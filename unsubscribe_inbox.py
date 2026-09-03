"""Read a Gmail mailbox over IMAP and decide who asked to be unsubscribed.

Ported from the FastAPI sibling project (Trax-Os/Clients/Raymonde,
apps/api/app/core/imap_poller.py), minus everything that needed Redis and a
job scheduler. Three layers, kept apart on purpose:

  1. PURE parsers on an email.message.Message. No network, no Notion, no
     streamlit. This is where every safety rule lives, so every safety rule is
     testable with a three-line fixture.
  2. IMAP I/O, guarded by a WALL-CLOCK deadline (see _Deadline). This is the
     property that makes the June 2026 grey-screen freeze impossible to repeat.
  3. scan_mailbox(), which returns a ScanReport and writes nothing anywhere.
     The caller decides what to push to Notion, because only the caller knows
     the suppression list.

Design notes worth keeping in mind while reading:

* No cursor, no state file. The Notion list itself is the idempotency store:
  the caller skips addresses it already has. Re-scanning the same window
  therefore costs one SEARCH plus a few hundred bytes per message and zero
  writes -- which is what lets this run on Streamlit Cloud, whose disk is wiped
  on every restart.
* Suppression is irreversible in this product (there is no un-suppress path).
  A false positive loses a prospect for good, so the classifier is
  conservative: when in doubt it does not suppress, it asks for review.
* Detection is bound to the SUBJECT. Free-text wording in a body only ever
  raises a REVIEW flag. A reply saying "stop, I'm interested, but end the other
  sequence" must not unsubscribe anyone by itself.
"""

from __future__ import annotations

import dataclasses
import email
import email.header
import email.utils
import imaplib
import re
import socket
import ssl
import time
from dataclasses import dataclass, field
from datetime import date, timedelta
from email.message import Message
from typing import Callable, FrozenSet, List, Optional, Sequence, Tuple

# --- constants --------------------------------------------------------------

# RFC 3501 SINCE wants English month abbreviations. strftime("%b") is
# locale-dependent -- under fr_FR.UTF-8 it yields "sep", and on glibc "sept.",
# both of which the server rejects. The upstream poller has this bug
# (imap_poller.py:481); we do not port it.
_IMAP_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

# What our own mailto produces ("unsubscribe {address}") plus the obvious human
# variants, FR/EN. Bound to the subject on purpose.
_UNSUB_SUBJECT_RE = re.compile(r"\bunsubscribe\b|d[ée]sinscri", re.I)

_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# Free-text opposition wording, searched in a reply BODY. Deliberately separate
# from _UNSUB_SUBJECT_RE: that one triggers an automatic action because an
# "unsubscribe" subject is unambiguous, this one only ever FLAGS. The law asks
# for the request to be handled, not for a regular expression to be trusted.
_OPPOSITION_BODY_RE = re.compile(
    r"d[ée]sinscri"
    r"|\bunsubscribe\b"
    r"|ne (?:plus|pas) (?:me |nous )?(?:re)?contacter"
    r"|ne me (?:re)?contactez plus"
    r"|retirez[- ](?:moi|nous)"
    r"|retirer mon adresse"
    r"|supprim(?:er|ez|e) (?:mon|mes) (?:adresse|donn[ée]es|coordonn[ée]es)"
    r"|opposition|RGPD|GDPR"
    r"|remove me|take me off|opt[- ]?out|do not (?:contact|email)",
    re.I,
)

# A 4.x.x DSN is NOT a bounce: the message is still in flight and the remote
# server will retry. Confusing it with a 5.x.x would PERMANENTLY suppress a
# perfectly valid prospect. Our most frequent case is Microsoft's 451 4.7.500,
# which fires precisely when a sender "changes its sending habits", i.e. at the
# start of a campaign.
_DSN_STATUS_RE = re.compile(r"^\s*Status:\s*([245])\.\d+\.\d+", re.I | re.M)
_DSN_ACTION_RE = re.compile(r"^\s*Action:\s*(\w+)", re.I | re.M)
_DSN_DIAG_RE = re.compile(r"^\s*Diagnostic-Code:\s*(.+)$", re.I | re.M)
_FINAL_RCPT_RE = re.compile(r"Final-Recipient:\s*[^;\n]*;\s*(\S+)", re.I)

# And within the 5.x.x class, only the ADDRESSING sub-class means "this address
# does not exist". RFC 3463 splits the detail digit by subject: 5.1.x is
# addressing (bad mailbox, bad destination system, null MX), while 5.4.x is
# routing and 5.7.x is security/policy. Microsoft overloads
# "550 5.4.1 Recipient address rejected: Access denied" for BOTH an unknown
# recipient AND a policy block aimed at the SENDER -- so a 5.4.1 may well be a
# perfectly valid prospect whose employer blocked us. Since suppression is
# irreversible in this product, only 5.1.x is auto-suppressed; every other
# permanent code, and every DSN with no status code at all, goes to review.
# (Upstream treated unknown as permanent, but upstream matched each DSN to a
# specific campaign recipient by Message-ID, which gave it corroboration this
# app does not have.)
_DSN_BAD_ADDRESS_RE = re.compile(r"^5\.1\.\d+$")

# Headers we ask for. Small and fixed, so a scan costs a few hundred bytes per
# message instead of the megabytes a full body would.
HEADER_FIELDS = (
    "SUBJECT FROM TO CC DATE MESSAGE-ID IN-REPLY-TO REFERENCES "
    "DELIVERED-TO X-ORIGINAL-TO X-FORWARDED-TO ENVELOPE-TO "
    "X-ORIGINAL-RECIPIENT RETURN-PATH CONTENT-TYPE "
    "AUTO-SUBMITTED PRECEDENCE X-AUTOREPLY X-AUTORESPOND"
)

_ADDRESS_HEADERS = ("Delivered-To", "X-Original-To", "X-Forwarded-To",
                    "Envelope-To", "X-Original-Recipient", "To", "Cc")

# A DSN embeds a copy of the original message, and ours carry inline images:
# 0.5 to 3 MB per bounce. The message/delivery-status part is almost always in
# the first few KB, before that copy, so a partial fetch is enough.
PARTIAL_BODY_BYTES = 32768

ACTION_SUPPRESS = "suppress"
ACTION_REVIEW = "review"
ACTION_IGNORE = "ignore"

ROLE_UNSUB = "unsub"      # the mailbox behind the desinscription@ alias
ROLE_BOUNCE = "bounce"    # a sender's own mailbox: hard bounces
ROLE_REPLY = "reply"      # a sender's own mailbox: "remove me" wording, flag only

REASON_UNSUB = "unsubscribe_request"
REASON_BOUNCE = "bounce_permanent"


# --- config and results -----------------------------------------------------

@dataclass(frozen=True)
class ImapConfig:
    """One mailbox, one role. Frozen so it can be logged and compared safely."""

    user: str = ""
    password: str = ""
    host: str = "imap.gmail.com"
    port: int = 993
    mailbox: str = "INBOX"
    role: str = ROLE_UNSUB
    # The address the message must have been delivered to. MANDATORY when the
    # unsubscribe address is an alias into a real person's mailbox: without it
    # the INBOX holds all of their mail and any subject containing
    # "desinscription" becomes a candidate.
    recipient_filter: str = ""
    since_days: int = 30
    max_messages: int = 60
    max_bodies: int = 5
    socket_timeout_s: int = 10
    deadline_s: float = 20.0
    own_domains: FrozenSet[str] = frozenset()
    never_suppress: FrozenSet[str] = frozenset()

    @property
    def is_configured(self) -> bool:
        return bool(self.user and self.password)


@dataclass
class Candidate:
    """One decision about one message."""

    action: str = ACTION_IGNORE
    email: str = ""
    reason: str = ""          # what goes into the Notion page body
    why: str = ""             # which barrier decided, for the journal
    display_name: str = ""    # From's display name -> Notion "Nom"
    subject: str = ""
    from_addr: str = ""
    date_hdr: str = ""
    uid: str = ""
    mailbox: str = ""
    dsn_status: str = ""      # extended code for a bounce
    review_hint: str = ""     # a secondary address a human should look at


@dataclass
class ScanReport:
    mailbox: str = ""
    role: str = ""
    searched: int = 0         # UIDs the server matched
    examined: int = 0         # messages actually fetched and classified
    ignored: int = 0
    truncated: bool = False   # SEARCH matched more than max_messages
    suppress: List[Candidate] = field(default_factory=list)
    review: List[Candidate] = field(default_factory=list)
    deferred: List[Candidate] = field(default_factory=list)  # 4.x.x, log only
    duration_s: float = 0.0
    error: str = ""
    error_kind: str = ""      # "" | auth | quota | network | mailbox | config
    added: int = 0            # filled by the caller after writing to Notion
    duplicates: int = 0       # filled by the caller

    @property
    def ok(self) -> bool:
        return not self.error_kind


class ImapScanError(Exception):
    """Raised only inside this module, never leaked out of scan_mailbox()."""

    def __init__(self, kind: str, message: str):
        super().__init__(message)
        self.kind = kind
        self.message = message


# --- layer 1: pure parsers --------------------------------------------------

def imap_since_token(day: date) -> str:
    """'03-Sep-2026'. Locale-independent by construction."""
    return "%02d-%s-%04d" % (day.day, _IMAP_MONTHS[day.month - 1], day.year)


def decode_rfc2047(value: Optional[str]) -> str:
    """'=?UTF-8?Q?d=C3=A9sinscription?=' -> 'désinscription'.

    Without this the subject regex silently misses every accented request --
    a false NEGATIVE, which for a legal obligation is the worse kind.
    """
    if not value:
        return ""
    out = []
    try:
        parts = email.header.decode_header(value)
    except Exception:
        return str(value)
    for data, charset in parts:
        if isinstance(data, bytes):
            try:
                out.append(data.decode(charset or "utf-8", "replace"))
            except (LookupError, UnicodeDecodeError):
                out.append(data.decode("utf-8", "replace"))
        else:
            out.append(data)
    return "".join(out)


def normalize_address(addr: str) -> str:
    """Lowercase and strip angle brackets. No +tag stripping: 'a+x@b.fr' is a
    distinct address as far as an unsubscribe request is concerned, and
    guessing otherwise could suppress someone who never asked."""
    if not addr:
        return ""
    return addr.strip().strip("<>").strip().lower()


def sender(msg: Message) -> Tuple[str, str]:
    """(display name, address) from From, both decoded and normalized."""
    raw = decode_rfc2047(msg.get("From") or "")
    name, addr = email.utils.parseaddr(raw)
    return name.strip(), normalize_address(addr)


def subject_addresses(subject: str) -> List[str]:
    """Every address in a subject, lowercased and deduplicated, order kept.

    More than one means we must not guess: 'Fwd: unsubscribe -- see the mail
    from jean@other.fr' could designate the wrong person.
    """
    seen: List[str] = []
    seen_set = set()
    for hit in _EMAIL_RE.findall(subject or ""):
        low = hit.strip().lower()
        if low and low not in seen_set:
            seen_set.add(low)
            seen.append(low)
    return seen


def recipient_addresses(msg: Message) -> FrozenSet[str]:
    """Every address this message was delivered or addressed to.

    Gmail fills Delivered-To with the envelope recipient, alias included, which
    is what makes the alias case tractable.
    """
    found = set()
    for header in _ADDRESS_HEADERS:
        for raw in msg.get_all(header) or []:
            for _n, addr in email.utils.getaddresses([decode_rfc2047(raw)]):
                low = normalize_address(addr)
                if low:
                    found.add(low)
    return frozenset(found)


def targets_mailbox(msg: Message, wanted: str) -> bool:
    """Barrier 0: was this message actually addressed to us?"""
    if not wanted:
        return True
    return normalize_address(wanted) in recipient_addresses(msg)


def is_auto_reply(msg: Message) -> bool:
    """Vacation / out-of-office / auto-acknowledgement.

    Without this an 'Réponse automatique : Désinscription' sent by someone who
    asked for nothing gets them unsubscribed.
    """
    auto = (msg.get("Auto-Submitted") or "").strip().lower()
    if auto and auto != "no":
        return True
    if msg.get("X-Autoreply") or msg.get("X-Autorespond"):
        return True
    precedence = (msg.get("Precedence") or "").strip().lower()
    if precedence in ("auto_reply", "bulk", "junk", "list"):
        return True
    subject = decode_rfc2047(msg.get("Subject")).lower()
    return any(k in subject for k in (
        "out of office", "absence du bureau", "réponse automatique",
        "reponse automatique", "automatic reply", "autoreply",
    ))


def looks_like_bounce(msg: Message) -> bool:
    """Barrier 1, on HEADERS ONLY: is this a delivery report?

    Must run BEFORE the subject regex. Our own List-Unsubscribe header puts the
    prospect's address IN THE SUBJECT, so a DSN about that mail would otherwise
    surface a real prospect, and mailer-daemon@googlemail.com would land in the
    suppression list. Upstream got this for free because parse_bounce() returned
    early before parse_unsubscribe_request(); porting the second function alone
    would silently drop that guarantee, so the filter lives in the classifier.
    """
    ctype = (msg.get("Content-Type") or "").lower()
    # A spam complaint is ALSO a multipart/report, so a bare startswith() would
    # swallow it here and it would never reach the review queue. Report-type is
    # what distinguishes them, so honour it before the generic test.
    if "report-type=feedback-report" in ctype:
        return False
    if "report-type=delivery-status" in ctype:
        return True
    if ctype.startswith("multipart/report"):
        return True
    for header in ("From", "Return-Path", "Sender"):
        raw = (msg.get(header) or "").lower()
        if "mailer-daemon" in raw or "postmaster" in raw:
            return True
    return False


def looks_like_arf(msg: Message) -> bool:
    """A spam complaint (Abuse Reporting Format).

    Some ARFs quote the original subject, so 'unsubscribe {address}' appears in
    them. An ARF IS a legitimate suppression signal, but the address to suppress
    sits in the report body, not in From: parsing that properly is out of scope,
    and guessing is worse than asking a human.
    """
    ctype = (msg.get("Content-Type") or "").lower()
    if "report-type=feedback-report" in ctype:
        return True
    # The sender heuristic only applies to an actual report. On its own it would
    # send a genuine request from, say, staff@somecompany.fr to review instead of
    # honouring it -- safe, but needlessly manual.
    if not ctype.startswith("multipart/report"):
        return False
    from_raw = (msg.get("From") or "").lower()
    return any(k in from_raw for k in ("staff@", "abuse@", "fbl@", "feedback-loop"))


def looks_like_bulk(msg: Message) -> bool:
    """Mailing-list or bulk mail: never a personal message about us.

    Needed because EVERY newsletter footer contains the word "unsubscribe", so
    without this the opposition-signal queue fills with Substack digests instead
    of the prospect replies a human actually needs to see. RFC 2369 list headers
    are the reliable discriminator: a genuine reply from a prospect carries none
    of them. Note is_auto_reply() catches "Precedence: bulk" but plenty of
    senders (Substack among them) set no Precedence at all.
    """
    for header in ("List-Unsubscribe", "List-Id", "List-Post", "List-Help",
                   "List-Subscribe", "X-Mailer-LID", "X-Campaign-Id",
                   "Feedback-ID"):
        if msg.get(header):
            return True
    return False


def is_reply(msg: Message) -> bool:
    return bool(msg.get("In-Reply-To") or msg.get("References"))


def indice_opposition(text: Optional[str]) -> Optional[str]:
    """The opposition wording found, or None. Used ONLY to raise a review flag:
    nothing is ever decided from this result."""
    if not text:
        return None
    hit = _OPPOSITION_BODY_RE.search(text)
    return hit.group(0).lower() if hit else None


def body_text(msg: Message, limit: int = 20000) -> str:
    """First text/plain part, decoded. Tolerates a truncated message (a partial
    fetch parses its leading parts correctly)."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if (part.get_content_type() or "") == "text/plain":
                    raw = part.get_payload(decode=True)
                    if raw:
                        return raw.decode("utf-8", "replace")[:limit]
            return ""
        raw = msg.get_payload(decode=True)
        if isinstance(raw, bytes):
            return raw.decode("utf-8", "replace")[:limit]
        return str(msg.get_payload())[:limit]
    except Exception:
        return ""


def parse_dsn_verdict(msg: Message) -> Tuple[str, str, str]:
    """('permanent'|'temporaire'|'inconnu', extended code, diagnostic text).

    The code overrides the action: an 'Action: failed' carrying a 4.x.x status
    stays temporary. With neither present we return 'inconnu', and the caller
    treats unknown as permanent, preserving upstream behaviour on malformed
    DSNs.
    """
    raw = ""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if (part.get_content_type() or "").lower() == "message/delivery-status":
                    raw = part.as_string()
                    break
        if not raw:
            raw = msg.as_string()
    except Exception:
        raw = ""

    diag = ""
    hit = _DSN_DIAG_RE.search(raw)
    if hit:
        diag = hit.group(1).strip()[:300]

    hit = _DSN_STATUS_RE.search(raw)
    if hit:
        code = hit.group(0).split(":", 1)[1].strip()
        klass = hit.group(1)
        if klass == "4":
            return "temporaire", code, diag
        if klass == "5":
            return "permanent", code, diag
        return "inconnu", code, diag   # 2.x.x is a delivery receipt, not a failure

    hit = _DSN_ACTION_RE.search(raw)
    if hit and hit.group(1).lower() == "delayed":
        return "temporaire", "", diag
    return "inconnu", "", diag


def parse_bounce_recipient(msg: Message) -> str:
    """The address that failed, from Final-Recipient."""
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if (part.get_content_type() or "").lower() == "message/delivery-status":
                    payload = part.get_payload()
                    for sub in payload if isinstance(payload, list) else []:
                        if isinstance(sub, Message):
                            value = sub.get("Final-Recipient")
                            if value:
                                return normalize_address(value.split(";", 1)[-1])
                    hit = _FINAL_RCPT_RE.search(part.as_string())
                    if hit:
                        return normalize_address(hit.group(1))
        hit = _FINAL_RCPT_RE.search(msg.as_string())
        if hit:
            return normalize_address(hit.group(1))
    except Exception:
        pass
    return ""


def _is_internal(addr: str, own_domains: FrozenSet[str]) -> bool:
    if not addr or "@" not in addr:
        return False
    return addr.rsplit("@", 1)[-1].lower() in {d.lower() for d in own_domains}


def _protected(addr: str, cfg: ImapConfig) -> bool:
    """Barrier 6: never, under any circumstance, suppress one of our own.

    Verified: the subject 'unsubscribe desinscription@merciraymond.fr' -- exactly
    what an auto-responder on that mailbox would produce -- extracts our own
    address.
    """
    if not addr:
        return True
    if addr in {a.lower() for a in cfg.never_suppress}:
        return True
    if cfg.recipient_filter and addr == normalize_address(cfg.recipient_filter):
        return True
    return _is_internal(addr, cfg.own_domains)


def _base(msg: Message, uid: str, cfg: ImapConfig) -> Candidate:
    name, addr = sender(msg)
    return Candidate(
        uid=str(uid),
        mailbox=cfg.user,
        subject=decode_rfc2047(msg.get("Subject"))[:300],
        from_addr=addr,
        display_name=name,
        date_hdr=(msg.get("Date") or "").strip()[:120],
    )


def classify_unsubscribe(msg: Message, uid: str, cfg: ImapConfig) -> Candidate:
    """The eight barriers, in decreasing order of safety.

    Returns a Candidate whose .action is suppress / review / ignore, and whose
    .why names the barrier that decided -- so the journal explains itself.
    """
    out = _base(msg, uid, cfg)
    out.reason = REASON_UNSUB

    # 0. Was it addressed to us? Mandatory with an alias.
    if not targets_mailbox(msg, cfg.recipient_filter):
        out.action, out.why = ACTION_IGNORE, "pas adresse a la boite de desinscription"
        return out

    # 1. Delivery reports never take this path.
    if looks_like_bounce(msg):
        out.action, out.why = ACTION_IGNORE, "avis de non-remise (traite comme rebond)"
        return out

    # 2. Spam complaints: a real signal, but the address is in the body.
    if looks_like_arf(msg):
        out.action, out.why = ACTION_REVIEW, "rapport de spam (adresse dans le corps)"
        return out

    # 3. Auto-responders, and bulk mail. A footer click or Gmail's Unsubscribe
    #    button produces ordinary personal mail, never a list mailing, so a
    #    message carrying RFC 2369 list headers is not a request from a human.
    if is_auto_reply(msg):
        out.action, out.why = ACTION_IGNORE, "reponse automatique"
        return out
    if looks_like_bulk(msg):
        out.action, out.why = ACTION_IGNORE, "courrier de masse (en-tetes de liste)"
        return out

    # 4. The subject must say so. Decoded first, or accented requests are lost.
    if not _UNSUB_SUBJECT_RE.search(out.subject):
        out.action, out.why = ACTION_IGNORE, "sujet sans demande de desinscription"
        return out

    addrs = subject_addresses(out.subject)
    from_addr = out.from_addr

    # 7. Several addresses in the subject: we never pick between two people.
    if len(addrs) > 1:
        out.action = ACTION_REVIEW
        out.why = "plusieurs adresses dans le sujet"
        out.review_hint = ", ".join(addrs[:5])
        return out

    subject_addr = addrs[0] if addrs else ""
    from_internal = _is_internal(from_addr, cfg.own_domains)
    subject_internal = _is_internal(subject_addr, cfg.own_domains)

    # 5. Internal / external arbitration. "The subject wins" is only correct
    #    when the sender is a colleague relaying a request; when the sender is
    #    external, the sender is the one who acted and the subject address is
    #    just a string that happened to be there.
    if from_internal:
        if subject_addr and not subject_internal:
            out.action, out.email = ACTION_SUPPRESS, subject_addr
            out.why = "transfert interne, adresse du sujet"
            out.display_name = ""   # the colleague's name is not the prospect's
        else:
            out.action = ACTION_REVIEW
            out.why = "expediteur interne sans adresse externe (test ou note interne)"
            out.review_hint = subject_addr or from_addr
            return out
    else:
        if not subject_addr or subject_addr == from_addr:
            out.action, out.email = ACTION_SUPPRESS, from_addr
            out.why = "expediteur externe (clic pied de page ou bouton Gmail)"
        elif subject_internal:
            out.action = ACTION_REVIEW
            out.why = "le sujet designe une de nos propres adresses"
            out.review_hint = subject_addr
            return out
        else:
            out.action, out.email = ACTION_SUPPRESS, from_addr
            out.why = "expediteur externe, adresse du sujet differente"
            out.review_hint = subject_addr   # a human should look at that one too

    # 6. Last gate before any write.
    if _protected(out.email, cfg):
        out.review_hint = out.email
        out.action, out.email = ACTION_REVIEW, ""
        out.why = "adresse protegee (domaine interne ou liste never_suppress)"
    return out


def classify_bounce(msg: Message, uid: str, cfg: ImapConfig) -> Candidate:
    """A delivery report from a sender's own mailbox.

    Only a 5.x.x suppresses. A 4.x.x is recorded and nothing else: the message
    is still in flight, and a rising deferral rate is the earliest signal that
    a sender's reputation is degrading.
    """
    out = _base(msg, uid, cfg)
    if not looks_like_bounce(msg):
        out.action, out.why = ACTION_IGNORE, "pas un avis de non-remise"
        return out

    verdict, code, diag = parse_dsn_verdict(msg)
    out.dsn_status = code
    failed = parse_bounce_recipient(msg)
    out.display_name = ""       # "MAILER-DAEMON" is not a person's name

    if not failed:
        out.action, out.why = ACTION_IGNORE, "avis de non-remise sans Final-Recipient"
        return out

    if verdict == "temporaire":
        out.action = ACTION_IGNORE
        out.email = failed
        out.why = "DSN temporaire %s -- aucun effet" % (code or "4.x.x")
        out.reason = "deferred"
        return out

    detail = (" -- " + diag[:80]) if diag else ""

    if verdict == "inconnu" or not code:
        # No usable status code. Suppressing on a guess would be irreversible.
        out.action = ACTION_REVIEW
        out.email = ""
        out.review_hint = failed
        out.reason = "bounce_indetermine"
        out.why = "avis de non-remise sans code exploitable%s" % detail
        return out

    if not _DSN_BAD_ADDRESS_RE.match(code):
        # Permanent, but about routing or policy rather than the address.
        out.action = ACTION_REVIEW
        out.email = ""
        out.review_hint = failed
        out.reason = "bounce_politique"
        out.why = ("rebond permanent %s (routage ou politique, pas une adresse "
                   "inexistante)%s" % (code, detail))
        return out

    out.email = failed
    out.reason = REASON_BOUNCE
    out.action = ACTION_SUPPRESS
    out.why = "adresse inexistante %s%s" % (code, detail)
    if _protected(out.email, cfg):
        out.review_hint = out.email
        out.action, out.email = ACTION_REVIEW, ""
        out.why = "rebond sur une adresse protegee"
    return out


def classify_reply(msg: Message, uid: str, cfg: ImapConfig) -> Candidate:
    """Free-text opposition wording in a reply body. NEVER auto-suppresses.

    The upstream comment explains why and it is right: a reply saying "stop,
    I'm interested, but end the other sequence" must not unsubscribe anyone on
    its own. The law asks for the request to be handled, not for a regex to be
    trusted. This only puts the message in the review queue.
    """
    out = _base(msg, uid, cfg)
    out.reason = "opposition_reponse"
    if looks_like_bounce(msg) or is_auto_reply(msg) or looks_like_arf(msg):
        out.action, out.why = ACTION_IGNORE, "avis machine, pas une reponse humaine"
        return out
    if looks_like_bulk(msg):
        out.action, out.why = ACTION_IGNORE, "courrier de masse (en-tetes de liste)"
        return out
    if _is_internal(out.from_addr, cfg.own_domains):
        out.action, out.why = ACTION_IGNORE, "expediteur interne"
        return out
    hit = indice_opposition(body_text(msg))
    if not hit:
        out.action, out.why = ACTION_IGNORE, "aucune formule d'opposition"
        return out
    out.action = ACTION_REVIEW
    out.review_hint = out.from_addr
    out.why = "formule d'opposition dans le corps: %r" % hit
    return out


# --- layer 2: IMAP I/O ------------------------------------------------------

class _Deadline:
    """A wall-clock budget for one whole cycle.

    Why this and not just the constructor's timeout: imaplib passes `timeout`
    to socket.create_connection AND makes it the timeout of every subsequent
    read. A hanging server therefore costs `timeout` PER COMMAND -- upstream's
    timeout=30 over eight commands is 240 s. That was survivable there only
    because rq-scheduler imposed a job timeout on top (hence the re-raised
    JobTimeoutException at imap_poller.py:527). We have no such umbrella, so we
    carry our own and re-arm the socket before every command.
    """

    __slots__ = ("end",)

    def __init__(self, budget_s: float):
        self.end = time.monotonic() + float(budget_s)

    def left(self) -> float:
        return self.end - time.monotonic()

    def blown(self) -> bool:
        return self.left() <= 0.5


def _arm(conn, deadline: _Deadline, per_read_s: int) -> None:
    """Make the deadline hard, even inside a blocking read."""
    try:
        conn.sock.settimeout(max(1.0, min(float(per_read_s), deadline.left())))
    except Exception:
        pass   # semi-private attribute: fall back to the constructor's timeout


def classify_error(exc: BaseException) -> str:
    """Map a failure to a backoff class."""
    text = str(exc).upper()
    if "AUTHENTICATIONFAILED" in text or "INVALID CREDENTIALS" in text:
        return "auth"
    if "APPLICATION-SPECIFIC PASSWORD" in text or "WEBLOGIN" in text:
        return "auth"
    if "OVERQUOTA" in text or "LIMIT" in text or "TOO MANY" in text:
        return "quota"
    if isinstance(exc, (socket.timeout, socket.gaierror, ssl.SSLError, OSError)):
        return "network"
    if "NONEXISTENT" in text or "CANNOT OPEN" in text:
        return "mailbox"
    return "unknown"


def _connect(cfg: ImapConfig, deadline: _Deadline):
    try:
        try:
            conn = imaplib.IMAP4_SSL(cfg.host, cfg.port,
                                     timeout=cfg.socket_timeout_s)
        except TypeError:
            # The `timeout` keyword landed in Python 3.9; the machine's default
            # python3 is 3.7.9 and the test scripts run on the default.
            conn = imaplib.IMAP4_SSL(cfg.host, cfg.port)
            try:
                conn.sock.settimeout(cfg.socket_timeout_s)
            except Exception:
                pass
    except Exception as exc:
        raise ImapScanError(classify_error(exc), "connexion: %s" % exc)

    try:
        _arm(conn, deadline, cfg.socket_timeout_s)
        conn.login(cfg.user, cfg.password)
        _arm(conn, deadline, cfg.socket_timeout_s)
        # readonly: never mark a human's mail as read. BODY.PEEK everywhere is
        # the belt; this is the braces.
        typ, _ = conn.select(cfg.mailbox, readonly=True)
        if typ != "OK":
            raise ImapScanError("mailbox", "impossible d'ouvrir %s" % cfg.mailbox)
    except ImapScanError:
        _quiet_logout(conn)
        raise
    except Exception as exc:
        _quiet_logout(conn)
        raise ImapScanError(classify_error(exc), "ouverture: %s" % exc)
    return conn


def _quiet_logout(conn) -> None:
    for method in ("close", "logout"):
        try:
            getattr(conn, method)()
        except Exception:
            pass


def search_criteria(cfg: ImapConfig, today: date) -> str:
    """The exact SEARCH string, per role. Server-side filtering is what keeps a
    scan cheap and keeps us out of mail that is none of our business."""
    since = imap_since_token(today - timedelta(days=max(1, cfg.since_days)))
    if cfg.role == ROLE_UNSUB:
        if cfg.recipient_filter:
            # The value comes from a secret, so this is belt-and-braces rather
            # than a real injection worry -- but an unescaped quote here would
            # produce a malformed IMAP command and a confusing "search failed".
            if '"' in cfg.recipient_filter or "\\" in cfg.recipient_filter:
                raise ImapScanError(
                    "config",
                    "UNSUB_IMAP_FILTER contient un caractere interdit (guillemet "
                    "ou antislash)")
            # Highly selective, and the only thing that makes reading a shared
            # personal mailbox acceptable.
            return 'SINCE %s HEADER Delivered-To "%s"' % (since, cfg.recipient_filter)
        # A dedicated unsubscribe mailbox: everything in it concerns us, and
        # its traffic is low. No subject filter, because a MIME-encoded subject
        # would not reliably match server-side.
        return "SINCE %s" % since
    if cfg.role == ROLE_BOUNCE:
        return 'SINCE %s OR FROM "mailer-daemon" FROM "postmaster"' % since
    if cfg.role == ROLE_REPLY:
        # ASCII-only terms on purpose: server-side tokenization of accents is
        # not something to bet on. This is only a candidate filter -- the
        # Python regex, which does handle accents, is the actual judge.
        return ('SINCE %s OR BODY "unsubscribe" OR BODY "desinscri" '
                'OR BODY "opposition" BODY "retirez"' % since)
    raise ImapScanError("config", "role inconnu: %r" % cfg.role)


def _fetch_items(conn, uids: Sequence[bytes], spec: str,
                 deadline: _Deadline, cfg: ImapConfig) -> List[Tuple[str, bytes]]:
    """Fetch `spec` for a batch of UIDs. Batched: one command for up to 50 UIDs
    is far cheaper than fifty round trips."""
    out: List[Tuple[str, bytes]] = []
    batch_size = 50
    for start in range(0, len(uids), batch_size):
        if deadline.blown():
            break
        batch = uids[start:start + batch_size]
        joined = b",".join(batch).decode("ascii")
        _arm(conn, deadline, cfg.socket_timeout_s)
        try:
            typ, data = conn.uid("fetch", joined, spec)
        except Exception as exc:
            raise ImapScanError(classify_error(exc), "fetch: %s" % exc)
        if typ != "OK" or not data:
            continue
        # imaplib returns a flat list alternating (b'<seq> (UID n ...', payload)
        # tuples and closing b')' strings.
        for item in data:
            if not isinstance(item, tuple) or len(item) < 2:
                continue
            prefix = item[0] if isinstance(item[0], bytes) else b""
            hit = re.search(rb"UID (\d+)", prefix)
            uid = hit.group(1).decode("ascii") if hit else ""
            payload = item[1]
            if isinstance(payload, bytes) and payload:
                out.append((uid, payload))
    return out


def fetch_candidates(cfg: ImapConfig, today: Optional[date] = None,
                     deadline: Optional[_Deadline] = None
                     ) -> Tuple[List[Tuple[str, Message]], int]:
    """([(uid, Message)], number of UIDs the server matched).

    Headers only for the unsubscribe role. For bounces and replies, headers
    first to screen, then a PARTIAL body fetch (32 KB) for at most max_bodies
    of them -- which is what makes reading delivery reports affordable.
    """
    today = today or date.today()
    deadline = deadline or _Deadline(cfg.deadline_s)
    if not cfg.is_configured:
        raise ImapScanError("config", "identifiants IMAP absents")

    conn = _connect(cfg, deadline)
    try:
        criteria = search_criteria(cfg, today)
        _arm(conn, deadline, cfg.socket_timeout_s)
        try:
            typ, data = conn.uid("search", None, criteria)
        except Exception as exc:
            raise ImapScanError(classify_error(exc), "search: %s" % exc)
        if typ != "OK":
            raise ImapScanError("unknown", "search a repondu %s" % typ)
        all_uids = data[0].split() if data and data[0] else []
        total = len(all_uids)
        # The NEWEST ones. Upstream slices uids[:max] because a Redis cursor
        # moves forward for it; with no cursor that would mean never reaching
        # recent mail once a window exceeds the cap.
        uids = all_uids[-cfg.max_messages:]

        header_spec = "(BODY.PEEK[HEADER.FIELDS (%s)])" % HEADER_FIELDS
        raw_headers = _fetch_items(conn, uids, header_spec, deadline, cfg)
        parsed = [(uid, email.message_from_bytes(raw)) for uid, raw in raw_headers]

        if cfg.role == ROLE_UNSUB:
            return parsed, total

        # Screen on headers, then pull partial bodies for the few that matter.
        if cfg.role == ROLE_BOUNCE:
            wanted = [(uid, msg) for uid, msg in parsed if looks_like_bounce(msg)]
        else:
            wanted = [(uid, msg) for uid, msg in parsed
                      if is_reply(msg) and not looks_like_bounce(msg)
                      and not is_auto_reply(msg)]
        wanted = wanted[-cfg.max_bodies:]
        if not wanted:
            return [], total

        body_spec = "(BODY.PEEK[]<0.%d>)" % PARTIAL_BODY_BYTES
        raw_bodies = _fetch_items(conn, [u.encode("ascii") for u, _ in wanted],
                                  body_spec, deadline, cfg)
        by_uid = {uid: raw for uid, raw in raw_bodies}
        full: List[Tuple[str, Message]] = []
        for uid, header_msg in wanted:
            raw = by_uid.get(uid)
            # A partial fetch is a truncated message; message_from_bytes still
            # parses its leading parts, which is where delivery-status sits.
            full.append((uid, email.message_from_bytes(raw) if raw else header_msg))
        return full, total
    finally:
        _quiet_logout(conn)


# --- layer 3: orchestration -------------------------------------------------

_CLASSIFIERS = {
    ROLE_UNSUB: classify_unsubscribe,
    ROLE_BOUNCE: classify_bounce,
    ROLE_REPLY: classify_reply,
}


def scan_mailbox(cfg: ImapConfig, today: Optional[date] = None,
                 fetcher: Optional[Callable] = None) -> ScanReport:
    """One mailbox, one role. Writes nothing, and never raises.

    `fetcher` is injectable so the whole decision path can be tested without a
    server. Returning instead of raising is deliberate: a scan is a comfort,
    and a broken mailbox must never be able to take the app down with it.
    """
    started = time.monotonic()
    report = ScanReport(mailbox=cfg.user, role=cfg.role)
    if not cfg.is_configured:
        report.error_kind, report.error = "config", "identifiants IMAP absents"
        return report

    classify = _CLASSIFIERS.get(cfg.role)
    if classify is None:
        report.error_kind, report.error = "config", "role inconnu: %r" % cfg.role
        return report

    try:
        items, total = (fetcher or fetch_candidates)(cfg, today)
    except ImapScanError as exc:
        report.error_kind, report.error = exc.kind, exc.message
        report.duration_s = time.monotonic() - started
        return report
    except (KeyboardInterrupt, SystemExit):
        raise
    except BaseException as exc:                      # noqa: BLE001
        report.error_kind = classify_error(exc)
        report.error = "%s: %s" % (exc.__class__.__name__, exc)
        report.duration_s = time.monotonic() - started
        return report

    report.searched = total
    report.truncated = total > len(items) and total > cfg.max_messages
    seen_suppress = set()
    for uid, msg in items:
        report.examined += 1
        try:
            verdict = classify(msg, uid, cfg)
        except (KeyboardInterrupt, SystemExit):
            raise
        except BaseException as exc:                  # noqa: BLE001
            report.ignored += 1
            report.review.append(Candidate(
                action=ACTION_REVIEW, uid=str(uid), mailbox=cfg.user,
                why="classification impossible: %s" % exc))
            continue
        if verdict.action == ACTION_SUPPRESS and verdict.email:
            if verdict.email in seen_suppress:
                continue                              # dedup within one scan
            seen_suppress.add(verdict.email)
            report.suppress.append(verdict)
            if verdict.review_hint:
                report.review.append(verdict)
        elif verdict.action == ACTION_REVIEW:
            report.review.append(verdict)
        else:
            report.ignored += 1
            if verdict.reason == "deferred" and verdict.email:
                report.deferred.append(verdict)
    report.duration_s = time.monotonic() - started
    return report


def with_deadline(cfg: ImapConfig, seconds: float) -> ImapConfig:
    """Same config with a shorter wall-clock budget.

    The caller scans several mailboxes in one cycle and holds a budget for the
    whole cycle, so each scan must be handed what is actually left rather than
    its own full allowance.
    """
    return dataclasses.replace(
        cfg, deadline_s=max(1.0, min(cfg.deadline_s, float(seconds))))


def provenance_lines(cand: Candidate) -> List[str]:
    """The audit trail written into the Notion page body.

    With no 'Raison' column this is the only place the reason can live, and it
    stays the way to tell a machine-created row from a hand-typed one.
    """
    lines = ["Ajoute automatiquement par le Raymongraphe."]
    if cand.reason:
        lines.append("Raison : %s" % cand.reason)
    if cand.dsn_status:
        lines.append("Code SMTP : %s" % cand.dsn_status)
    if cand.date_hdr:
        lines.append("Recu le : %s" % cand.date_hdr)
    if cand.from_addr:
        lines.append("Expediteur : %s%s" % (
            (cand.display_name + " ") if cand.display_name else "", cand.from_addr))
    if cand.subject:
        lines.append("Sujet : %s" % cand.subject)
    if cand.mailbox:
        lines.append("Boite relevee : %s" % cand.mailbox)
    if cand.why:
        lines.append("Detection : %s" % cand.why)
    return lines
