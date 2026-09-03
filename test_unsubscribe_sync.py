#!/usr/bin/env python3
"""
Tests for the unsubscribe sync: Notion payload building and mailbox parsing.
Run from project root: python3.12 test_unsubscribe_sync.py
(3.12 is the interpreter that has the app's deps; the pure-module tests run
on any 3.7+, and the app-level ones skip themselves if streamlit is absent.)

No test opens a socket and no test touches Notion. Everything below is either a
pure function or scan_mailbox() with an injected fetcher -- which is the whole
point of keeping notion_props.py and unsubscribe_inbox.py free of I/O.

Two of these tests exist because of real bugs:
  * test_notion_payload_real_schema locks the "Raison" regression: the app used
    to POST a property that does not exist, so every automatic write returned
    HTTP 400 and the suppression list could only be filled by hand.
  * test_since_token_ignores_locale locks the upstream strftime("%d-%b-%Y") bug,
    which yields "03-sep-2026" under a French locale and is rejected by IMAP.
"""

import locale
import os
import pathlib
import socket
import sys
from datetime import date
from email import message_from_string

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import notion_props as np
import unsubscribe_inbox as ui

OWN = frozenset({"merciraymond.fr"})
UNSUB_BOX = "desinscription@merciraymond.fr"

CFG = ui.ImapConfig(
    user="prenom.nom@merciraymond.fr", password="app-password",
    role=ui.ROLE_UNSUB, recipient_filter=UNSUB_BOX, own_domains=OWN,
)
# Same thing without the alias filter, for tests about a dedicated mailbox.
CFG_NOFILTER = ui.ImapConfig(user="a@b.fr", password="p", role=ui.ROLE_UNSUB,
                             own_domains=OWN)
CFG_BOUNCE = ui.ImapConfig(user="salome@merciraymond.fr", password="p",
                           role=ui.ROLE_BOUNCE, since_days=7, own_domains=OWN)
CFG_REPLY = ui.ImapConfig(user="salome@merciraymond.fr", password="p",
                          role=ui.ROLE_REPLY, since_days=7, own_domains=OWN)


def msg(subject="", frm="", extra="", body="", to=UNSUB_BOX):
    """Build a Message from headers. Delivered-To defaults to the unsub box so
    barrier 0 passes unless a test says otherwise."""
    raw = ""
    if subject:
        raw += "Subject: %s\n" % subject
    if frm:
        raw += "From: %s\n" % frm
    if to is not None:
        raw += "Delivered-To: %s\n" % to
    raw += "Date: Wed, 3 Sep 2026 09:12:44 +0200\n"
    raw += extra
    raw += "\n" + body
    return message_from_string(raw)


# --- notion_props -----------------------------------------------------------

REAL_SCHEMA = {"Email": "title", "Date": "date", "Nom": "rich_text"}


def test_notion_payload_real_schema():
    """The exact schema of the live database: no 'Raison' must be emitted."""
    props = np.build_suppression_properties(
        REAL_SCHEMA, "jean@client.fr", reason="unsubscribe_request",
        day="2026-09-03", nom="Jean Dupont")
    assert set(props) == {"Email", "Date", "Nom"}, props
    assert "Raison" not in props
    assert props["Email"]["title"][0]["text"]["content"] == "jean@client.fr"
    assert props["Date"]["date"]["start"] == "2026-09-03"
    # Nom carries the human name, never the reason: it is the only column the
    # team reads, and all 15 hand-typed rows have a real name in it.
    assert props["Nom"]["rich_text"][0]["text"]["content"] == "Jean Dupont"
    print("  payload on the real schema (no Raison): OK")
    return True


def test_notion_payload_invariant():
    """set(properties) <= set(schema), whatever we are asked to write."""
    for schema in (REAL_SCHEMA,
                   {"Adresse": "title"},
                   {"Email": "title", "Raison": "select", "Nom": "rich_text"},
                   {"Email": "title", "Quand": "date", "Autre": "number"}):
        props = np.build_suppression_properties(
            schema, "a@b.fr", reason="bounce_permanent", day="2026-09-03",
            nom="A B")
        assert set(props) <= set(schema), (schema, props)
    print("  invariant properties <= schema: OK")
    return True


def test_notion_payload_reason_when_column_exists():
    """The day someone adds Raison, it fills itself -- whatever its type."""
    for ptype, reader in (
        ("rich_text", lambda p: p["Raison"]["rich_text"][0]["text"]["content"]),
        ("select", lambda p: p["Raison"]["select"]["name"]),
        ("multi_select", lambda p: p["Raison"]["multi_select"][0]["name"]),
    ):
        props = np.build_suppression_properties(
            {**REAL_SCHEMA, "Raison": ptype}, "a@b.fr", reason="bounce_permanent")
        assert reader(props) == "bounce_permanent", (ptype, props)
    # 'status' options cannot be created through the API, so writing one is a
    # guaranteed 400: we must skip it rather than try.
    props = np.build_suppression_properties(
        {**REAL_SCHEMA, "Raison": "status"}, "a@b.fr", reason="x")
    assert "Raison" not in props
    print("  Raison filled per real type, status skipped: OK")
    return True


def test_notion_title_found_by_type():
    """The title property is found by TYPE, so a renamed column still works."""
    props = np.build_suppression_properties({"Courriel": "title"}, "a@b.fr")
    assert set(props) == {"Courriel"}
    assert np.title_property({"X": "rich_text", "Y": "title"}) == "Y"
    assert np.title_property({"X": "rich_text"}) is None
    print("  title found by type: OK")
    return True


def test_notion_no_title_raises():
    """No title means nowhere to put the address: fail loudly, not partially."""
    try:
        np.build_suppression_properties({"Date": "date"}, "a@b.fr")
    except np.NotionSchemaError:
        print("  missing title raises NotionSchemaError: OK")
        return True
    raise AssertionError("NotionSchemaError attendue")


def test_notion_ambiguous_date_is_skipped():
    """Two unnamed date columns: writing the wrong date is worse than none."""
    props = np.build_suppression_properties(
        {"Email": "title", "Debut": "date", "Fin": "date"}, "a@b.fr",
        day="2026-09-03")
    assert set(props) == {"Email"}, props
    # A single date column, whatever its name, is unambiguous.
    props = np.build_suppression_properties(
        {"Email": "title", "Quand": "date"}, "a@b.fr", day="2026-09-03")
    assert props["Quand"]["date"]["start"] == "2026-09-03"
    print("  ambiguous date skipped, single date used: OK")
    return True


def test_notion_clip_and_schema_readers():
    assert len(np.clip("x" * 5000)) == np.NOTION_TEXT_LIMIT
    assert np.clip("") == ""
    page = {"properties": {"Email": {"type": "title", "title": []},
                           "Date": {"type": "date", "date": None}}}
    assert np.schema_from_page(page) == {"Email": "title", "Date": "date"}
    ds = {"properties": {"Nom": {"type": "rich_text"}}}
    assert np.schema_from_data_source(ds) == {"Nom": "rich_text"}
    assert np.schema_from_page({}) == {}
    children = np.build_provenance_children(["Raison : x", "", "Sujet : y"])
    text = children[0]["paragraph"]["rich_text"][0]["text"]["content"]
    assert text == "Raison : x\nSujet : y"
    assert np.build_provenance_children([]) == []
    print("  clip, schema readers, page body: OK")
    return True


# --- pure parsers -----------------------------------------------------------

def test_since_token_ignores_locale():
    """RFC 3501 wants 'Sep'. strftime('%b') under fr_FR gives 'sep'/'sept.'."""
    assert ui.imap_since_token(date(2026, 9, 3)) == "03-Sep-2026"
    saved = locale.setlocale(locale.LC_TIME)
    try:
        locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")
    except locale.Error:
        print("  SINCE token (fr_FR locale unavailable, ASCII case only): OK")
        return True
    try:
        assert ui.imap_since_token(date(2026, 9, 3)) == "03-Sep-2026"
        assert date(2026, 9, 3).strftime("%b") != "Sep", (
            "la locale fr n'a pas pris, le test ne prouve rien")
    finally:
        locale.setlocale(locale.LC_TIME, saved)
    print("  SINCE token immune to locale: OK")
    return True


def test_decode_rfc2047():
    """An accented subject arrives MIME-encoded; missing it loses the request."""
    encoded = "=?UTF-8?Q?D=C3=A9sinscription?="
    assert ui.decode_rfc2047(encoded) == "Désinscription"
    assert ui.decode_rfc2047("") == ""
    assert ui.decode_rfc2047("plain") == "plain"
    m = msg(subject=encoded, frm="jean@client.fr")
    v = ui.classify_unsubscribe(m, "1", CFG)
    assert v.action == ui.ACTION_SUPPRESS and v.email == "jean@client.fr", v
    print("  RFC 2047 decoding, accented subject detected: OK")
    return True


def test_barrier0_not_addressed_to_us():
    """The alias case: the INBOX holds all of a person's mail."""
    m = msg(subject="unsubscribe", frm="jean@client.fr",
            to="prenom.nom@merciraymond.fr")
    v = ui.classify_unsubscribe(m, "1", CFG)
    assert v.action == ui.ACTION_IGNORE, v
    assert "adresse" in v.why
    # No filter configured (dedicated mailbox) accepts everything.
    v2 = ui.classify_unsubscribe(m, "1", CFG_NOFILTER)
    assert v2.action == ui.ACTION_SUPPRESS, v2
    print("  barrier 0 Delivered-To: OK")
    return True


def test_barrier1_dsn_never_unsubscribes():
    """Our List-Unsubscribe puts the prospect's address in the subject, so a
    DSN about that mail would otherwise suppress a real prospect."""
    m = msg(subject="unsubscribe jean@client.fr",
            frm="Mail Delivery Subsystem <mailer-daemon@googlemail.com>")
    v = ui.classify_unsubscribe(m, "1", CFG)
    assert v.action == ui.ACTION_IGNORE, v
    assert v.email == ""
    m2 = msg(subject="unsubscribe jean@client.fr", frm="x@y.fr",
             extra="Content-Type: multipart/report; report-type=delivery-status;\n"
                   ' boundary="b"\n')
    assert ui.classify_unsubscribe(m2, "2", CFG).action == ui.ACTION_IGNORE
    print("  barrier 1 delivery reports: OK")
    return True


def test_barrier2_arf_goes_to_review():
    m = msg(subject="unsubscribe jean@client.fr", frm="staff@hotmail.com",
            extra="Content-Type: multipart/report; report-type=feedback-report;\n"
                  ' boundary="b"\n')
    v = ui.classify_unsubscribe(m, "1", CFG)
    assert v.action == ui.ACTION_REVIEW and v.email == "", v
    print("  barrier 2 spam report to review: OK")
    return True


def test_barrier3_auto_reply():
    """'Réponse automatique : Désinscription' from someone who asked nothing."""
    for extra in ("Auto-Submitted: auto-replied\n", "X-Autoreply: yes\n",
                  "Precedence: bulk\n"):
        m = msg(subject="unsubscribe", frm="bob@client.fr", extra=extra)
        assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_IGNORE, extra
    m = msg(subject="Absence du bureau: desinscription", frm="bob@client.fr")
    assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_IGNORE
    # Auto-Submitted: no is the explicit "I am a real message" value.
    m = msg(subject="unsubscribe", frm="bob@client.fr",
            extra="Auto-Submitted: no\n")
    assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_SUPPRESS
    print("  barrier 3 auto-responders: OK")
    return True


def test_barrier4_subject_must_say_so():
    m = msg(subject="Votre devis pour le projet", frm="jean@client.fr")
    assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_IGNORE
    # A body-only request is never an automatic unsubscribe on this path.
    m = msg(subject="Re: notre proposition", frm="jean@client.fr",
            body="Bonjour, merci de me desinscrire de votre liste.")
    assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_IGNORE
    for subject in ("unsubscribe", "Desinscription", "DÉSINSCRIRE",
                    "Fwd: unsubscribe jean@client.fr"):
        m = msg(subject=subject, frm="jean@client.fr")
        assert ui.classify_unsubscribe(m, "1", CFG).action == ui.ACTION_SUPPRESS, subject
    print("  barrier 4 subject-bound detection: OK")
    return True


def test_barrier5_internal_external_arbitration():
    """One case per row of the arbitration table."""
    # external sender, no address in subject -> the sender acted
    v = ui.classify_unsubscribe(msg("unsubscribe", "Jean <jean@client.fr>"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_SUPPRESS, "jean@client.fr"), v
    assert v.display_name == "Jean"

    # external sender, same address in subject
    v = ui.classify_unsubscribe(
        msg("unsubscribe jean@client.fr", "jean@client.fr"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_SUPPRESS, "jean@client.fr"), v

    # external sender, DIFFERENT external address: the sender is who clicked,
    # the subject address only happened to be there (forwarded newsletter).
    v = ui.classify_unsubscribe(
        msg("unsubscribe alice@other.fr", "bob@client.fr"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_SUPPRESS, "bob@client.fr"), v
    assert v.review_hint == "alice@other.fr"

    # internal sender relaying an external request: here the subject wins
    v = ui.classify_unsubscribe(
        msg("unsubscribe jean@client.fr", "salome@merciraymond.fr"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_SUPPRESS, "jean@client.fr"), v
    assert v.display_name == ""   # the colleague's name is not the prospect's

    # internal sender, no external address: a colleague typing "désinscription"
    # in a subject must not be unsubscribed.
    v = ui.classify_unsubscribe(
        msg("Desinscription: point sur le process", "salome@merciraymond.fr"),
        "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v

    # external sender pointing at one of OUR addresses
    v = ui.classify_unsubscribe(
        msg("unsubscribe salome@merciraymond.fr", "bob@client.fr"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v
    print("  barrier 5 internal/external arbitration (6 cases): OK")
    return True


def test_barrier6_never_our_own_addresses():
    """The exact subject an auto-responder on the unsub box would produce."""
    v = ui.classify_unsubscribe(
        msg("unsubscribe " + UNSUB_BOX, UNSUB_BOX), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v
    # never_suppress catches addresses outside our own domain too
    cfg = ui.ImapConfig(user="a@b.fr", password="p", own_domains=OWN,
                        never_suppress=frozenset({"vip@client.fr"}))
    v = ui.classify_unsubscribe(msg("unsubscribe", "vip@client.fr"), "1", cfg)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v
    print("  barrier 6 our own addresses are never suppressed: OK")
    return True


def test_barrier7_multiple_addresses():
    v = ui.classify_unsubscribe(
        msg("Fwd: unsubscribe -- voir le mail de jean@autre.fr et de a@b.fr",
            "bob@client.fr"), "1", CFG)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v
    assert "jean@autre.fr" in v.review_hint
    assert ui.subject_addresses("a@b.fr et A@B.fr") == ["a@b.fr"]
    print("  barrier 7 several addresses in a subject: OK")
    return True


# --- bounces ----------------------------------------------------------------

def _dsn(status, final="mort@client.fr", diag=None):
    return message_from_string(
        "From: Mail Delivery Subsystem <mailer-daemon@googlemail.com>\n"
        "Subject: Delivery Status Notification (Failure)\n"
        'Content-Type: multipart/report; report-type=delivery-status; boundary="b"\n'
        "\n--b\nContent-Type: message/delivery-status\n\n"
        "Reporting-MTA: dns; googlemail.com\n\n"
        "Final-Recipient: rfc822; %s\nAction: failed\nStatus: %s\n%s"
        "\n--b--\n" % (final, status,
                       ("Diagnostic-Code: smtp; %s\n" % diag) if diag else "")
    )


def test_bounce_bad_address_suppresses():
    """Only the ADDRESSING sub-class (5.1.x) means the address does not exist."""
    for code in ("5.1.1", "5.1.2", "5.1.10"):
        v = ui.classify_bounce(_dsn(code, diag="550 no such user"), "1", CFG_BOUNCE)
        assert (v.action, v.email) == (ui.ACTION_SUPPRESS, "mort@client.fr"), (code, v)
        assert v.reason == ui.REASON_BOUNCE and v.dsn_status == code
        assert v.display_name == ""       # never "Mail Delivery Subsystem"
    print("  bad-address bounce 5.1.x suppresses: OK")
    return True


def test_bounce_policy_goes_to_review():
    """Exchange Online answers "550 5.4.1 Recipient address rejected: Access
    denied" both for an unknown recipient AND for a policy block aimed at the
    SENDER. Auto-suppressing on that lost four live prospects the first time
    this ran. Permanent, yes; non-existent, not established."""
    v = ui.classify_bounce(
        _dsn("5.4.1", final="valide@altarea.com",
             diag="550 5.4.1 Recipient address rejected: Access denied"),
        "1", CFG_BOUNCE)
    assert v.action == ui.ACTION_REVIEW, v
    assert v.email == "", "un blocage de politique ne doit JAMAIS supprimer"
    assert v.review_hint == "valide@altarea.com"
    assert v.dsn_status == "5.4.1"
    for code in ("5.7.1", "5.7.26", "5.2.2", "5.4.4"):
        assert ui.classify_bounce(_dsn(code), "1", CFG_BOUNCE).action == ui.ACTION_REVIEW, code
    print("  policy/routing bounce 5.4.x/5.7.x to review, never suppressed: OK")
    return True


def test_bounce_without_status_code_goes_to_review():
    """Upstream treated an unparseable DSN as permanent, but it matched each one
    to a campaign recipient by Message-ID first. We have no such corroboration,
    so a guess must not be irreversible."""
    m = message_from_string(
        "From: mailer-daemon@googlemail.com\nSubject: failure\n"
        'Content-Type: multipart/report; report-type=delivery-status; boundary="b"\n'
        "\n--b\nContent-Type: message/delivery-status\n\n"
        "Final-Recipient: rfc822; qui@sait.fr\nAction: failed\n--b--\n")
    v = ui.classify_bounce(m, "1", CFG_BOUNCE)
    assert v.action == ui.ACTION_REVIEW and v.email == "", v
    assert v.review_hint == "qui@sait.fr"
    print("  DSN with no status code to review: OK")
    return True


def test_bounce_temporary_does_nothing():
    """Microsoft's 451 4.7.500 fires when a sender changes its sending habits,
    i.e. at the start of a campaign. Treating it as permanent would delete a
    perfectly valid prospect for good."""
    v = ui.classify_bounce(_dsn("4.7.500"), "1", CFG_BOUNCE)
    assert v.action == ui.ACTION_IGNORE, v
    assert v.reason == "deferred" and v.email == "mort@client.fr"
    verdict, code, _ = ui.parse_dsn_verdict(_dsn("4.7.500"))
    assert (verdict, code) == ("temporaire", "4.7.500")
    # A 2.x.x is a delivery receipt, not a failure.
    assert ui.parse_dsn_verdict(_dsn("2.0.0"))[0] == "inconnu"
    print("  soft bounce 4.x.x has no destructive effect: OK")
    return True


def test_bounce_without_recipient_is_ignored():
    m = message_from_string(
        "From: postmaster@x.fr\nSubject: failure\n"
        'Content-Type: multipart/report; report-type=delivery-status; boundary="b"\n'
        "\n--b\nContent-Type: message/delivery-status\n\nStatus: 5.1.1\n--b--\n")
    assert ui.classify_bounce(m, "1", CFG_BOUNCE).action == ui.ACTION_IGNORE
    assert ui.classify_bounce(msg("hello", "a@b.fr"), "1",
                              CFG_BOUNCE).action == ui.ACTION_IGNORE
    print("  bounce with no Final-Recipient ignored: OK")
    return True


def test_bounce_on_our_own_address_goes_to_review():
    v = ui.classify_bounce(_dsn("5.1.1", final="salome@merciraymond.fr"),
                           "1", CFG_BOUNCE)
    assert (v.action, v.email) == (ui.ACTION_REVIEW, ""), v
    print("  bounce on one of our addresses to review: OK")
    return True


# --- "remove me" replies ----------------------------------------------------

def test_reply_opposition_only_flags():
    m = msg(subject="Re: notre proposition", frm="jean@client.fr",
            extra="In-Reply-To: <abc@merciraymond.fr>\n",
            body="Bonjour, merci de ne plus me contacter. Cordialement",
            to="salome@merciraymond.fr")
    v = ui.classify_reply(m, "1", CFG_REPLY)
    assert v.action == ui.ACTION_REVIEW, v
    assert v.email == "", "une formule en texte libre ne doit JAMAIS desinscrire"
    assert v.review_hint == "jean@client.fr"
    # A reply with no such wording, and an internal sender, are both ignored.
    m2 = msg("Re: proposition", "jean@client.fr",
             extra="In-Reply-To: <a@b>\n", body="Oui, tres interesse !")
    assert ui.classify_reply(m2, "1", CFG_REPLY).action == ui.ACTION_IGNORE
    m3 = msg("Re: x", "salome@merciraymond.fr", extra="In-Reply-To: <a@b>\n",
             body="il faut le desinscrire")
    assert ui.classify_reply(m3, "1", CFG_REPLY).action == ui.ACTION_IGNORE
    assert ui.indice_opposition("merci de me RETIREZ-moi") is not None
    assert ui.indice_opposition("tout va bien") is None
    print("  body opposition flags only, never suppresses: OK")
    return True


# --- scan_mailbox with an injected fetcher ----------------------------------

def test_scan_dedups_and_sorts_verdicts():
    items = [
        ("10", msg("unsubscribe", "jean@client.fr")),
        ("11", msg("unsubscribe", "jean@client.fr")),          # same address
        ("12", msg("unsubscribe alice@other.fr", "bob@client.fr")),
        ("13", msg("Devis", "paul@client.fr")),                # no match
        ("14", msg("unsubscribe " + UNSUB_BOX, UNSUB_BOX)),    # protected
    ]
    report = ui.scan_mailbox(CFG, fetcher=lambda cfg, today: (items, 40))
    assert [c.email for c in report.suppress] == ["jean@client.fr", "bob@client.fr"]
    assert report.examined == 5 and report.ignored == 1
    assert report.searched == 40 and report.truncated is False
    # bob carries a review_hint, and the protected one is in review too
    assert len(report.review) == 2, report.review
    assert report.ok is True
    print("  scan dedups, splits suppress/review/ignore: OK")
    return True


def test_scan_never_raises():
    """A broken mailbox must never be able to take the app down."""
    def boom(cfg, today):
        raise RuntimeError("AUTHENTICATIONFAILED (Failure)")
    report = ui.scan_mailbox(CFG, fetcher=boom)
    assert report.ok is False and report.error_kind == "auth", report
    assert report.suppress == []

    def imap_err(cfg, today):
        raise ui.ImapScanError("quota", "OVERQUOTA")
    assert ui.scan_mailbox(CFG, fetcher=imap_err).error_kind == "quota"

    # An unconfigured mailbox is a no-op, not an error path.
    blank = ui.scan_mailbox(ui.ImapConfig())
    assert blank.error_kind == "config" and blank.suppress == []
    print("  scan_mailbox never raises: OK")
    return True


def test_scan_marks_truncation():
    items = [("1", msg("unsubscribe", "a@client.fr"))]
    report = ui.scan_mailbox(
        ui.ImapConfig(user="u", password="p", own_domains=OWN, max_messages=1),
        fetcher=lambda cfg, today: (items, 500))
    assert report.truncated is True and report.searched == 500
    print("  truncation reported when the window overflows: OK")
    return True


def test_error_classification():
    assert ui.classify_error(Exception("AUTHENTICATIONFAILED")) == "auth"
    assert ui.classify_error(Exception("Application-specific password required")) == "auth"
    assert ui.classify_error(Exception("OVERQUOTA")) == "quota"
    assert ui.classify_error(socket.timeout("timed out")) == "network"
    assert ui.classify_error(Exception("[NONEXISTENT] no such mailbox")) == "mailbox"
    assert ui.classify_error(Exception("weird")) == "unknown"
    print("  error to backoff-class mapping: OK")
    return True


def test_deadline_is_wall_clock():
    d = ui._Deadline(0.0)
    assert d.blown() is True
    d = ui._Deadline(30.0)
    assert d.blown() is False and 29.0 < d.left() <= 30.0
    print("  wall-clock deadline: OK")
    return True


def test_search_criteria():
    today = date(2026, 9, 3)
    assert ui.search_criteria(CFG, today) == (
        'SINCE 04-Aug-2026 HEADER Delivered-To "%s"' % UNSUB_BOX)
    assert ui.search_criteria(CFG_NOFILTER, today) == "SINCE 04-Aug-2026"
    assert 'FROM "mailer-daemon"' in ui.search_criteria(CFG_BOUNCE, today)
    assert 'BODY "desinscri"' in ui.search_criteria(CFG_REPLY, today)
    try:
        ui.search_criteria(ui.ImapConfig(role="nope"), today)
    except ui.ImapScanError as exc:
        assert exc.kind == "config"
    else:
        raise AssertionError("role inconnu doit lever")
    print("  SEARCH criteria per role: OK")
    return True


def test_provenance_lines():
    v = ui.classify_unsubscribe(msg("unsubscribe", "Jean <jean@client.fr>"), "7", CFG)
    lines = ui.provenance_lines(v)
    body = "\n".join(lines)
    assert "Raison : unsubscribe_request" in body
    assert "jean@client.fr" in body and "Sujet : unsubscribe" in body
    assert "Detection :" in body
    print("  Notion page body (the reason lives here): OK")
    return True


def test_no_test_opens_a_socket():
    """Guard: every test above must be pure. If one of them ever reaches the
    network, this fails instead of silently hitting Gmail from CI."""
    real = socket.socket

    class Tripwire(real):
        def __init__(self, *a, **k):
            raise AssertionError("un test a ouvert une socket")

    socket.socket = Tripwire
    try:
        ui.classify_unsubscribe(msg("unsubscribe", "a@client.fr"), "1", CFG)
        ui.classify_bounce(_dsn("5.1.1"), "1", CFG_BOUNCE)
        ui.scan_mailbox(CFG, fetcher=lambda cfg, today: ([], 0))
        np.build_suppression_properties(REAL_SCHEMA, "a@b.fr")
    finally:
        socket.socket = real
    print("  no socket opened by the pure paths: OK")
    return True



# --- app wiring -------------------------------------------------------------
# email_automation_app pulls in streamlit, pandas and html2text, so it is
# imported lazily: the 30 pure tests above must keep running on an interpreter
# that only has the standard library.

def _app():
    """The app module, or None when its dependencies are not installed."""
    try:
        import email_automation_app as app
        return app
    except Exception as exc:
        print("  (app non importable: %s)" % str(exc)[:60])
        return None


class FakeResp:
    """Just enough of a requests.Response for the write path."""

    def __init__(self, status, payload=None, text=""):
        self.status_code = status
        self._payload = payload if payload is not None else {}
        self.text = text or repr(self._payload)
        self.headers = {}

    def json(self):
        return self._payload


def test_app_gate_all_branches():
    """Every branch reachable without secrets.toml, because `configured` is
    injected -- reading the module global directly made the blocking case
    untestable."""
    app = _app()
    if app is None:
        return True
    gate = app._suppression_gate
    assert gate({}, configured=False)[0] == "warn"
    action, msg = gate({"loaded_ever": False, "last_error_msg": "404 object_not_found"},
                       configured=True)
    assert action == "block" and "404" in msg, (action, msg)
    assert gate({"loaded_ever": True, "last_fetch": 1000.0, "last_error": 0.0},
                now=1000.0, configured=True) == ("ok", "")
    assert gate({"loaded_ever": True, "last_fetch": 990.0, "last_error": 995.0,
                 "last_error_msg": "timeout"}, now=1000.0, configured=True)[0] == "warn"
    assert gate({"loaded_ever": True, "last_fetch": 0.0, "last_error": 0.0},
                now=5000.0, configured=True)[0] == "warn"
    print("  gate: warn / block / ok / error / stale: OK")
    return True


def test_app_throttle_and_force():
    app = _app()
    if app is None:
        return True
    state = {"last_attempt": 1000.0, "running": False}
    go, why = app._should_sync_now(state, now=1000.0 + 10)
    assert go is False and "trop" in why, (go, why)
    assert app._should_sync_now(state, now=1000.0 + 10, force=True)[0] is True
    assert app._should_sync_now(state, now=1000.0 + 10_000)[0] is True
    # A cycle already running is never doubled up, force or not.
    assert app._should_sync_now({"running": True}, now=0.0, force=True)[0] is False
    print("  throttle skipped by force, never a double cycle: OK")
    return True


def test_app_cc_filter():
    """The CC path used to bypass is_suppressed() entirely."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    with patch.object(app, "_load_suppression", return_value={"bad@x.fr": {}}):
        assert app._filter_cc("good@x.fr, bad@x.fr") == ["good@x.fr"]
        assert app._filter_cc(" BAD@X.FR ") == []
        assert app._filter_cc("") == []
    print("  CC filtered against the suppression list: OK")
    return True


def test_app_add_suppression_is_idempotent():
    """The guard that makes the whole automatic sync safe to replay."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    with patch.object(app, "SUPPRESSION_NOTION_DS_ID", "ds"), \
         patch.object(app, "_load_suppression", return_value={"a@b.fr": {}}), \
         patch.object(app, "_post_suppression_to_notion") as post:
        assert app.add_suppression("A@B.FR") is True
        assert post.call_count == 0, "une adresse deja presente ne doit pas etre reecrite"
    print("  add_suppression writes nothing for a known address: OK")
    return True


def test_app_write_payload_has_no_unknown_property():
    """THE regression test. The old code posted a 'Raison' property that does
    not exist, so Notion answered 400 and every automatic write failed."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    with patch.object(app, "_resolve_suppression_ds_id", return_value="ds1"), \
         patch.object(app, "_suppression_schema", return_value=REAL_SCHEMA), \
         patch.object(app.requests, "post", return_value=FakeResp(200)) as post:
        ok = app._post_suppression_to_notion(
            "a@b.fr", "unsubscribe_request", nom="A B",
            provenance=["Raison : unsubscribe_request"])
    assert ok is True
    body = post.call_args[1]["json"]
    assert set(body["properties"]) == {"Email", "Date", "Nom"}, body["properties"]
    assert "Raison" not in body["properties"]
    assert body["parent"] == {"type": "data_source_id", "data_source_id": "ds1"}
    # With no Raison column, the reason lives in the page body.
    assert "children" in body and body["children"]
    text = body["children"][0]["paragraph"]["rich_text"][0]["text"]["content"]
    assert "unsubscribe_request" in text
    print("  write payload restricted to the real schema: OK")
    return True


def test_app_write_retries_once_on_validation_error():
    """A renamed column must self-heal, not fail forever."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    bad = FakeResp(400, {"code": "validation_error"},
                   '{"message":"Raison is not a property that exists"}')
    with patch.object(app, "_resolve_suppression_ds_id", return_value="ds1"), \
         patch.object(app, "_suppression_schema", return_value=REAL_SCHEMA), \
         patch.object(app.requests, "post",
                      side_effect=[bad, FakeResp(200)]) as post:
        ok = app._post_suppression_to_notion("a@b.fr", "x")
    assert ok is True and post.call_count == 2
    # And the response BODY is kept, which is the part whose absence hid the bug.
    with patch.object(app, "_resolve_suppression_ds_id", return_value="ds1"), \
         patch.object(app, "_suppression_schema", return_value=REAL_SCHEMA), \
         patch.object(app.requests, "post", side_effect=[bad, bad]):
        assert app._post_suppression_to_notion("a@b.fr", "x") is False
    assert "is not a property that exists" in app._suppression_state["last_write_error"]
    print("  validation_error retried once, Notion's words kept: OK")
    return True


def test_app_legacy_fallback_uses_resolved_id():
    """The old fallback passed the RAW secret as database_id, so whenever the
    secret held a data source id the fallback was structurally wrong."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    notfound = FakeResp(404, {"code": "object_not_found"}, "{}")
    with patch.object(app, "SUPPRESSION_NOTION_DS_ID", "RAW-SECRET"), \
         patch.object(app, "_resolve_suppression_ds_id", return_value="resolved-ds"), \
         patch.object(app, "_suppression_schema", return_value=REAL_SCHEMA), \
         patch.object(app.requests, "post",
                      side_effect=[notfound, FakeResp(200)]) as post:
        ok = app._post_suppression_to_notion("a@b.fr", "x")
    assert ok is True and post.call_count == 2
    parent = post.call_args_list[1][1]["json"]["parent"]
    assert parent == {"type": "database_id", "database_id": "resolved-ds"}, parent
    print("  legacy fallback uses the resolved id: OK")
    return True


def test_app_resolve_has_a_negative_cache():
    """Without it, a hanging Notion cost two probes on EVERY call -- and the
    sidebar paste loop calls it once per address."""
    app = _app()
    if app is None:
        return True
    import requests as _rq
    from unittest.mock import patch
    saved = dict(app._suppression_state)
    try:
        app._suppression_state["resolved_ds_id"] = None
        app._suppression_state["resolve_failed_at"] = 0.0
        boom = _rq.RequestException("connection reset")
        with patch.object(app, "SUPPRESSION_NOTION_DS_ID", "ds"), \
             patch.object(app, "NOTION_API_KEY", "k"), \
             patch.object(app.requests, "post", side_effect=boom) as post, \
             patch.object(app.requests, "get", side_effect=boom) as get:
            assert app._resolve_suppression_ds_id() is None
            first = post.call_count + get.call_count
            assert app._resolve_suppression_ds_id() is None
            assert post.call_count + get.call_count == first, (
                "les sondes sont rejouees: le cache negatif ne mord pas")
        assert app._suppression_state["resolve_failed_at"] > 0
    finally:
        app._suppression_state.update(saved)
    print("  negative cache on data-source resolution: OK")
    return True


def test_app_own_domains_and_footer_anchor():
    app = _app()
    if app is None:
        return True
    assert "merciraymond.fr" in app._own_domains()
    # If the footer constant is ever reworded, the per-recipient rewrite stops
    # biting silently. Fail here instead.
    anchor = 'mailto:%s?subject=unsubscribe"' % app.UNSUBSCRIBE_MAILTO
    assert anchor in app.UNSUBSCRIBE_FOOTER_HTML, (
        "le pied de page a change: le remplacement par destinataire ne mord plus")
    out = app.UNSUBSCRIBE_FOOTER_HTML.replace(
        anchor, anchor[:-1] + '%20jean@client.fr"')
    assert "subject=unsubscribe%20jean@client.fr" in out
    assert app.UNSUBSCRIBE_FOOTER_MARKER in out, (
        "le marqueur doit survivre, sinon _ensure_unsubscribe_footer duplique")
    print("  own_domains derived, footer anchor still matches: OK")
    return True



def test_app_humanize_age():
    """A tiny negative age means "just now", not "never": the clock is read
    before the fetch it measures, and printing "jamais" next to a list that had
    just loaded read as a failure."""
    app = _app()
    if app is None:
        return True
    h = app._humanize_age
    assert h(None) == "jamais"
    assert h(-0.3) == "il y a 0 s", h(-0.3)
    assert h(0) == "il y a 0 s"
    assert h(45) == "il y a 45 s"
    assert h(400) == "il y a 6 min"
    assert h(9000) == "il y a 2 h"
    assert h(200000) == "il y a 2 j"
    print("  humanize_age: negative means just now, None means never: OK")
    return True



def test_app_state_survives_a_rerun():
    """The subtlest constraint in the whole feature, and the easiest to undo by
    "tidying" the state back into the main script.

    Streamlit re-executes the entire top level of the main script in a BRAND-NEW
    module namespace on every interaction (verified by stamping a uuid at import
    time: three reruns gave three different namespaces, while an imported module
    kept its identity). So a dict defined in email_automation_app.py is
    recreated on every click, and a TTL, a circuit breaker, a throttle, a
    backoff or a lock built on one is inert across reruns.

    Everything that must outlive a rerun therefore lives in sync_state, which
    sys.modules keeps for the life of the process. If someone moves it back,
    this fails instead of the app silently losing its throttle and its
    cross-session lock.
    """
    app = _app()
    if app is None:
        return True
    import sync_state
    assert app._suppression_state is sync_state.suppression, (
        "le cache de suppression doit vivre dans sync_state, pas dans le script")
    assert app._unsub_state is sync_state.unsub, (
        "l'etat de la releve doit vivre dans sync_state, pas dans le script")
    assert app._suppression_lock is sync_state.suppression_lock
    assert app._unsub_lock is sync_state.imap_lock, (
        "le verrou inter-sessions doit vivre dans sync_state, sinon chaque "
        "rerun en cree un neuf et deux sessions scannent en parallele")
    # And the state module must stay free of streamlit, or it would drag the
    # whole UI into a test that only wants to inspect a dict.
    src = pathlib.Path(sync_state.__file__).read_text(encoding="utf-8")
    assert "import streamlit" not in src
    print("  shared state lives in a module that survives a rerun: OK")
    return True



def test_app_round_robin_covers_every_mailbox():
    """The round-robin used to advance TWICE per cycle, because _unsub_configs()
    both built the configs and was called again to test whether there was any
    work. With six senders and a step of two that visited indices 0, 2 and 4
    forever: three of the team's mailboxes were never swept, so their hard
    bounces were never seen. The counter now moves once, in the worker."""
    app = _app()
    if app is None:
        return True
    from unittest.mock import patch
    senders = [{"name": "u%d" % i, "email": "u%d@merciraymond.fr" % i,
                "password": "p"} for i in range(6)]
    with patch.object(app, "_load_users", return_value=senders), \
         patch.object(app, "UNSUB_SCAN_BOUNCES", True), \
         patch.object(app, "UNSUB_IMAP_USER", ""):
        app._unsub_state["rr"] = 0
        app._unsub_state["backoff"] = {}
        visited = []
        for _ in range(6):
            app._unsub_advance_rr()
            boxes = {c.user for c in app._unsub_configs() if c.role == "bounce"}
            visited.append(boxes.pop() if boxes else None)
        assert len(set(visited)) == 6, visited
        # And the "is there work" test must not move the counter.
        before = app._unsub_state["rr"]
        assert app._unsub_has_work() is True
        assert app._unsub_state["rr"] == before, (
            "_unsub_has_work a un effet de bord: c'est ce bug qui sautait des boites")
    print("  round-robin visits all six mailboxes, has_work is pure: OK")
    return True


def test_app_has_work_respects_backoff():
    app = _app()
    if app is None:
        return True
    import time as _t
    from unittest.mock import patch
    senders = [{"name": "u", "email": "u@merciraymond.fr", "password": "p"}]
    with patch.object(app, "_load_users", return_value=senders), \
         patch.object(app, "UNSUB_SCAN_BOUNCES", True), \
         patch.object(app, "UNSUB_IMAP_USER", ""):
        app._unsub_state["backoff"] = {}
        assert app._unsub_has_work() is True
        app._unsub_state["backoff"] = {"u@merciraymond.fr": _t.time() + 3600}
        assert app._unsub_has_work() is False, (
            "une boite en backoff ne doit pas relancer de cycle")
        app._unsub_state["backoff"] = {}
    print("  a backed-off mailbox starts no cycle: OK")
    return True



def test_bulk_mail_is_never_a_request():
    """Every newsletter footer says "unsubscribe", so without a bulk filter the
    review queue fills with Substack digests instead of prospect replies -- and
    a list mailing that happened to have a matching subject would look like a
    request. RFC 2369 list headers are the discriminator."""
    bulk = ("Subject: unsubscribe\nFrom: news@substack.com\n"
            "Delivered-To: %s\nList-Unsubscribe: <mailto:x@substack.com>\n\n"
            % UNSUB_BOX)
    v = ui.classify_unsubscribe(message_from_string(bulk), "1", CFG)
    assert v.action == ui.ACTION_IGNORE and v.email == "", v
    assert "masse" in v.why

    digest = ("Subject: How CI/CD Works\nFrom: news@substack.com\n"
              "In-Reply-To: <a@b>\nList-Id: <sysdesign.substack.com>\n\n"
              "... cliquez ici pour vous desinscrire ...")
    v = ui.classify_reply(message_from_string(digest), "1", CFG_REPLY)
    assert v.action == ui.ACTION_IGNORE, v

    # A genuine prospect reply carries none of those headers and still flags.
    real = ("Subject: Re: notre proposition\nFrom: jean@client.fr\n"
            "In-Reply-To: <abc@merciraymond.fr>\n\n"
            "Merci de ne plus me contacter.")
    v = ui.classify_reply(message_from_string(real), "1", CFG_REPLY)
    assert v.action == ui.ACTION_REVIEW and v.review_hint == "jean@client.fr", v
    assert ui.looks_like_bulk(message_from_string(real)) is False
    print("  bulk mail ignored on both roles, real reply still flagged: OK")
    return True


TESTS = [
    ("notion payload, real schema", test_notion_payload_real_schema),
    ("notion payload invariant", test_notion_payload_invariant),
    ("notion Raison when it exists", test_notion_payload_reason_when_column_exists),
    ("notion title by type", test_notion_title_found_by_type),
    ("notion missing title", test_notion_no_title_raises),
    ("notion ambiguous date", test_notion_ambiguous_date_is_skipped),
    ("notion clip and readers", test_notion_clip_and_schema_readers),
    ("SINCE token vs locale", test_since_token_ignores_locale),
    ("RFC 2047 decoding", test_decode_rfc2047),
    ("barrier 0 Delivered-To", test_barrier0_not_addressed_to_us),
    ("barrier 1 DSN", test_barrier1_dsn_never_unsubscribes),
    ("barrier 2 ARF", test_barrier2_arf_goes_to_review),
    ("barrier 3 auto-reply", test_barrier3_auto_reply),
    ("barrier 4 subject-bound", test_barrier4_subject_must_say_so),
    ("barrier 5 arbitration", test_barrier5_internal_external_arbitration),
    ("barrier 6 own addresses", test_barrier6_never_our_own_addresses),
    ("barrier 7 several addresses", test_barrier7_multiple_addresses),
    ("bounce bad address", test_bounce_bad_address_suppresses),
    ("bounce policy to review", test_bounce_policy_goes_to_review),
    ("bounce no status code", test_bounce_without_status_code_goes_to_review),
    ("bounce temporary", test_bounce_temporary_does_nothing),
    ("bounce without recipient", test_bounce_without_recipient_is_ignored),
    ("bounce on our address", test_bounce_on_our_own_address_goes_to_review),
    ("reply opposition flags only", test_reply_opposition_only_flags),
    ("bulk mail never a request", test_bulk_mail_is_never_a_request),
    ("scan dedup and split", test_scan_dedups_and_sorts_verdicts),
    ("scan never raises", test_scan_never_raises),
    ("scan truncation", test_scan_marks_truncation),
    ("error classification", test_error_classification),
    ("wall-clock deadline", test_deadline_is_wall_clock),
    ("SEARCH criteria", test_search_criteria),
    ("provenance lines", test_provenance_lines),
    ("no socket opened", test_no_test_opens_a_socket),
    ("app gate branches", test_app_gate_all_branches),
    ("app throttle/force", test_app_throttle_and_force),
    ("app CC filter", test_app_cc_filter),
    ("app add_suppression idempotent", test_app_add_suppression_is_idempotent),
    ("app write payload", test_app_write_payload_has_no_unknown_property),
    ("app validation_error retry", test_app_write_retries_once_on_validation_error),
    ("app legacy fallback id", test_app_legacy_fallback_uses_resolved_id),
    ("app resolve negative cache", test_app_resolve_has_a_negative_cache),
    ("app domains + footer anchor", test_app_own_domains_and_footer_anchor),
    ("app humanize_age", test_app_humanize_age),
    ("app state survives rerun", test_app_state_survives_a_rerun),
    ("app round-robin coverage", test_app_round_robin_covers_every_mailbox),
    ("app has_work backoff", test_app_has_work_respects_backoff),
]


def main():
    print("Tests: unsubscribe sync (Notion payload + mailbox parsing)")
    print("-" * 60)
    ok = True
    for name, fn in TESTS:
        try:
            fn()
        except Exception as exc:
            print("  %s: FAIL - %s" % (name, exc))
            ok = False
    print("-" * 60)
    print("%d tests" % len(TESTS), "- OK" if ok else "- FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
