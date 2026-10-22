"""Senders added from the app: Notion rows <-> sender dicts, and the password cipher.

Why this module exists: the "Ajouter un utilisateur" form used to write
users.json next to the script. On Streamlit Cloud that disk is thrown away at
every restart and every redeploy, so a sender added on Monday was gone by
Tuesday. The persistent store is now a Notion database ("Expéditeurs
Raymographe", a child of the Raymographe page), next to the suppression list.

A Gmail app password opens the whole mailbox (SMTP to send, IMAP to read), and
that database is readable by everyone the Raymographe page is shared with. So
the password is never written in clear. It is encrypted with Fernet
(AES-128-CBC for secrecy, HMAC-SHA256 so a tampered value is refused instead
of decrypted to garbage) under a key that exists only in the Streamlit
secrets. Notion holds ciphertext it cannot read; the secrets hold a key but no
data. Either one alone is useless.

Example of what one row looks like in Notion:

    Nom             Email                         Mot de passe chiffré
    Jean Dupont     jean.dupont@merciraymond.fr   gAAAAABm...(about 120 chars)

Same rule as notion_props: columns are found BY TYPE on the live schema, so
renaming a column in Notion breaks nothing. Precedence rule, enforced in
merge_senders(): the [users] blocks of the Streamlit secrets win over Notion,
and the form refuses an address that is already there, so a given address has
exactly one place where its password is edited.

No streamlit and no network here: importable and testable on its own.
`cryptography` is imported lazily, so the pure tests run without it.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from notion_props import NotionSchemaError, clip, schema_from_page, title_property

# Same shape as parse_email_list() in the app, anchored to the whole string.
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")

# A Gmail app password is 16 letters (19 with the spaces Google displays).
# The ceiling only keeps the Fernet token far below Notion's 2000-char limit
# per rich_text value: clipping a token would corrupt it silently.
PASSWORD_MAX_LEN = 200

# Preferred names per role. A column is accepted under another name when it
# is the only one of its type, which is what makes a rename harmless.
_EMAIL_NAMES = ("Email", "E-mail", "Adresse")
_TOKEN_NAMES = ("Mot de passe chiffré", "Mot de passe", "Password")


class SenderKeyError(ValueError):
    """The encryption key is missing, malformed, or cannot be used."""


# --- cipher ------------------------------------------------------------------

def make_cipher(key: str):
    """Fernet cipher for the key held in the secrets. Raises SenderKeyError.

    The key is 32 random bytes written in url-safe base64 (44 characters),
    as produced by Fernet.generate_key().
    """
    if not key or not str(key).strip():
        raise SenderKeyError("clé de chiffrement absente des secrets")
    try:
        from cryptography.fernet import Fernet
    except ImportError as exc:
        raise SenderKeyError("paquet 'cryptography' non installé") from exc
    try:
        return Fernet(str(key).strip().encode("ascii"))
    except (ValueError, TypeError, UnicodeEncodeError) as exc:
        raise SenderKeyError(
            "clé de chiffrement invalide (attendu : 44 caractères base64, "
            "générés par Fernet.generate_key())"
        ) from exc


def encrypt_password(cipher, password: str) -> str:
    """Fernet token as ASCII text. A fresh random IV each call, so encrypting
    the same password twice gives two different tokens."""
    return cipher.encrypt(password.encode("utf-8")).decode("ascii")


def decrypt_password(cipher, token: str) -> Optional[str]:
    """Clear password, or None when this key cannot open the token (another
    key, or a value edited by hand in Notion: the HMAC check catches both)."""
    from cryptography.fernet import InvalidToken
    try:
        return cipher.decrypt(str(token).strip().encode("ascii")).decode("utf-8")
    except (InvalidToken, ValueError, UnicodeError):
        return None


# --- schema ------------------------------------------------------------------

def _column(schema: Dict[str, str], wanted_type: str, names) -> Optional[str]:
    """A preferred name with the right type, else the SOLE column of that type.
    With two or more candidates we would be guessing, so None."""
    for name in names:
        if schema.get(name) == wanted_type:
            return name
    found = [n for n, t in schema.items() if t == wanted_type]
    return found[0] if len(found) == 1 else None


def sender_columns(schema: Dict[str, str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """(name column, email column, encrypted-password column), found by type."""
    schema = schema or {}
    return (
        title_property(schema),
        _column(schema, "email", _EMAIL_NAMES),
        _column(schema, "rich_text", _TOKEN_NAMES),
    )


def build_sender_properties(schema: Dict[str, str], name: str, email: str,
                            token: str) -> dict:
    """'properties' payload for POST or PATCH /v1/pages, restricted to the
    schema. Raises NotionSchemaError rather than write a partial row: a sender
    without a password column would look saved and fail at the first send."""
    name_col, email_col, token_col = sender_columns(schema)
    missing = [label for label, col in (("titre", name_col), ("email", email_col),
                                         ("mot de passe chiffré", token_col)) if not col]
    if missing:
        raise NotionSchemaError(
            "la base des expéditeurs n'a pas de colonne %s identifiable"
            % ", ".join(missing)
        )
    props = {
        name_col: {"title": [{"text": {"content": clip(name)}}]},
        email_col: {"email": email},
        token_col: {"rich_text": [{"text": {"content": token}}]},
    }
    assert set(props) <= set(schema), (
        "propriete(s) hors schema: %s" % sorted(set(props) - set(schema))
    )
    return props


def _plain(items) -> str:
    return "".join(it.get("plain_text", "") for it in (items or [])).strip()


def parse_sender_page(page: dict) -> Optional[dict]:
    """{"name", "email", "token", "page_id"} from one Notion row, or None when
    the row is incomplete (a blank line someone added by hand in Notion)."""
    props = (page or {}).get("properties") or {}
    name_col, email_col, token_col = sender_columns(schema_from_page(page))
    if not (name_col and email_col and token_col):
        return None
    name = _plain(props[name_col].get("title"))
    email = (props[email_col].get("email") or "").strip()
    token = _plain(props[token_col].get("rich_text"))
    if not (name and email and token):
        return None
    return {"name": name, "email": email, "token": token, "page_id": page.get("id", "")}


# --- merge and validation -------------------------------------------------------

def merge_senders(base: List[dict], added: List[dict]) -> List[dict]:
    """The secrets senders first, then the Notion ones whose address is new.

    Deduplicated by lowercased email. The secrets win on a clash, because the
    form refuses those addresses anyway (see validate_new_sender): a clash can
    only come from someone editing both places by hand.
    """
    seen = set()
    out = []
    for user in list(base or []) + list(added or []):
        email_lc = (user.get("email") or "").strip().lower()
        if not email_lc or email_lc in seen:
            continue
        seen.add(email_lc)
        out.append(user)
    return out


def validate_new_sender(name: str, email: str, password: str,
                        base: List[dict], added: List[dict]) -> Optional[str]:
    """Why the form must refuse this sender (in French, shown as is), or None.

    `base` is the secrets senders, `added` the Notion ones. Re-entering an
    address already in Notion is allowed: it updates that row, which is how a
    revoked app password gets replaced.
    """
    name, email, password = (name or "").strip(), (email or "").strip(), (password or "").strip()
    if not (name and email and password):
        return "Nom, email et mot de passe d'application sont obligatoires."
    if not EMAIL_RE.match(email):
        return "Adresse email invalide : %s" % email
    if len(password) > PASSWORD_MAX_LEN:
        return "Mot de passe trop long pour un mot de passe d'application Gmail."
    email_lc = email.lower()
    if any((u.get("email") or "").strip().lower() == email_lc for u in base or []):
        return ("%s est déjà configuré dans les secrets Streamlit : c'est là "
                "qu'il faut modifier son mot de passe." % email)
    # The sidebar selectbox is keyed on the display name, so two senders with
    # one name would make the second one unreachable.
    name_cf = name.casefold()
    for user in list(base or []) + list(added or []):
        other = (user.get("email") or "").strip().lower()
        if (user.get("name") or "").strip().casefold() == name_cf and other != email_lc:
            return "Le nom « %s » est déjà utilisé par %s." % (name, other)
    return None
