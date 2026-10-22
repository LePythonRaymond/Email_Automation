#!/usr/bin/env python3
"""
Tests for the senders added from the app ("Ajouter un utilisateur").
Run from project root: python3.12 test_senders_store.py
(the pure tests need only `cryptography`; the app-level ones skip themselves
if streamlit is absent.)

No test touches Notion or Gmail: Notion is an in-memory fake behind
requests.post/get/patch, Gmail a fake smtplib.SMTP.

Two of these tests exist because of the real incident of 2026-10-22:
  * test_ui_add_sender_no_exception replays the exact clicks of the bug
    report. The old handler wrote st.session_state.utilisateur_select after
    the selectbox was drawn, and Streamlit raised StreamlitAPIException.
  * test_ui_sender_survives_restart locks the "temporary only" complaint: the
    old form wrote users.json on a disk Streamlit Cloud wipes at each restart.
"""

import os
import smtplib
import sys
import uuid
from unittest.mock import patch

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import notion_props
import senders_store as ss

from cryptography.fernet import Fernet

KEY = Fernet.generate_key().decode("ascii")
LIVE_SCHEMA = {"Nom": "title", "Email": "email",
               "Mot de passe chiffré": "rich_text", "Ajouté le": "created_time"}
SECRETS_USERS = [
    {"name": "Salomé Cremona", "email": "salome.cremona@merciraymond.fr", "password": "p1"},
    {"name": "Diane", "email": "dianedemagnitot@merciraymond.fr", "password": "p2"},
]


def page(name, email, token, schema=None, page_id=None):
    """A Notion page object as the query endpoint returns it."""
    schema = schema or LIVE_SCHEMA
    props = {}
    for col, typ in schema.items():
        if typ == "title":
            props[col] = {"type": "title", "title": [{"plain_text": name}] if name else []}
        elif typ == "email":
            props[col] = {"type": "email", "email": email or None}
        elif typ == "rich_text":
            props[col] = {"type": "rich_text", "rich_text": [{"plain_text": token}] if token else []}
        else:
            props[col] = {"type": typ, typ: None}
    return {"object": "page", "id": page_id or str(uuid.uuid4()), "properties": props}


# --- senders_store: cipher ----------------------------------------------------

def test_cipher_roundtrip():
    c = ss.make_cipher(KEY)
    t1, t2 = ss.encrypt_password(c, "abcd efgh ijkl mnop"), ss.encrypt_password(c, "abcd efgh ijkl mnop")
    assert t1 != t2, "un IV aleatoire par chiffrement: deux jetons differents"
    assert "abcd" not in t1
    assert ss.decrypt_password(c, t1) == ss.decrypt_password(c, t2) == "abcd efgh ijkl mnop"
    # A token is ~120 chars for an app password: far below Notion's 2000.
    assert len(t1) < 300, len(t1)
    print("  encrypt/decrypt round trip, random IV, short token: OK")
    return True


def test_cipher_refuses_wrong_key_and_tampering():
    c = ss.make_cipher(KEY)
    token = ss.encrypt_password(c, "secret")
    other = ss.make_cipher(Fernet.generate_key().decode())
    assert ss.decrypt_password(other, token) is None
    tampered = token[:-6] + ("A" if token[-6] != "A" else "B") + token[-5:]
    assert ss.decrypt_password(c, tampered) is None
    assert ss.decrypt_password(c, "pas un jeton") is None
    print("  wrong key, tampered or junk token -> None, never garbage: OK")
    return True


def test_make_cipher_errors():
    for bad in ("", "   ", None, "trop-court", "é" * 44):
        try:
            ss.make_cipher(bad)
        except ss.SenderKeyError:
            continue
        raise AssertionError("cle %r acceptee" % (bad,))
    print("  missing / malformed key -> SenderKeyError: OK")
    return True


# --- senders_store: schema ----------------------------------------------------

def test_columns_found_by_type():
    assert ss.sender_columns(LIVE_SCHEMA) == ("Nom", "Email", "Mot de passe chiffré")
    renamed = {"Expéditeur": "title", "Adresse Gmail": "email", "Secret": "rich_text"}
    assert ss.sender_columns(renamed) == ("Expéditeur", "Adresse Gmail", "Secret")
    # Two rich_text columns and neither has a known name: refuse to guess.
    ambiguous = {"Nom": "title", "Email": "email", "Note": "rich_text", "Autre": "rich_text"}
    assert ss.sender_columns(ambiguous)[2] is None
    # ...but the known name settles it even with a second rich_text around.
    with_note = dict(LIVE_SCHEMA, Note="rich_text")
    assert ss.sender_columns(with_note)[2] == "Mot de passe chiffré"
    print("  columns found by type, rename-proof, no guessing: OK")
    return True


def test_build_properties():
    props = ss.build_sender_properties(LIVE_SCHEMA, "Jean Dupont", "jean@merciraymond.fr", "gAAAA")
    assert set(props) == {"Nom", "Email", "Mot de passe chiffré"}
    assert set(props) <= set(LIVE_SCHEMA)
    assert props["Email"] == {"email": "jean@merciraymond.fr"}
    assert props["Mot de passe chiffré"]["rich_text"][0]["text"]["content"] == "gAAAA"
    try:
        ss.build_sender_properties({"Nom": "title", "Email": "email"}, "a", "a@b.fr", "t")
    except notion_props.NotionSchemaError as exc:
        assert "mot de passe" in str(exc)
    else:
        raise AssertionError("ecriture partielle acceptee")
    print("  payload restricted to the schema, partial row refused: OK")
    return True


def test_parse_page():
    row = ss.parse_sender_page(page("Jean Dupont", "jean@merciraymond.fr", "tok", page_id="p1"))
    assert row == {"name": "Jean Dupont", "email": "jean@merciraymond.fr", "token": "tok", "page_id": "p1"}
    # A blank line someone added by hand in Notion is skipped, not a crash.
    assert ss.parse_sender_page(page("", "", "")) is None
    assert ss.parse_sender_page(page("X", "x@y.fr", "")) is None
    assert ss.parse_sender_page({}) is None
    print("  row parsed, incomplete rows skipped: OK")
    return True


# --- senders_store: merge and validation ----------------------------------------

def test_merge_secrets_win():
    added = [
        {"name": "Ella DELAITRE", "email": "ella.delaitre@merciraymond.fr", "password": "n1"},
        {"name": "Autre Diane", "email": "DianeDeMagnitot@merciraymond.fr", "password": "n2"},
    ]
    merged = ss.merge_senders(SECRETS_USERS, added)
    assert [u["email"].lower() for u in merged] == [
        "salome.cremona@merciraymond.fr", "dianedemagnitot@merciraymond.fr",
        "ella.delaitre@merciraymond.fr"], merged
    assert merged[1]["password"] == "p2", "les secrets gagnent"
    assert ss.merge_senders([], []) == []
    print("  secrets first and authoritative, dedup case-insensitive: OK")
    return True


def test_validate():
    v = ss.validate_new_sender
    notion = [{"name": "Ella DELAITRE", "email": "ella.delaitre@merciraymond.fr"}]
    assert "obligatoires" in v("", "a@b.fr", "x", SECRETS_USERS, [])
    assert "obligatoires" in v("A", "a@b.fr", "   ", SECRETS_USERS, [])
    assert "invalide" in v("A", "pas-un-email", "x", SECRETS_USERS, [])
    assert "trop long" in v("A", "a@b.fr", "x" * 201, SECRETS_USERS, [])
    msg = v("Diane bis", " DIANEDEMAGNITOT@merciraymond.fr ", "x", SECRETS_USERS, [])
    assert msg and "secrets" in msg, msg
    msg = v("ella delaitre", "autre@merciraymond.fr", "x", SECRETS_USERS, notion)
    assert msg and "déjà utilisé" in msg, msg
    # Same address as an existing Notion row: allowed, it is an update.
    assert v("Ella Delaitre", "Ella.Delaitre@merciraymond.fr", "x", SECRETS_USERS, notion) is None
    assert v("Nouveau", "nouveau@merciraymond.fr", "abcd efgh ijkl mnop", SECRETS_USERS, notion) is None
    print("  validation: fields, format, secrets clash, name clash, update allowed: OK")
    return True


# --- app wiring -----------------------------------------------------------------

def _app():
    try:
        import email_automation_app as app
        return app
    except Exception as exc:
        print("  (app non importable: %s)" % str(exc)[:60])
        return None


class FakeResp:
    def __init__(self, status, payload=None):
        self.status_code = status
        self._payload = payload if payload is not None else {}
        self.text = repr(self._payload)
        self.headers = {}

    def json(self):
        return self._payload

    def raise_for_status(self):
        if self.status_code >= 400:
            import requests
            raise requests.HTTPError("HTTP %d" % self.status_code)


class FakeNotion:
    """The three endpoints the senders store uses, plus empty answers for the
    rest of the app (suppression list, sites) so a full script run works."""

    def __init__(self, ds_id, schema=None):
        self.ds_id = ds_id
        self.schema = schema or LIVE_SCHEMA
        self.rows = []          # page objects, creation order
        self.calls = []

    def post(self, url, headers=None, json=None, timeout=None, **_):
        self.calls.append(("POST", url, json))
        if url.endswith("/data_sources/%s/query" % self.ds_id):
            return FakeResp(200, {"results": list(self.rows), "has_more": False})
        if url.endswith("/v1/pages"):
            assert json["parent"] == {"type": "data_source_id", "data_source_id": self.ds_id}, json
            assert set(json["properties"]) <= set(self.schema), json["properties"]
            new = self._page_from_props(json["properties"], str(uuid.uuid4()))
            self.rows.append(new)
            return FakeResp(200, new)
        return FakeResp(200, {"results": [], "has_more": False})

    def patch(self, url, headers=None, json=None, timeout=None, **_):
        self.calls.append(("PATCH", url, json))
        pid = url.rsplit("/", 1)[-1]
        for i, row in enumerate(self.rows):
            if row["id"] == pid:
                self.rows[i] = self._page_from_props(json["properties"], pid)
                return FakeResp(200, self.rows[i])
        return FakeResp(404, {"code": "object_not_found"})

    def get(self, url, headers=None, timeout=None, **_):
        self.calls.append(("GET", url, None))
        if url.endswith("/data_sources/%s" % self.ds_id):
            return FakeResp(200, {"properties": {k: {"type": t} for k, t in self.schema.items()}})
        return FakeResp(404, {})

    def _page_from_props(self, props, pid):
        name_col, email_col, token_col = ss.sender_columns(self.schema)
        return page(props[name_col]["title"][0]["text"]["content"],
                    props[email_col]["email"],
                    props[token_col]["rich_text"][0]["text"]["content"],
                    schema=self.schema, page_id=pid)

    def patches(self, requests_module):
        return (patch.object(requests_module, "post", self.post),
                patch.object(requests_module, "patch", self.patch),
                patch.object(requests_module, "get", self.get))


class FakeSMTP:
    """smtplib.SMTP stand-in. Accepts a login only with GOOD_PASSWORD."""
    GOOD_PASSWORD = "abcd efgh ijkl mnop"
    logins = []

    def __init__(self, host, port, timeout=None):
        self.host = host

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def starttls(self):
        pass

    def login(self, user, password):
        FakeSMTP.logins.append(user)
        if password != FakeSMTP.GOOD_PASSWORD:
            raise smtplib.SMTPAuthenticationError(535, b"5.7.8 Username and Password not accepted")


def _fresh_state(app):
    import sync_state
    sync_state.reset_for_tests()
    FakeSMTP.logins = []


def _with_store(app, notion):
    """Context managers that point the app at the fake store with a real key."""
    p_post, p_patch, p_get = notion.patches(app.requests)
    return [p_post, p_patch, p_get,
            patch.object(app, "SENDERS_ENCRYPTION_KEY", KEY),
            patch.object(app, "SENDERS_NOTION_DS_ID", notion.ds_id),
            patch.object(app, "NOTION_API_KEY", "secret_test"),
            patch.object(app, "_load_users_from_secrets", return_value=list(SECRETS_USERS)),
            patch.object(app.smtplib, "SMTP", FakeSMTP)]


def _run(cms, fn):
    for cm in cms:
        cm.start()
    try:
        return fn()
    finally:
        for cm in reversed(cms):
            cm.stop()


def test_app_add_creates_then_updates():
    app = _app()
    if app is None:
        return True
    _fresh_state(app)
    notion = FakeNotion("ds-test")

    def scenario():
        err, verb = app.add_app_sender(" Ella DELAITRE ", "ella.delaitre@merciraymond.fr",
                                       FakeSMTP.GOOD_PASSWORD)
        assert (err, verb) == (None, "ajouté"), err
        assert len(notion.rows) == 1
        stored = notion.rows[0]["properties"]["Mot de passe chiffré"]["rich_text"][0]["plain_text"]
        assert FakeSMTP.GOOD_PASSWORD not in stored, "mot de passe en clair dans Notion"
        assert ss.decrypt_password(ss.make_cipher(KEY), stored) == FakeSMTP.GOOD_PASSWORD
        # Write-through: visible at once, without waiting for the 5 min TTL.
        assert any(u["email"] == "ella.delaitre@merciraymond.fr" for u in app._load_users())
        # Same address again: the row is updated, never duplicated.
        err, verb = app.add_app_sender("Ella Delaitre", "ELLA.DELAITRE@merciraymond.fr",
                                       FakeSMTP.GOOD_PASSWORD)
        assert (err, verb) == (None, "mis à jour"), err
        assert len(notion.rows) == 1
        assert [c[0] for c in notion.calls].count("PATCH") == 1
        names = [u["name"] for u in app._load_users()]
        assert names == ["Salomé Cremona", "Diane", "Ella Delaitre"], names

    _run(_with_store(app, notion), scenario)
    print("  add creates an encrypted row, re-add updates it in place: OK")
    return True


def test_app_refusals_write_nothing():
    app = _app()
    if app is None:
        return True
    _fresh_state(app)
    notion = FakeNotion("ds-test")

    def scenario():
        err, _ = app.add_app_sender("Ella", "ella@merciraymond.fr", "mauvais mot de passe")
        assert err and "Gmail refuse" in err, err
        err, _ = app.add_app_sender("Diane 2", "dianedemagnitot@merciraymond.fr", FakeSMTP.GOOD_PASSWORD)
        assert err and "secrets" in err, err
        assert FakeSMTP.logins == ["ella@merciraymond.fr"], \
            "les verifications pures passent AVANT tout appel a Gmail"
        assert not notion.rows
        assert not [c for c in notion.calls if c[0] in ("PATCH",) or c[1].endswith("/v1/pages")]

    _run(_with_store(app, notion), scenario)
    print("  wrong password or secrets clash: refused, nothing written: OK")
    return True


def test_app_no_key_no_network():
    app = _app()
    if app is None:
        return True
    _fresh_state(app)
    with patch.object(app, "SENDERS_ENCRYPTION_KEY", ""), \
         patch.object(app, "NOTION_API_KEY", "secret_test"), \
         patch.object(app.requests, "post", side_effect=AssertionError("reseau")):
        assert app._load_app_senders() == []
        err, _ = app.add_app_sender("A", "a@merciraymond.fr", "x")
    assert err and "clé" in err, err
    print("  no key: store off, no request, explicit refusal: OK")
    return True


def test_app_cache_and_breaker():
    """A Notion that fails must not be re-probed on every rerun (June 2026)."""
    app = _app()
    if app is None:
        return True
    _fresh_state(app)
    good = [{"name": "E", "email": "e@merciraymond.fr", "password": "x", "page_id": "p"}]
    with patch.object(app, "SENDERS_ENCRYPTION_KEY", KEY), \
         patch.object(app, "NOTION_API_KEY", "secret_test"), \
         patch.object(app, "_fetch_senders_from_notion", return_value=(good, [])) as f:
        assert app._load_app_senders(now=1000.0) == good
        assert app._load_app_senders(now=1000.0 + 299) == good
        assert f.call_count == 1, "TTL 5 min"
    with patch.object(app, "SENDERS_ENCRYPTION_KEY", KEY), \
         patch.object(app, "NOTION_API_KEY", "secret_test"), \
         patch.object(app, "_fetch_senders_from_notion", side_effect=RuntimeError("timeout")) as f:
        assert app._load_app_senders(now=1000.0 + 301) == good, "derniere copie valide servie"
        assert app._load_app_senders(now=1000.0 + 330) == good
        assert f.call_count == 1, "disjoncteur: pas de nouvel essai avant 60 s"
        assert "timeout" in app._senders_state["last_error_msg"]
    print("  5 min cache, breaker holds 60 s, last good copy served: OK")
    return True


def test_app_undecryptable_row_is_updated_not_duplicated():
    app = _app()
    if app is None:
        return True
    _fresh_state(app)
    notion = FakeNotion("ds-test")
    old_key = ss.make_cipher(Fernet.generate_key().decode())
    notion.rows.append(page("Ella DELAITRE", "ella.delaitre@merciraymond.fr",
                            ss.encrypt_password(old_key, "ancien"), page_id="p-old"))

    def scenario():
        assert app._load_app_senders() == []
        assert [s["name"] for s in app._senders_state["undecryptable"]] == ["Ella DELAITRE"]
        err, verb = app.add_app_sender("Ella DELAITRE", "ella.delaitre@merciraymond.fr",
                                       FakeSMTP.GOOD_PASSWORD)
        assert (err, verb) == (None, "mis à jour"), err
        assert len(notion.rows) == 1 and notion.rows[0]["id"] == "p-old"
        assert app._senders_state["undecryptable"] == []

    _run(_with_store(app, notion), scenario)
    print("  row under an old key: flagged, then re-entered in place: OK")
    return True


def test_verify_gmail_login_messages():
    app = _app()
    if app is None:
        return True

    class Down(FakeSMTP):
        def __init__(self, *a, **k):
            raise OSError("network unreachable")

    with patch.object(app.smtplib, "SMTP", FakeSMTP):
        assert app._verify_gmail_login("a@b.fr", FakeSMTP.GOOD_PASSWORD) is None
        assert "mot de passe d'application" in app._verify_gmail_login("a@b.fr", "bad")
    with patch.object(app.smtplib, "SMTP", Down):
        assert "Impossible de joindre Gmail" in app._verify_gmail_login("a@b.fr", "x")
    print("  Gmail check: accepted / refused / unreachable: OK")
    return True


# --- the real UI, driven by Streamlit's AppTest --------------------------------

def _apptest(notion):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "email_automation_app.py"), default_timeout=60)
    # Hermetic secrets: nothing from .streamlit/secrets.toml is used.
    at.secrets["NOTION_API_KEY"] = "secret_test"
    at.secrets["UNSUB_SYNC_ENABLED"] = "0"
    at.secrets["users"] = {"salome": dict(SECRETS_USERS[0]), "diane": dict(SECRETS_USERS[1])}
    at.secrets["senders_store"] = {"encryption_key": KEY, "notion_ds_id": notion.ds_id}
    return at


def _ui_run(notion, steps):
    import requests
    p_post, p_patch, p_get = notion.patches(requests)
    with p_post, p_patch, p_get, patch.object(smtplib, "SMTP", FakeSMTP):
        return steps()


def test_ui_add_sender_no_exception():
    """The exact clicks of the bug report, on the real script."""
    try:
        import streamlit  # noqa: F401
    except ImportError:
        return True
    import sync_state
    sync_state.reset_for_tests()
    notion = FakeNotion("ds-ui")

    def steps():
        at = _apptest(notion).run()
        assert not at.exception, [e.message for e in at.exception]
        assert at.selectbox(key="utilisateur_select").options == ["Salomé Cremona", "Diane"]
        at.text_input(key="add_user_name").input("Ella DELAITRE")
        at.text_input(key="add_user_email").input("ella.delaitre@merciraymond.fr")
        at.text_input(key="add_user_password").input(FakeSMTP.GOOD_PASSWORD)
        submit = [b for b in at.sidebar.button if b.label == "Ajouter"]
        assert len(submit) == 1, [b.label for b in at.sidebar.button]
        at = submit[0].click().run()
        assert not at.exception, [e.message for e in at.exception]
        sel = at.selectbox(key="utilisateur_select")
        assert sel.options == ["Salomé Cremona", "Diane", "Ella DELAITRE"], sel.options
        assert sel.value == "Ella DELAITRE", sel.value
        for k in ("add_user_name", "add_user_email", "add_user_password"):
            assert at.text_input(key=k).value == "", (k, at.text_input(key=k).value)
        assert any("Ella DELAITRE" in s.value and "permanent" in s.value
                   for s in at.sidebar.success), [s.value for s in at.sidebar.success]
        assert len(notion.rows) == 1
        return at

    _ui_run(notion, steps)
    print("  UI: add -> no exception, profile selected, fields emptied: OK")
    return notion


def test_ui_sender_survives_restart():
    """Restart = new process: empty caches, no session state, same Notion."""
    try:
        import streamlit  # noqa: F401
    except ImportError:
        return True
    import sync_state
    notion = test_ui_add_sender_no_exception()
    sync_state.reset_for_tests()

    def steps():
        at = _apptest(notion).run()
        assert not at.exception, [e.message for e in at.exception]
        assert "Ella DELAITRE" in at.selectbox(key="utilisateur_select").options
        return at

    _ui_run(notion, steps)
    print("  UI: after a restart the sender is still there: OK")
    return True


def test_ui_wrong_password_keeps_fields():
    try:
        import streamlit  # noqa: F401
    except ImportError:
        return True
    import sync_state
    sync_state.reset_for_tests()
    notion = FakeNotion("ds-ui")

    def steps():
        at = _apptest(notion).run()
        at.text_input(key="add_user_name").input("Ella DELAITRE")
        at.text_input(key="add_user_email").input("ella.delaitre@merciraymond.fr")
        at.text_input(key="add_user_password").input("faux")
        at = [b for b in at.sidebar.button if b.label == "Ajouter"][0].click().run()
        assert not at.exception, [e.message for e in at.exception]
        assert any("Gmail refuse" in e.value for e in at.sidebar.error), [e.value for e in at.sidebar.error]
        assert at.text_input(key="add_user_name").value == "Ella DELAITRE", "champ conserve"
        assert at.selectbox(key="utilisateur_select").options == ["Salomé Cremona", "Diane"]
        assert not notion.rows

    _ui_run(notion, steps)
    print("  UI: Gmail refusal shown, entry kept on screen, nothing stored: OK")
    return True


TESTS = [
    ("cipher round trip", test_cipher_roundtrip),
    ("cipher wrong key / tampering", test_cipher_refuses_wrong_key_and_tampering),
    ("cipher key errors", test_make_cipher_errors),
    ("columns by type", test_columns_found_by_type),
    ("build properties", test_build_properties),
    ("parse page", test_parse_page),
    ("merge secrets win", test_merge_secrets_win),
    ("validation", test_validate),
    ("app add create/update", test_app_add_creates_then_updates),
    ("app refusals write nothing", test_app_refusals_write_nothing),
    ("app no key", test_app_no_key_no_network),
    ("app cache + breaker", test_app_cache_and_breaker),
    ("app undecryptable row", test_app_undecryptable_row_is_updated_not_duplicated),
    ("app Gmail check", test_verify_gmail_login_messages),
    ("UI add (bug report)", test_ui_add_sender_no_exception),
    ("UI survives restart", test_ui_sender_survives_restart),
    ("UI wrong password", test_ui_wrong_password_keeps_fields),
]


def main():
    print("Tests: senders added from the app (Notion store + form)")
    print("-" * 60)
    ok = True
    for name, fn in TESTS:
        try:
            fn()
        except Exception as exc:
            import traceback
            traceback.print_exc()
            print("  %s: FAIL - %s" % (name, exc))
            ok = False
    print("-" * 60)
    print("%d tests" % len(TESTS), "- OK" if ok else "- FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
