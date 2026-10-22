#!/usr/bin/env python3
"""
Tests for plan features: to_name_case, user load, CSV load.
(Adding a sender from the form is covered by test_senders_store.py.)
Run from project root: python test_plan_features.py
"""

import json
import sys
import os
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd

from email_automation_app import (
    to_name_case,
    _load_users,
    EmailAutomation,
)


def test_to_name_case():
    """Names and surnames appear in proper name case."""
    assert to_name_case("JEAN-PIERRE DUPONT") == "Jean-Pierre Dupont"
    assert to_name_case("françois moreau") == "François Moreau"
    assert to_name_case("ARNOLD") == "Arnold"
    assert to_name_case("") == ""
    assert to_name_case("  ") == ""
    assert to_name_case("marie-claire") == "Marie-Claire"
    print("  to_name_case: OK")
    return True


def test_extract_contact_info_name_case():
    """extract_contact_info returns proper name case for full name columns."""
    automation = EmailAutomation()
    df = pd.DataFrame({
        "email": ["a@b.fr"],
        "Nom du contact": ["JEAN-PIERRE DUPONT"],
    })
    mapping = automation.detect_column_mapping(df)
    email_col = mapping["email_column"]
    placeholders = mapping["available_placeholders"]
    full_name_cols = mapping["full_name_columns"]
    row = df.iloc[0]
    info = automation.extract_contact_info(
        row, email_col, placeholders, full_name_cols, name_value_columns=mapping.get("name_value_columns", [])
    )
    assert "Nom du contact_first" in info and info["Nom du contact_first"] == "Jean-Pierre"
    assert "Nom du contact_last" in info and info["Nom du contact_last"] == "Dupont"
    assert info.get("Nom du contact") == "Jean-Pierre Dupont"
    print("  extract_contact_info name case: OK")
    return True


def test_name_value_columns_prenom_nom():
    """Separate Prénom / Nom columns get to_name_case."""
    automation = EmailAutomation()
    df = pd.DataFrame({
        "email": ["x@y.fr"],
        "Prénom": ["jean-pierre"],
        "Nom": ["MOREAU"],
    })
    mapping = automation.detect_column_mapping(df)
    row = df.iloc[0]
    info = automation.extract_contact_info(
        row,
        mapping["email_column"],
        mapping["available_placeholders"],
        mapping["full_name_columns"],
        name_value_columns=mapping.get("name_value_columns", []),
    )
    assert info.get("Prénom") == "Jean-Pierre"
    assert info.get("Nom") == "Moreau"
    print("  name_value_columns (Prénom/Nom): OK")
    return True


def test_load_users_no_file():
    """_load_users returns at least the [users] blocks of the secrets. The
    Notion part is stubbed out: this test must not depend on the network."""
    from unittest.mock import patch
    with patch("email_automation_app._load_app_senders", return_value=[]):
        users = _load_users()
    assert isinstance(users, list)
    assert len(users) >= 5
    names = [u["name"] for u in users]
    assert "Hugo Meunier" in names
    print("  _load_users (secrets senders): OK")
    return True


def test_csv_loading():
    """CSV load (comma and semicolon) produces correct DataFrame."""
    # Comma-separated
    csv_comma = b"email,name\ntest@a.com,Jean Dupont"
    buf = BytesIO(csv_comma)
    df = pd.read_csv(buf, encoding="utf-8", sep=None, engine="python")
    assert list(df.columns) == ["email", "name"]
    assert len(df) == 1 and df.iloc[0]["email"] == "test@a.com"

    # Semicolon-separated
    csv_semi = b"email;name\ntest2@b.fr;Marie Martin"
    buf2 = BytesIO(csv_semi)
    df2 = pd.read_csv(buf2, encoding="utf-8", sep=None, engine="python")
    assert list(df2.columns) == ["email", "name"]
    assert len(df2) == 1 and df2.iloc[0]["email"] == "test2@b.fr"
    print("  CSV load (comma/semicolon): OK")
    return True


def main():
    print("Tests: plan features (to_name_case, users, CSV)")
    print("-" * 50)
    ok = True
    for name, fn in [
        ("to_name_case", test_to_name_case),
        ("extract_contact_info name case", test_extract_contact_info_name_case),
        ("name_value_columns Prénom/Nom", test_name_value_columns_prenom_nom),
        ("_load_users", test_load_users_no_file),
        ("CSV loading", test_csv_loading),
    ]:
        try:
            fn()
        except Exception as e:
            print(f"  {name}: FAIL - {e}")
            ok = False
    print("-" * 50)
    print("OK" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
