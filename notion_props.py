"""Notion property payloads built FROM a live schema (never from hardcoded names).

Why this module exists: the app used to POST a "Raison" property that does not
exist in the suppression database (whose real schema is Email/title, Date/date,
Nom/rich_text). Notion rejects ANY unknown property name with HTTP 400
validation_error, so every automatic write failed silently -- the error body,
which literally names the offending property, was never read.

The fix is to stop guessing. Callers hand us {property_name: property_type} as
Notion actually reports it, and we emit only properties that exist. The single
invariant, asserted below and locked by a test:

    set(returned_properties) <= set(schema)

Consequence: the day someone adds a "Raison" column in Notion, it gets filled
automatically -- no code change, no redeploy. And anything we cannot store in a
property (the reason, today) goes into the page BODY via build_provenance_children,
which depends on no schema at all and therefore can never be rejected.

Pure stdlib, no I/O, no streamlit: importable and testable on its own.
"""

from __future__ import annotations

from typing import Dict, List, Optional

# Notion caps a single rich_text/title content at 2000 chars, and a select
# option name at 100. Exceeding either is a 400, so we clip rather than fail.
NOTION_TEXT_LIMIT = 2000
NOTION_SELECT_LIMIT = 100

# Candidate names per role, most likely first. Matching is exact (Notion
# property names are case-sensitive) and always cross-checked against the type.
_REASON_NAMES = ("Raison", "Reason", "Motif")
_DATE_NAMES = ("Date", "Date d'ajout", "Ajoute le", "Ajouté le", "Added")
_NAME_NAMES = ("Nom", "Name", "Contact", "Prenom", "Prénom")
_EMAIL_NAMES = ("Email", "E-mail", "Mail", "Adresse", "Adresse email")

# Types we know how to write a reason into, in order of preference.
_REASON_WRITABLE = ("rich_text", "select", "multi_select")


class NotionSchemaError(ValueError):
    """The schema cannot hold an email at all (no title property)."""


def clip(text: str, limit: int = NOTION_TEXT_LIMIT) -> str:
    """Trim to Notion's per-field ceiling. Empty-safe."""
    if not text:
        return ""
    text = str(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def schema_from_page(page: dict) -> Dict[str, str]:
    """{property_name: property_type} read off a Notion PAGE object.

    A page carries every property of its database, including the empty ones,
    so a plain query gives us the full schema for free -- no extra request on
    the hot path. The existing read code already relies on this (it does
    props.get("Date") on rows whose date is empty).
    """
    props = (page or {}).get("properties") or {}
    return {
        name: str(pdata.get("type") or "")
        for name, pdata in props.items()
        if isinstance(pdata, dict) and pdata.get("type")
    }


def schema_from_data_source(payload: dict) -> Dict[str, str]:
    """Same mapping, from GET /v1/data_sources/{id} (or a legacy database GET).

    Only needed when the database is empty: no rows means no page, means no
    free schema. Defensive about where 'properties' sits across API versions.
    """
    payload = payload or {}
    props = payload.get("properties")
    if not isinstance(props, dict):
        props = ((payload.get("schema") or {}).get("properties")) or {}
    return {
        name: str(pdata.get("type") or "")
        for name, pdata in props.items()
        if isinstance(pdata, dict) and pdata.get("type")
    }


def title_property(schema: Dict[str, str]) -> Optional[str]:
    """Name of the title property, found BY TYPE.

    The read path already does this (it scans for type == "title" and ignores
    the name); doing the same on write is what removes the last hardcoded
    property name, and keeps read and write symmetric.
    """
    for name, ptype in (schema or {}).items():
        if ptype == "title":
            return name
    return None


def _pick(schema: Dict[str, str], names, wanted_types) -> Optional[str]:
    """First candidate name that exists AND has one of the accepted types."""
    for name in names:
        if schema.get(name) in wanted_types:
            return name
    return None


def build_suppression_properties(
    schema: Dict[str, str],
    email_lc: str,
    *,
    reason: str = "",
    day: str = "",
    nom: str = "",
) -> dict:
    """'properties' payload for POST /v1/pages, restricted to what exists.

    Raises NotionSchemaError when the schema has no title property, because
    then there is nowhere to put the address and a silent partial write would
    be worse than a loud failure.
    """
    schema = schema or {}
    props: Dict[str, dict] = {}

    title_name = title_property(schema)
    if not title_name:
        raise NotionSchemaError(
            "la base de desinscription n'a aucune propriete de type 'title' "
            "-- impossible d'y ecrire une adresse"
        )
    props[title_name] = {"title": [{"text": {"content": clip(email_lc)}}]}

    # Date: an exact known name first. Otherwise the SOLE date property -- with
    # two or more we would be guessing, and writing the wrong date is worse
    # than writing none.
    if day:
        dname = _pick(schema, _DATE_NAMES, ("date",))
        if dname is None:
            dates = [n for n, t in schema.items() if t == "date"]
            dname = dates[0] if len(dates) == 1 else None
        if dname:
            props[dname] = {"date": {"start": day}}

    # Reason: written according to its REAL type. Absent today, so this block
    # is a no-op until someone adds the column -- then it fills itself.
    # 'status' is deliberately excluded: its options cannot be created through
    # the API, so writing an unknown one is a guaranteed 400.
    if reason:
        rname = _pick(schema, _REASON_NAMES, _REASON_WRITABLE)
        rtype = schema.get(rname) if rname else None
        if rtype == "rich_text":
            props[rname] = {"rich_text": [{"text": {"content": clip(reason)}}]}
        elif rtype == "select":
            props[rname] = {"select": {"name": clip(reason, NOTION_SELECT_LIMIT)}}
        elif rtype == "multi_select":
            props[rname] = {
                "multi_select": [{"name": clip(reason, NOTION_SELECT_LIMIT)}]
            }

    # Nom: the sender's DISPLAY NAME, never the reason. Every one of the 15
    # hand-typed rows has a real human name there; writing "bounce_permanent"
    # into it would corrupt the only column the team reads. Left absent (hence
    # empty in Notion) when we have nothing.
    if nom:
        nname = _pick(schema, _NAME_NAMES, ("rich_text",))
        if nname and nname != title_name:
            props[nname] = {"rich_text": [{"text": {"content": clip(nom)}}]}

    # A property of Notion type 'email' whose name looks like an address field:
    # free redundancy, useful for sorting/filtering in Notion views.
    ename = _pick(schema, _EMAIL_NAMES, ("email",))
    if ename:
        props[ename] = {"email": clip(email_lc, NOTION_SELECT_LIMIT)}

    # The invariant. An assert is right here: a violation is a programming
    # error in this function, not a runtime condition to recover from.
    assert set(props) <= set(schema), (
        "propriete(s) hors schema: %s" % sorted(set(props) - set(schema))
    )
    return props


def build_provenance_children(lines: List[str]) -> list:
    """One paragraph block for the page body.

    'children' is validated against no schema whatsoever, so this is the only
    place where a trace can be written that Notion cannot refuse. With no
    'Raison' column, this is where the reason lives -- and it stays useful
    afterwards, as the way to tell a machine-created row from a hand-typed one.
    """
    text = "\n".join(str(line) for line in (lines or []) if line)
    if not text:
        return []
    return [
        {
            "object": "block",
            "type": "paragraph",
            "paragraph": {
                "rich_text": [{"type": "text", "text": {"content": clip(text)}}]
            },
        }
    ]
