"""Process-wide mutable state that has to survive a Streamlit rerun.

This module exists because of a fact about Streamlit that is easy to get wrong,
and that we verified rather than assumed: on every interaction, Streamlit
re-executes the ENTIRE top level of the main script inside a BRAND-NEW module
namespace. Measured by stamping a uuid at import time, three consecutive reruns
gave three different namespaces for email_automation_app, while an imported
module kept the same object identity throughout.

So a dict defined at module level in email_automation_app.py is recreated on
every click. Anything meant to be a process-wide cache, throttle, circuit
breaker or lock has to live in an IMPORTED module instead, because those are
held in sys.modules and are therefore created once per process.

Two things depended on this and were quietly broken:

  * The suppression cache and its circuit breaker (added after the June 2026
    incident, where an unreachable Notion database froze the app for minutes).
    Being reset every rerun, the 5-minute TTL only ever helped WITHIN one run
    -- which is still the case it was written for, since a send loop calls
    is_suppressed() once per contact -- but a broken Notion was re-probed on
    every single rerun, at up to 20 s a time. The breaker could never hold.

  * The mailbox sweep's throttle, backoff, review queue and cross-session lock.
    All of them reset per rerun would have meant one IMAP cycle per user click,
    with no lock between concurrent sessions: precisely the storm the design
    was meant to prevent.

Holding the state here fixes both. The app binds local names to these objects
on each rerun; the names are fresh, the objects are not.

Nothing here imports streamlit, so the state is inspectable from a test.
"""

from __future__ import annotations

import collections
import threading

UNSUB_JOURNAL_MAXLEN = 200

#: Suppression list cache. See the F3 block in email_automation_app.py.
suppression = {
    "data": {},               # {email_lc: {"date": iso, "reason": str}}
    "last_fetch": 0.0,
    "resolved_ds_id": None,   # cached after the first successful resolution
    "resolve_failed_at": 0.0,  # negative cache: stops the two-probe storm
    "loaded_ever": False,
    "last_error": 0.0,        # timestamp of last failed fetch; arms the breaker
    "last_error_msg": "",     # why it failed, so the UI can say more than "0"
    "last_write_error": "",   # Notion's own words on the last refused write
    "schema": None,           # {property_name: property_type}, from the live DB
    "schema_at": 0.0,
}

#: Mailbox sweep state. See the F4 block in email_automation_app.py.
unsub = {
    "last_attempt": 0.0,
    "last_success": 0.0,
    "running": False,
    "thread": None,
    "backoff": {},        # {mailbox: timestamp until which we leave it alone}
    "errors": {},         # {mailbox: "message"}
    "last_error": "",     # most recent error, for the pre-send warning
    "last_error_kind": "",
    "review": {},         # {"mailbox|role": [Candidate]}, replaced per key
    "added_last": 0,
    "duplicates_last": 0,
    "examined_last": 0,
    "added_total": 0,
    "rr": 0,              # round-robin index over the senders' mailboxes
    "journal": collections.deque(maxlen=UNSUB_JOURNAL_MAXLEN),
}

#: Guards the copy-on-write updates of suppression["data"], which the sweep
#: thread writes while reruns iterate it.
suppression_lock = threading.Lock()

#: One IMAP cycle at a time for the whole process. Streamlit Cloud runs every
#: session in one process, so a module-level lock IS the cross-session lock --
#: exactly the role Redis SETNX played in the sibling FastAPI project. It only
#: works because this module persists.
imap_lock = threading.Lock()


def reset_for_tests() -> None:
    """Put both dicts back to a pristine state. Tests only."""
    suppression.update({
        "data": {}, "last_fetch": 0.0, "resolved_ds_id": None,
        "resolve_failed_at": 0.0, "loaded_ever": False, "last_error": 0.0,
        "last_error_msg": "", "last_write_error": "", "schema": None,
        "schema_at": 0.0,
    })
    unsub.update({
        "last_attempt": 0.0, "last_success": 0.0, "running": False,
        "thread": None, "backoff": {}, "errors": {}, "last_error": "",
        "last_error_kind": "", "review": {}, "added_last": 0,
        "duplicates_last": 0, "examined_last": 0, "added_total": 0, "rr": 0,
    })
    unsub["journal"].clear()
