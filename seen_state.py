"""Memory of emails already handled, so an hourly run can tell "nothing new" from "something to do".

Saved as DOI_Seen_Messages.csv next to the run log. One row per email the program has judged:
  ref              the email's short tag (same 'ref xxxxxxxx' as in the tracker's Notes)
  result           done     - accepted and written to the tracker: never opened again
                   skipped  - accepted, but its request isn't in the tracker (re-checked quietly)
                   review   - waiting for a person (re-checked quietly - e.g. after you approve
                              a sender in DOI_Contacts.csv it is processed on the next run)
                   reject   - not a DOI request (re-checked quietly)
  request_id, conversation_id   so later replies with a changed subject still find their request
  first_seen, last_seen

An email is "news" only when it's new, or its outcome changed. Deleting the file is safe: the
next run simply treats everything as new once (the tracker itself never duplicates).
"""
import csv
from datetime import datetime, timedelta
from pathlib import Path

FIELDS = ["ref", "result", "request_id", "conversation_id", "first_seen", "last_seen"]
QUIET = {"done", "skipped"}                  # accepted outcomes; "accept" again is not news


def load(path) -> dict:
    path = Path(path)
    if not path.exists():
        return {}
    try:
        with open(path, newline="", encoding="utf-8") as f:
            return {row["ref"]: {k: row.get(k, "") or "" for k in FIELDS} for row in csv.DictReader(f) if row.get("ref")}
    except (OSError, KeyError, csv.Error):
        return {}                                # unreadable memory = start fresh, never crash


def done_refs(state) -> set:
    """Emails fully processed - the Outlook fetch doesn't even open these again."""
    return {ref for ref, row in state.items() if row["result"] == "done"}


def known_threads(state) -> dict:
    """conversation_id -> request_id for requests already in the tracker."""
    return {row["conversation_id"]: row["request_id"] for row in state.values()
            if row["conversation_id"] and row["request_id"] and row["result"] in QUIET}


def is_news(state, ref, status) -> bool:
    """status: accept | review | reject (this run's verdict)."""
    previous = state.get(ref, {}).get("result")
    if status == "accept":
        return previous not in QUIET
    return previous != status


def record(state, ref, result, request_id="", conversation_id="", when=None):
    when = f"{(when or datetime.now()):%Y-%m-%d %H:%M}"
    row = state.setdefault(ref, {k: "" for k in FIELDS} | {"ref": ref, "first_seen": when})
    row.update(result=result, last_seen=when)
    if request_id:
        row["request_id"] = request_id
    if conversation_id:
        row["conversation_id"] = conversation_id


def touch(state, ref, when=None):
    if ref in state:
        state[ref]["last_seen"] = f"{(when or datetime.now()):%Y-%m-%d %H:%M}"


def prune(state, keep_days: int, now=None):
    """Forget emails not seen for keep_days - they're outside every fetch window by then."""
    cutoff = f"{((now or datetime.now()) - timedelta(days=keep_days)):%Y-%m-%d %H:%M}"
    for ref in [r for r, row in state.items() if row["last_seen"] and row["last_seen"] < cutoff]:
        del state[ref]


def save(path, state) -> str:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS)
            w.writeheader()
            w.writerows(sorted(state.values(), key=lambda r: r["first_seen"]))
    except PermissionError:
        return f"seen-messages memory NOT saved: {path.name} is open - next run re-checks those emails"
    return "saved"
