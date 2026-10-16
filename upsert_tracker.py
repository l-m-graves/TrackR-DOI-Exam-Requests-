"""Write qualified DOI requests into the tracker workbook (tblRequests).

Rules (see CLAUDE.md):
  - Key = Request ID. New ID -> append a row; existing ID -> update it.
  - Never type into formula columns; new rows get the formulas copied from the row above.
  - Idempotent: each email leaves a short "ref" tag in Notes, so a re-run skips it.
  - Timestamped backup before every save; refuse to run while Excel has the file open.
"""
import hashlib
import re
import shutil
import zipfile
from copy import copy
from datetime import datetime, time
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

import doi_filter as F
from doi_filter import REQUEST_ID, canonical_id, message_ref
from due_dates import find_due, find_expected
from email_types import (ADVANCES_TO_IN_PROGRESS, CREATES_ROW, DUE_DATE_KINDS, LABEL,
                         classify, topic_from_subject)
from doi_analysis import analyze

SHEET = "DOI Exam Tracker"
TABLE = "tblRequests"
LAST_ROW = 1000         # the trk* named ranges and dropdowns stop at row 1000

# Columns the script fills (letters match the sheet). Everything else is either
# a formula (A, E, F, H, J, N) or left for a person to type (B, K).
COL_RECEIVED, COL_REQUEST, COL_STATUS, COL_CONTACT, COL_NOTES = "C", "D", "I", "L", "O"
COL_DUE, COL_LAST_SENT = "G", "M"

# File parts openpyxl cannot carry through a save. If the tracker ever gains
# one of these (e.g. a chart on Summary), stop instead of silently deleting it.
LOSSY_PARTS = ("xl/charts/", "xl/media/", "xl/slicers/", "xl/slicerCaches/", "xl/timelines/")
LOSSY_DRAWING = re.compile(r"^xl/drawings/drawing\d+\.xml$")     # comment .vml files are fine

REPLY_PREFIX = re.compile(r"^\s*((re|fw|fwd)\s*:\s*)+", re.IGNORECASE)


class TrackerError(Exception):
    """Something about the workbook means we must not write to it."""


# --- small helpers (pure, unit-tested) ---------------------------------------
def note_ref(entry_id: str) -> str:
    """Short, stable tag for one email (see doi_filter.message_ref - one definition for both)."""
    return message_ref(entry_id)


def first_word_id(text) -> str:
    """Python twin of column A's formula: D's first word with ':' removed ('FIN-52:' -> 'FIN-52'),
    or '' when that word has no digit (the sheet shows '(no ID)')."""
    word = str(text or "").strip().split(" ", 1)[0].replace(":", "")
    return word if any(ch.isdigit() for ch in word) else ""


def request_text(subject: str, rid: str) -> str:
    """Column D text. Team-template subjects ('DOI Exam Request FIN-53 – Topic') become the
    tracker's own pattern 'FIN-53 Request - Topic'. Otherwise: drop RE:/FW: and make sure the
    Request ID is the first word (the sheet's formulas read it from there). 'FIN 45' in a subject
    is written as 'FIN-45' - column A's formula needs the hyphenated form."""
    subject = REQUEST_ID.sub(lambda h: canonical_id(h.group(0)), subject or "")
    topic = topic_from_subject(subject, rid)
    if topic:
        return f"{rid} Request - {topic}"
    text = REPLY_PREFIX.sub("", subject).strip()
    if first_word_id(text) != rid:
        text = f"{rid} {text}"
    return text


def note_line(m, ref: str, extra: str = "", kind: str = "other") -> str:
    extra = f" - {extra}" if extra else ""
    return f"{m.received:%Y-%m-%d} {LABEL[kind]} from {m.sender}: {m.subject}{extra} [{ref}]"


def sent_note_line(s, ref: str, extra: str, kind: str) -> str:
    extra = f" - {extra}" if extra else ""
    return f"{s.sent_on:%Y-%m-%d} {LABEL[kind]} sent: {s.subject}{extra} [{ref}]"


def _day(value):
    """A cell value or date as midnight datetime (the sheet's dates have no time), else None."""
    if isinstance(value, datetime):
        return datetime.combine(value.date(), time())
    if hasattr(value, "year") and hasattr(value, "day"):
        return datetime(value.year, value.month, value.day)
    return None


def contact(m, kind: str = "other", asked=None) -> str:
    """The request's Contact: the person INSIDE the company who was asked to provide the
    information - the To line of an internal email that assigns or asks for it. Left out:
    you (CC'd or forwarded, never the contact), the accountants coordinating the request, the
    team mailbox (MUST_BE_SENT_TO), and anyone on the examiner side. CC is FYI, not asked.
    Nobody left (e.g. the examiner's own email) -> '' and column L stays blank until the
    internal email naming the person arrives. Nothing is guessed."""
    people = asked if asked else (m.to if kind == "assignment" else [])
    keep = [a for a in people
            if a not in F.MY_ADDRESSES and a not in F.MUST_BE_SENT_TO
            and F.role_of(a) not in F.EXAMINER_SIDE | {"accountant"}]
    return "; ".join(f"{m.names[a]} <{a}>" if m.names.get(a) else a for a in keep)


def only_you(value) -> bool:
    """A Contact cell holding nothing but your own address - written by an earlier version of
    this program, never by you on purpose - may be replaced."""
    found = set(re.findall(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+", str(value or "").lower()))
    return bool(found) and found <= F.MY_ADDRESSES


# --- safety checks ------------------------------------------------------------
def check_unlocked(path: Path):
    """openpyxl rewrites the whole file, so Excel must not have it open."""
    if (path.parent / ("~$" + path.name)).exists():       # Excel's owner file while open
        raise TrackerError(f"{path.name} is open in Excel (or Excel crashed) - close it and re-run.")
    try:
        with open(path, "r+b"):                            # Windows refuses this while Excel holds it
            pass
    except PermissionError:
        raise TrackerError(f"{path.name} is locked by another program - close it and re-run.")


def check_lossless(path: Path):
    with zipfile.ZipFile(path) as z:
        risky = [n for n in z.namelist() if n.startswith(LOSSY_PARTS) or LOSSY_DRAWING.match(n)]
    if risky:
        raise TrackerError("Workbook contains parts openpyxl would delete on save "
                           f"(charts/images/slicers): {risky}. Not writing.")


def backup(path: Path, backup_dir: Path = None) -> Path:
    backup_dir = Path(backup_dir or path.parent / "Backups")
    backup_dir.mkdir(parents=True, exist_ok=True)
    dest = backup_dir / f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, dest)                                # copy2 keeps the file's timestamps
    return dest


# --- the upsert ---------------------------------------------------------------
def _table_bounds(ws):
    tbl = ws.tables[TABLE]
    first, last = re.match(r"[A-Z]+(\d+):[A-Z]+(\d+)", tbl.ref).groups()
    return tbl, int(first) + 1, int(last)                  # +1 skips the header row


def _index(ws, first, last) -> dict:
    """Request ID -> row, using the same rule as column A's formula (see first_word_id)."""
    rows = {}
    for r in range(first, last + 1):
        rid = first_word_id(ws[f"{COL_REQUEST}{r}"].value)
        if rid:
            rows.setdefault(rid, r)
    return rows


def _append_row(ws, tbl, first, last) -> int:
    """Add a row under the table: copy formulas (shifted to the new row) and formatting
    from the FIRST data row, then stretch the table to include it. The first row is the
    pattern - not the row above - so one badly formatted row at the bottom (e.g. a stray
    entry typed straight into the table) doesn't spread to every new row."""
    new = last + 1
    if new > LAST_ROW:
        raise TrackerError(f"Tracker is full: named ranges stop at row {LAST_ROW}.")
    if any(ws.cell(new, c).value is not None for c in range(1, ws.max_column + 1)):
        raise TrackerError(f"Row {new} under the table is not empty - won't overwrite it.")
    for src in ws[first]:
        dst = ws.cell(new, src.column)
        dst._style = copy(src._style)
        if isinstance(src.value, str) and src.value.startswith("="):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
    ws.row_dimensions[new].height = ws.row_dimensions[first].height
    start = tbl.ref.split(":")[0]
    end_col = re.match(r"[A-Z]+", tbl.ref.split(":")[1]).group(0)
    tbl.ref = f"{start}:{end_col}{new}"
    if tbl.autoFilter is not None:
        tbl.autoFilter.ref = tbl.ref
    return new


def _due_date(ws, row, stated, changes) -> str:
    """Write a stated due date into an EMPTY Due Date cell; otherwise return a note when it differs."""
    stated, cur = _day(stated), _day(ws[f"{COL_DUE}{row}"].value)
    if ws[f"{COL_DUE}{row}"].value in (None, ""):
        ws[f"{COL_DUE}{row}"] = stated
        changes.append("due date")
        return ""
    if cur != stated:
        return (f"email states due {stated:%Y-%m-%d}; tracker has {cur:%Y-%m-%d} - not changed" if cur
                else f"email states due {stated:%Y-%m-%d} - Due Date cell not changed")
    return ""


def _date_notes(ws, row, kind, body, when, changes) -> list:
    """What an email's dates mean for this row, by email type."""
    extra = []
    if kind in DUE_DATE_KINDS:
        due = find_due(body, when)
        if due["date"]:
            extra.append(_due_date(ws, row, due["date"], changes))
        elif due["note"]:
            extra.append(due["note"])
    elif kind == "status_update":                     # a promise of delivery, not the DOI's due date
        expected = find_expected(body, when)
        if expected:
            extra.append(f"expects to provide by {expected:%Y-%m-%d}")
    return [e for e in extra if e]


def _advance_status(ws, row, kind, changes):
    """Acknowledgments, status updates, responses and questions mean work has started: move
    Not Started (or blank) to In Progress. Never touches any status a person chose."""
    cur = str(ws[f"{COL_STATUS}{row}"].value or "").strip()
    if kind in ADVANCES_TO_IN_PROGRESS and cur in ("", "Not Started"):
        ws[f"{COL_STATUS}{row}"] = "In Progress"
        changes.append(f"status In Progress ({LABEL[kind].lower()})")


def _append_note(ws, row, ref, line, changes) -> bool:
    notes = ws[f"{COL_NOTES}{row}"].value or ""
    if ref in notes:
        return False
    ws[f"{COL_NOTES}{row}"] = f"{notes}\n{line}" if notes else line
    changes.append("note")
    return True


def upsert(path, accepted, backup_dir=None, dry_run=False, last_sent=None, sent=None) -> list:
    """accepted:  [(Mail, request_id)] from fetch_candidates() - emails you received.
    sent:      [(SentMail, request_id)] from doi_filter.match_sent() - emails you sent.
    last_sent: {request_id: datetime or (datetime, source)} - from doi_filter.merge_last_sent()
               (accountants' emails, plus your Sent Items if TRACK_SENT).
    Returns one action dict per received email (added / updated / unchanged / skipped), plus one
    per sent email that changed a row, plus one per row whose Last Email Sent moved forward."""
    path = Path(path)
    check_unlocked(path)
    check_lossless(path)

    wb = load_workbook(path)
    ws = wb[SHEET]
    tbl, first, last = _table_bounds(ws)
    rows = _index(ws, first, last)

    actions = []
    for m, rid in sorted(accepted, key=lambda x: x[0].received):   # oldest first
        ref = note_ref(m.entry_id)
        an = m.analysis if m.analysis is not None else analyze(m, rid)
        kind = an.kind
        row = rows.get(rid)
        changes, review = [], list(an.review)

        if row is None:
            if kind not in CREATES_ROW:                 # e.g. a status update for an unknown request
                actions.append({"request_id": rid, "row": "-", "action": "skipped", "ref": ref, "mail": m,
                                "changes": f"{LABEL[kind].lower()} for a request not in the tracker - add it by hand"})
                continue
            row = _append_row(ws, tbl, first, last)
            last = rows[rid] = row
            text = request_text(m.subject, rid)
            if an.parent_request_id and "follow" not in text.lower():     # feeds 'Follows Up On'
                text = f"{text} - {an.parent_request_id} Follow-up"
            ws[f"{COL_REQUEST}{row}"] = text
            ws[f"{COL_STATUS}{row}"] = "Not Started"            # new rows only
            changes.append("new row")

        # Initial Request Received: keep the earliest date - a later reply is not the "initial" one.
        received_day = datetime.combine(m.received.date(), time())
        cur = ws[f"{COL_RECEIVED}{row}"].value
        if cur is None or (isinstance(cur, datetime) and received_day < cur):
            ws[f"{COL_RECEIVED}{row}"] = received_day
            changes.append("received date")

        # Contact: fill if blank; never replace what a person typed (except a cell holding only
        # your own address, which an earlier version wrote by mistake).
        current = ws[f"{COL_CONTACT}{row}"].value
        if not current or only_you(current):
            who = contact(m, kind, an.asked)
            if who:
                ws[f"{COL_CONTACT}{row}"] = who
                changes.append("contact" if not current else "contact (replaced your own address)")

        # Due Date: the tracking date (examiner's deadline; else the internal expected date as a
        # labelled fallback), only into an EMPTY cell. A different date is never chosen silently:
        # it is noted and flagged for review.
        extra = []
        if an.tracking_due_date:
            conflict = _due_date(ws, row, an.tracking_due_date, changes)
            if conflict:
                review.append(f"{conflict} ({an.date_source})")
            elif "due date" in changes and an.date_source.startswith("internal"):
                extra.append("Due Date set from the internal expected date - no examiner deadline seen yet")
        an.review = review                                           # incl. any conflict just found
        if "new row" in changes and an.send_to and not an.send_to_source.startswith("stated"):
            extra.append(f"send to {an.send_to} ({an.send_to_source.split(' (')[0]}, default)")
        if an.summary():
            extra.append(an.summary())
        _advance_status(ws, row, kind, changes)

        # Notes: append-only, one line per email, tagged so re-runs can tell it's done.
        _append_note(ws, row, ref, note_line(m, ref, "; ".join(extra), kind), changes)

        action = "added" if "new row" in changes else ("updated" if changes else "unchanged")
        actions.append({"request_id": rid, "row": row, "action": action, "review": "; ".join(review),
                        "changes": ", ".join(changes), "ref": ref, "mail": m})

    # Emails you sent in the team's templates: progress notes and status, never new rows.
    for s, rid in sorted(sent or [], key=lambda x: x[0].sent_on):
        kind, row = classify(s.subject, s.body), rows.get(rid)
        if kind == "other" or row is None:
            continue
        ref = note_ref(s.entry_id or f"{s.subject}|{s.sent_on:%Y-%m-%d %H:%M}")
        if ref in (ws[f"{COL_NOTES}{row}"].value or ""):
            continue                                       # already recorded on an earlier run
        changes = []
        extra = _date_notes(ws, row, kind, s.body, s.sent_on, changes)
        _advance_status(ws, row, kind, changes)
        _append_note(ws, row, ref, sent_note_line(s, ref, "; ".join(extra), kind), changes)
        actions.append({"request_id": rid, "row": row, "action": "updated", "ref": ref, "mail": None,
                        "changes": f"{LABEL[kind].lower()} sent: " + ", ".join(changes)})

    # Last Email Sent: your newest sent email about each request already in the tracker. Only
    # ever moves forward, so a date you typed that is newer than anything found is kept.
    for rid, value in sorted((last_sent or {}).items()):
        sent_on, source = value if isinstance(value, tuple) else (value, "")
        row = rows.get(rid)
        if row is None:                                   # an email never creates a request row here
            continue
        sent_day, cur = _day(sent_on), _day(ws[f"{COL_LAST_SENT}{row}"].value)
        if cur is None or sent_day > cur:
            ws[f"{COL_LAST_SENT}{row}"] = sent_day
            label = f"last email sent {sent_day:%Y-%m-%d}" + (f" ({source})" if source else "")
            actions.append({"request_id": rid, "row": row, "action": "updated",
                            "changes": label, "ref": "", "mail": None})

    if any(a["action"] not in ("unchanged", "skipped") for a in actions) and not dry_run:
        backup(path, backup_dir)                # nothing changed -> no backup, no save
        wb.save(path)
    return actions
