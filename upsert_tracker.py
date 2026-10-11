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

from doi_filter import REQUEST_ID

SHEET = "DOI Exam Tracker"
TABLE = "tblRequests"
LAST_ROW = 1000         # the trk* named ranges and dropdowns stop at row 1000

# Columns the script fills (letters match the sheet). Everything else is either
# a formula (A, E, F, H, J, N) or left for a person to type (B, G, K, M).
COL_RECEIVED, COL_REQUEST, COL_STATUS, COL_CONTACT, COL_NOTES = "C", "D", "I", "L", "O"

# File parts openpyxl cannot carry through a save. If the tracker ever gains
# one of these (e.g. a chart on Summary), stop instead of silently deleting it.
LOSSY_PARTS = ("xl/charts/", "xl/media/", "xl/slicers/", "xl/slicerCaches/", "xl/timelines/")
LOSSY_DRAWING = re.compile(r"^xl/drawings/drawing\d+\.xml$")     # comment .vml files are fine

REPLY_PREFIX = re.compile(r"^\s*((re|fw|fwd)\s*:\s*)+", re.IGNORECASE)


class TrackerError(Exception):
    """Something about the workbook means we must not write to it."""


# --- small helpers (pure, unit-tested) ---------------------------------------
def note_ref(entry_id: str) -> str:
    """Short, stable tag for one email. EntryIDs are ~140 hex chars; 8 chars of a hash is plenty."""
    return "ref " + hashlib.sha1(entry_id.encode()).hexdigest()[:8]


def request_text(subject: str, rid: str) -> str:
    """Column D text. The sheet's formulas take the Request ID as D's first word, so drop
    RE:/FW: prefixes and, if the ID still isn't the first word, put it in front."""
    text = REPLY_PREFIX.sub("", subject).strip()
    if text.split(" ", 1)[0] != rid:
        text = f"{rid} {text}"
    return text


def note_line(m, ref: str) -> str:
    return f"{m.received:%Y-%m-%d} email from {m.sender}: {m.subject} [{ref}]"


def contact(m) -> str:
    return f"{m.sender_name} <{m.sender}>" if m.sender_name else m.sender


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
    backup_dir = Path(backup_dir or path.parent / "backups")
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
    """Request ID -> row, using the same rule as column A's formula (first word of D)."""
    rows = {}
    for r in range(first, last + 1):
        text = ws[f"{COL_REQUEST}{r}"].value
        if text:
            rows.setdefault(str(text).split(" ", 1)[0], r)
    return rows


def _append_row(ws, tbl, last) -> int:
    """Add a row under the table: copy formulas (shifted to the new row) and formatting
    from the row above, then stretch the table to include it."""
    new = last + 1
    if new > LAST_ROW:
        raise TrackerError(f"Tracker is full: named ranges stop at row {LAST_ROW}.")
    if any(ws.cell(new, c).value is not None for c in range(1, ws.max_column + 1)):
        raise TrackerError(f"Row {new} under the table is not empty - won't overwrite it.")
    for src in ws[last]:
        dst = ws.cell(new, src.column)
        dst._style = copy(src._style)
        if isinstance(src.value, str) and src.value.startswith("="):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
    ws.row_dimensions[new].height = ws.row_dimensions[last].height
    start = tbl.ref.split(":")[0]
    end_col = re.match(r"[A-Z]+", tbl.ref.split(":")[1]).group(0)
    tbl.ref = f"{start}:{end_col}{new}"
    if tbl.autoFilter is not None:
        tbl.autoFilter.ref = tbl.ref
    return new


def upsert(path, accepted, backup_dir=None, dry_run=False) -> list:
    """accepted: list of (Mail, request_id) from fetch_candidates().
    Returns one action dict per email: added / updated / unchanged."""
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
        row = rows.get(rid)
        changes = []

        if row is None:
            row = _append_row(ws, tbl, last)
            last = rows[rid] = row
            ws[f"{COL_REQUEST}{row}"] = request_text(m.subject, rid)
            ws[f"{COL_STATUS}{row}"] = "Not Started"            # new rows only
            changes.append("new row")

        # Initial Request Received: keep the earliest date - a later reply is not the "initial" one.
        received_day = datetime.combine(m.received.date(), time())
        cur = ws[f"{COL_RECEIVED}{row}"].value
        if cur is None or (isinstance(cur, datetime) and received_day < cur):
            ws[f"{COL_RECEIVED}{row}"] = received_day
            changes.append("received date")

        # Contact: fill if blank; never replace what a person typed.
        if not ws[f"{COL_CONTACT}{row}"].value:
            ws[f"{COL_CONTACT}{row}"] = contact(m)
            changes.append("contact")

        # Notes: append-only, one line per email, tagged so re-runs can tell it's done.
        notes = ws[f"{COL_NOTES}{row}"].value or ""
        if ref not in notes:
            line = note_line(m, ref)
            ws[f"{COL_NOTES}{row}"] = f"{notes}\n{line}" if notes else line
            changes.append("note")

        kind = "added" if "new row" in changes else ("updated" if changes else "unchanged")
        actions.append({"request_id": rid, "row": row, "action": kind,
                        "changes": ", ".join(changes), "ref": ref, "mail": m})

    if any(a["action"] != "unchanged" for a in actions) and not dry_run:
        backup(path, backup_dir)                # nothing changed -> no backup, no save
        wb.save(path)
    return actions
