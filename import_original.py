"""One-time import: rebuild the original request list in the tracker's format.

    python import_original.py ORIGINAL.xlsx TRACKER_TEMPLATE.xlsx OUTPUT.xlsx

- ORIGINAL          the old list (sheet "Requests", one row per request, headers in row 1)
- TRACKER_TEMPLATE  DOI_Exam_Test_Tracker_v2.xlsx - supplies the layout, formulas, dropdowns
                    and Summary sheet; its example rows are replaced
- OUTPUT            a NEW file. Both inputs are left untouched.

Anything the script had to change or could not decide is printed as a review list and written
to OUTPUT_Import_Log.csv (source row -> tracker row), so every imported row can be traced
back to the original - the evidence trail for this one-off migration.
"""
import argparse
import csv
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from upsert_tracker import LAST_ROW, SHEET, TABLE, _table_bounds, check_lossless, first_word_id

SOURCE_SHEET = "Requests"
# original header (lower-case) -> tracker column
COLUMNS = {
    "assigned to": "B",
    "initial request email received": "C",
    "request name": "D",
    "due date": "G",
    "status": "I",
    "completion date": "K",
    "contact": "L",
    "last email sent": "M",
    "follow up": None,              # the follow-up link lives in the request text (D) instead
}
LEAD_ID = re.compile(r"^[A-Z]{2,}(?:-[A-Z]+)*-\d+(?=\s|$)")    # same "first word" rule as column A


def clean_request(text):
    """Make column D follow the pattern the tracker's formulas read:
    'ID Request - Subject - OTHERID Follow-up'. Returns (text, list of changes made)."""
    if text is None or not str(text).strip():
        return "", ["request name is blank"]
    original = str(text)
    s = original.replace("–", "-").replace("—", "-")          # en/em dash -> hyphen
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\b([A-Z]{2,}) (\d+)(?= Follow[ -]?up\b)", r"\1-\2", s, flags=re.I)  # 'ABC 27 Follow up'
    s = re.sub(r"\bFollow ?up\b", "Follow-up", s, flags=re.I)        # Type/Follows Up On look for 'Follow-up'
    s = re.sub(r"^([A-Z]{2,}(?:-[A-Z]+)*-\d+)[:;,.]\s*", r"\1 ", s)   # 'ABC-12: ...' -> 'ABC-12 ...'
    changes = []
    if s != original:
        changes.append(f"request name tidied (was: {original!r})")
    if not LEAD_ID.match(s):
        changes.append("no request ID at the start - fix column D by hand")
    return s, changes


def list_block(ws, col):
    """Values from row 4 down to the first blank cell. The Lists sheet has help text further
    down column A, so 'every non-empty cell' would treat help paragraphs as statuses."""
    out = []
    for r in range(4, 61):
        v = ws[f"{col}{r}"].value
        if v in (None, ""):
            break
        out.append(v)
    return out


def _header_map(ws):
    found = {}
    for cell in ws[1]:
        key = str(cell.value or "").strip().lower()
        if key in COLUMNS:
            found[key] = cell.column
    missing = [k for k, v in COLUMNS.items() if v and k not in found]
    if missing:
        raise SystemExit(f"Original is missing columns: {missing}")
    return found


def read_original(path):
    ws = load_workbook(path, data_only=True)[SOURCE_SHEET]
    cols = _header_map(ws)
    rows = []
    for r in range(2, ws.max_row + 1):
        rec = {k: ws.cell(r, c).value for k, c in cols.items()}
        if any(v not in (None, "") for v in rec.values()):
            rec["_source_row"] = r
            rows.append(rec)
    return rows


def convert(rec, statuses):
    """One original row -> (tracker values by column letter, notes text, review items)."""
    out, notes, review = {}, [], []
    out["D"], review_d = clean_request(rec.get("request name"))
    review += review_d
    for key in ("assigned to", "contact"):
        v = rec.get(key)
        out[COLUMNS[key]] = str(v).strip() if v not in (None, "") else None
    for key in ("initial request email received", "completion date", "last email sent", "due date"):
        v = rec.get(key)
        if isinstance(v, datetime) or v in (None, ""):
            out[COLUMNS[key]] = v or None
        else:                                                   # e.g. 'TBD' typed in a date column
            out[COLUMNS[key]] = None
            notes.append(f"{key.title()} in original list: {v}")
            review.append(f"{key} was text ({v!r}) - moved to Notes")
    status = str(rec.get("status") or "").strip()
    canonical = {s.lower(): s for s in statuses}
    if status.lower() in canonical:
        out["I"] = canonical[status.lower()]
    elif status:                                                # free-text update in the Status column
        out["I"] = None
        notes.append(status)
        review.append("status was free text - moved to Notes; pick a Status from the dropdown")
    else:
        out["I"] = None
        review.append("no status - pick one from the dropdown")
    if rec.get("follow up") not in (None, ""):
        notes.append(f"Follow Up (original column): {rec['follow up']}")
    out["O"] = "\n".join(notes) or None
    return out, review


def build(original, template, output):
    check_lossless(Path(template))
    rows = read_original(original)
    wb = load_workbook(template)
    ws = wb[SHEET]
    lists = wb["Lists"]
    statuses = list_block(lists, "A")
    tbl, first, old_last = _table_bounds(ws)
    if first + len(rows) - 1 > LAST_ROW:
        raise SystemExit(f"{len(rows)} rows won't fit: named ranges stop at row {LAST_ROW}.")

    # The template's first data row is the pattern for every row: formatting + formulas.
    pattern = {c.column_letter: (copy(c._style), c.value) for c in ws[first]}
    height = ws.row_dimensions[first].height
    blank_style = {c.column_letter: copy(c._style) for c in ws[old_last + 1]}

    log, seen_ids = [], {}
    for i, rec in enumerate(rows):
        r = first + i
        for col, (style, value) in pattern.items():
            cell = ws[f"{col}{r}"]
            cell._style = copy(style)
            is_formula = isinstance(value, str) and value.startswith("=")
            cell.value = Translator(value, origin=f"{col}{first}").translate_formula(f"{col}{r}") if is_formula else None
        ws.row_dimensions[r].height = height
        values, review = convert(rec, statuses)
        for col, v in values.items():
            ws[f"{col}{r}"] = v
        rid = first_word_id(values["D"])
        if rid and rid in seen_ids:
            review.append(f"duplicate request ID - also on tracker row {seen_ids[rid]}")
        seen_ids.setdefault(rid, r)
        log.append({"source_row": rec["_source_row"], "tracker_row": r, "request_id": rid,
                    "review": "; ".join(review)})

    # Clear leftover example rows if the original is shorter than the template's sample data.
    new_last = first + len(rows) - 1
    for r in range(new_last + 1, old_last + 1):
        for col, style in blank_style.items():
            ws[f"{col}{r}"].value = None
            ws[f"{col}{r}"]._style = copy(style)
        ws.row_dimensions[r].height = None
    end_col = re.match(r"[A-Z]+", tbl.ref.split(":")[1]).group(0)
    tbl.ref = f"{tbl.ref.split(':')[0]}:{end_col}{max(new_last, first)}"
    if tbl.autoFilter is not None:
        tbl.autoFilter.ref = tbl.ref

    # Every assignee needs to be on the Lists sheet or the dropdown marks the cell invalid.
    assignees = list_block(lists, "D")
    for name in sorted({ws[f"B{r}"].value for r in range(first, new_last + 1)} - set(assignees) - {None}):
        lists.cell(4 + len(assignees), 4, name)
        assignees.append(name)

    wb.calculation.fullCalcOnLoad = True
    wb.save(output)
    log_path = Path(output).with_name(Path(output).stem + "_Import_Log.csv")
    with open(log_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["source_row", "tracker_row", "request_id", "review"])
        w.writeheader()
        w.writerows(log)
    return log, log_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("original")
    ap.add_argument("template")
    ap.add_argument("output")
    args = ap.parse_args()
    out = Path(args.output)
    if out.exists():
        sys.exit(f"{out} already exists - pick a new name (this script never overwrites).")
    if out.resolve() in (Path(args.original).resolve(), Path(args.template).resolve()):
        sys.exit("OUTPUT must be a new file, not one of the inputs.")

    log, log_path = build(args.original, args.template, out)
    flagged = [x for x in log if x["review"]]
    print(f"Imported {len(log)} requests into {out}")
    print(f"Import log: {log_path}")
    if flagged:
        print(f"\nCheck these {len(flagged)} rows by hand:")
        for x in flagged:
            print(f"  tracker row {x['tracker_row']} (original row {x['source_row']}) {x['request_id']}: {x['review']}")


if __name__ == "__main__":
    main()
