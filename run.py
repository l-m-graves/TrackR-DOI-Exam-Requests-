"""Entry point: fetch from Outlook -> qualify -> upsert tracker -> publish copy -> log.

    python run.py                 # last 14 days, writes the tracker
    python run.py --days 30       # look further back
    python run.py --dry-run       # show what would change, write nothing
"""
import argparse
import csv
from datetime import datetime, timedelta
from pathlib import Path

import doi_filter
from publish import publish
from upsert_tracker import TrackerError, note_ref, upsert

LOG_FIELDS = ["run_at", "received", "ref", "sender", "subject", "result", "detail"]


def write_log(path, rows):
    """Append to a CSV audit log: every email looked at, why it was accepted or rejected,
    and what happened to the tracker."""
    path = Path(path)
    new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if new:
            w.writeheader()
        w.writerows(rows)


def log_row(run_at, m, result, detail):
    return {"run_at": f"{run_at:%Y-%m-%d %H:%M:%S}", "received": f"{m.received:%Y-%m-%d %H:%M}",
            "ref": note_ref(m.entry_id), "sender": m.sender, "subject": m.subject,
            "result": result, "detail": detail}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--days", type=int, default=14, help="how many days back to look (default 14)")
    ap.add_argument("--dry-run", action="store_true", help="don't save the tracker, publish, or write the log")
    args = ap.parse_args()

    run_at = datetime.now()
    accepted, rejected = doi_filter.fetch_candidates(since=run_at - timedelta(days=args.days))
    print(f"{len(accepted)} qualified, {len(rejected)} rejected")

    log = [log_row(run_at, m, "rejected", reason) for m, reason in rejected]
    try:
        actions = upsert(doi_filter.TRACKER_PATH, accepted, dry_run=args.dry_run)
    except TrackerError as e:
        print(f"NOT WRITTEN: {e}")
        log += [log_row(run_at, m, "error", str(e)) for m, _ in accepted]
        actions = []
    for a in actions:
        print(f"  {a['action']:9} {a['request_id']:<16} row {a['row']:<4} {a['changes']}")
        log.append(log_row(run_at, a["mail"], a["action"], f"row {a['row']}: {a['changes']}"))
    for m, reason in rejected:
        print(f"  rejected  {m.subject[:60]!r}: {reason}")

    if args.dry_run:
        print("Dry run - tracker, shared copy and log not written.")
        return
    # Publish even if the upsert was refused (e.g. tracker open in Excel): the copy is made
    # from your last saved version, and it also carries edits you typed yourself.
    if doi_filter.PUBLISH_PATH:
        result = publish(doi_filter.TRACKER_PATH, doi_filter.PUBLISH_PATH)
        print(result)
        log.append({"run_at": f"{run_at:%Y-%m-%d %H:%M:%S}", "result": "publish", "detail": result})
    write_log(doi_filter.LOG_PATH, log)


if __name__ == "__main__":
    main()
