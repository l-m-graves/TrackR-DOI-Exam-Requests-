"""Entry point: fetch from Outlook -> detect (accept / review / reject) -> analyze -> upsert
tracker -> publish copy -> log. Read-only towards Outlook: nothing is sent, moved or deleted.

    python run.py                 # last 14 days, writes the tracker
    python run.py --days 30       # look further back
    python run.py --dry-run       # show what would change, write/save nothing
"""
import argparse
import csv
from datetime import datetime, timedelta
from pathlib import Path

import doi_contacts
import doi_filter
from publish import publish
from doi_analysis import analyze
from upsert_tracker import TrackerError, note_ref, upsert

LOG_FIELDS = ["run_at", "result", "request_id", "topic", "classification", "confidence", "signals",
              "message_id", "conversation_id", "current_sender", "original_sender", "received", "ref",
              "subject", "exam_due_date", "internal_expected_date", "tracking_due_date", "date_source",
              "date_evidence", "attachments", "xlsx_update", "detail", "review_reason"]


def write_log(path, rows):
    """Append to a CSV audit log: every email looked at, why it was accepted, sent to review or
    rejected, and what happened to the tracker. Creates the file (and folder) on first use. If the
    log is open in Excel (locked), writes a dated file next to it instead of losing the entries.
    A log written by an older version (different columns) is renamed, never mixed."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        with open(path, newline="", encoding="utf-8") as f:
            header = next(csv.reader(f), [])
        if header and header != LOG_FIELDS:
            old = path.with_name(f"{path.stem}_Old_Format_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
            try:
                path.rename(old)
                print(f"Log columns changed - previous log kept as {old.name}.")
            except PermissionError:
                pass                                   # locked: the fallback below handles it
    try:
        _append_csv(path, rows)
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        _append_csv(fallback, rows)
        print(f"Log {path.name} is open/locked - this run was logged to {fallback.name} instead.")
        return fallback


def _append_csv(path, rows):
    new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if new:
            w.writeheader()
        w.writerows(rows)


def _date(d):
    return f"{d:%Y-%m-%d}" if d else ""


def log_row(run_at, m, result, detail="", xlsx_update="", review_reason=""):
    """One log line about one received email, with its detection and analysis when available."""
    det, an = m.detection, m.analysis
    row = {"run_at": f"{run_at:%Y-%m-%d %H:%M:%S}", "result": result, "detail": detail,
           "message_id": m.entry_id, "conversation_id": m.conversation_id, "current_sender": m.sender,
           "received": f"{m.received:%Y-%m-%d %H:%M}" if m.received else "", "ref": note_ref(m.entry_id),
           "subject": m.subject, "attachments": "; ".join(m.attachment_status),
           "xlsx_update": xlsx_update, "review_reason": review_reason}
    if det is not None:
        row.update(request_id=det.request_id, confidence=det.confidence, signals="; ".join(det.reasons),
                   original_sender=det.original_sender)
    if an is not None:
        row.update(request_id=an.request_id, topic=an.subject_area, classification=f"{an.request_type}/{an.kind}",
                   original_sender=an.original_sender, exam_due_date=_date(an.exam_due_date),
                   internal_expected_date=_date(an.internal_expected_date),
                   tracking_due_date=_date(an.tracking_due_date), date_source=an.date_source,
                   date_evidence=an.date_evidence)
    return row


def run_row(run_at, result, detail, **extra):
    """A log line that isn't about one received email (publish, contacts, Sent Items...)."""
    return {"run_at": f"{run_at:%Y-%m-%d %H:%M:%S}", "result": result, "detail": detail, **extra}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--days", type=int, default=14, help="how many days back to look (default 14)")
    ap.add_argument("--dry-run", action="store_true", help="don't save the tracker, publish, or write the log")
    args = ap.parse_args()

    run_at = datetime.now()
    contacts = doi_contacts.load(doi_filter.CONTACTS_PATH) if doi_filter.CONTACTS_PATH else None
    approved = doi_contacts.approved_senders(contacts) if contacts else set()
    attach_dir = None if args.dry_run else doi_filter.ATTACH_DIR     # dry run saves nothing
    accepted, rejected, review = doi_filter.fetch_candidates(since=run_at - timedelta(days=args.days),
                                                             attach_dir=attach_dir, approved_senders=approved)
    for m, rid in accepted:
        m.analysis = analyze(m, rid, approved)
    print(f"{len(accepted)} accepted, {len(review)} need review, {len(rejected)} rejected")

    log = [log_row(run_at, m, "rejected", reason) for m, reason in rejected]
    log += [log_row(run_at, m, "review", review_reason=det.reasons[0]) for m, det in review]
    last_sent, sent = {}, []
    if doi_filter.TRACK_SENT:                       # Sent Items -> Last Email Sent (column M)
        try:
            sent = doi_filter.match_sent(doi_filter.fetch_sent(since=run_at - timedelta(days=args.days)), accepted)
            last_sent = doi_filter.last_sent_by_request([x for x, _ in sent], accepted)
        except Exception as e:                      # never let Sent Items stop the main update
            print(f"Sent Items not read: {e}")
            log.append(run_row(run_at, "error", f"Sent Items not read: {e}"))
    try:
        actions = upsert(doi_filter.TRACKER_PATH, accepted, dry_run=args.dry_run,
                         last_sent=last_sent, sent=sent)
    except TrackerError as e:
        print(f"NOT WRITTEN: {e}")
        log += [log_row(run_at, m, "error", str(e)) for m, _ in accepted]
        actions = []
    for a in actions:
        print(f"  {a['action']:9} {a['request_id']:<16} row {str(a['row']):<4} {a['changes']}")
        detail = f"row {a['row']}: {a['changes']}"
        if a.get("review"):
            print(f"            review: {a['review']}")
        log.append(log_row(run_at, a["mail"], a["action"], detail, xlsx_update=a["action"],
                           review_reason=a.get("review", "")) if a["mail"]
                   else run_row(run_at, a["action"], f"{a['request_id']} {detail}",
                                request_id=a["request_id"], xlsx_update=a["action"]))
    for m, rid in accepted:                         # attachments are evidence - log each one
        for result in m.saved:
            print(f"  attachment {rid:<15} {result}")
            log.append(log_row(run_at, m, "attachment", result))
    for m, det in review:
        print(f"  REVIEW    {m.subject[:60]!r}: {det.reasons[0]}")
    for m, reason in rejected:
        print(f"  rejected  {m.subject[:60]!r}: {reason}")

    # Contact directory: record everyone on request emails (accepted, or rejected after the ID
    # check passed). New senders stay untrusted until you type Y in Approved.
    if contacts is not None:
        new = []
        for m, why in accepted + rejected + review:
            rid = doi_filter.REQUEST_ID.search(m.subject)
            if rid:
                new += doi_contacts.observe(contacts, m, rid.group(0), m.received or run_at)
        pending = doi_contacts.pending_senders(contacts)
        print(f"Contacts: {len(new)} new, {len(pending)} sender(s) awaiting your approval"
              + (f" - open {Path(doi_filter.CONTACTS_PATH).name} and type Y in Approved" if pending else ""))
        if not args.dry_run:
            result = doi_contacts.save(doi_filter.CONTACTS_PATH, contacts)
            log.append(run_row(run_at, "contacts", f"{len(new)} new: {', '.join(new)}; {result}" if new else result))

    if args.dry_run:
        print("Dry run - tracker, shared copy, contacts and log not written.")
        return
    # Publish even if the upsert was refused (e.g. tracker open in Excel): the copy is made
    # from your last saved version, and it also carries edits you typed yourself.
    if doi_filter.PUBLISH_PATH:
        result = publish(doi_filter.TRACKER_PATH, doi_filter.PUBLISH_PATH)
        print(result)
        log.append(run_row(run_at, "publish", result))
    write_log(doi_filter.LOG_PATH, log)


if __name__ == "__main__":
    main()
