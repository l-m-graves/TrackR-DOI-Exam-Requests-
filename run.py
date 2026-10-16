"""Entry point: fetch from Outlook -> detect (accept / review / reject) -> analyze -> upsert
tracker -> publish copy -> log. Read-only towards Outlook: nothing is sent, moved or deleted.

    python run.py                 # last 14 days, writes the tracker
    python run.py --days 30       # look further back
    python run.py --dry-run       # show what would change, write/save nothing
    python run.py --all           # report every email again, not only new ones
    python run.py --scheduled     # how the hourly Task Scheduler job runs it (see schedule_hourly.py)

Each run remembers which emails it has handled (DOI_Seen_Messages.csv next to the log). If
nothing is new it stops straight away - no tracker save, no log lines, no attachment copies.
"""
import argparse
import csv
import sys
from datetime import datetime, timedelta
from pathlib import Path

import doi_contacts
import doi_filter
import seen_state
from publish import needs_publish, publish
from doi_analysis import analyze, examiner_included
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


def seen_path() -> Path:
    """The memory of handled emails lives next to the run log (a Title_Case saved file)."""
    return Path(doi_filter.LOG_PATH).with_name("DOI_Seen_Messages.csv")


def errors_path() -> Path:
    return Path(doi_filter.LOG_PATH).with_name("DOI_Run_Errors.txt")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--days", type=int, default=14, help="how many days back to look (default 14)")
    ap.add_argument("--dry-run", action="store_true", help="don't save the tracker, publish, or write the log")
    ap.add_argument("--all", action="store_true", help="ignore the memory of handled emails: report everything")
    ap.add_argument("--scheduled", action="store_true",
                    help="unattended hourly run (Task Scheduler): any crash goes to DOI_Run_Errors.txt")
    args = ap.parse_args(argv)
    if not args.scheduled:
        return _run(args)
    try:
        return _run(args)
    except Exception:                               # no console to show it on - keep it in a file
        import traceback
        path = errors_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"--- {datetime.now():%Y-%m-%d %H:%M:%S}\n{traceback.format_exc()}\n")
        return 1


def _run(args):
    run_at = datetime.now()
    since = run_at - timedelta(days=args.days)
    state = {} if args.all else seen_state.load(seen_path())
    contacts = doi_contacts.load(doi_filter.CONTACTS_PATH) if doi_filter.CONTACTS_PATH else None
    approved = doi_contacts.approved_senders(contacts) if contacts else set()
    if contacts:                        # people the examiner side brought onto threads earlier
        doi_filter.THREAD_THIRD_PARTIES = doi_contacts.thread_third_parties(contacts)
    attach_dir = None if args.dry_run else doi_filter.ATTACH_DIR     # dry run saves nothing
    accepted, rejected, review = doi_filter.fetch_candidates(since=since, attach_dir=attach_dir,
                                                             approved_senders=approved,
                                                             skip_refs=seen_state.done_refs(state))

    # Is anything NEW? An email is news if never judged before, or its outcome changed (e.g. a
    # review item whose sender you have since approved). Already-handled emails are not.
    def ref(m):
        return note_ref(m.entry_id)
    news = {ref(m) for m, _ in accepted if seen_state.is_news(state, ref(m), "accept")}
    news |= {ref(m) for m, _ in review if seen_state.is_news(state, ref(m), "review")}
    news |= {ref(m) for m, _ in rejected if seen_state.is_news(state, ref(m), "reject")}
    if not news:
        print("No new DOI-related emails - nothing to do.")
        # The one exception: edits you typed into the tracker still reach the coworkers' copy.
        if not args.dry_run and doi_filter.PUBLISH_PATH and Path(doi_filter.TRACKER_PATH).exists() \
                and needs_publish(Path(doi_filter.TRACKER_PATH), Path(doi_filter.PUBLISH_PATH)):
            result = publish(doi_filter.TRACKER_PATH, doi_filter.PUBLISH_PATH)
            print(result)
            write_log(doi_filter.LOG_PATH, [run_row(run_at, "publish", f"tracker edited by hand: {result}")])
        return 0

    for m, rid in accepted:
        m.analysis = analyze(m, rid, approved)
    print(f"{len(news)} new or changed: {len(accepted)} accepted, {len(review)} need review, "
          f"{len(rejected)} rejected (unchanged ones are not repeated)")

    log = [log_row(run_at, m, "rejected", reason) for m, reason in rejected if ref(m) in news]
    log += [log_row(run_at, m, "review", review_reason=det.reasons[0]) for m, det in review if ref(m) in news]
    # Last Email Sent (column M): the newest email in each request's exchange, whoever sent it -
    # every fetched message, not only accepted ones - plus, if TRACK_SENT, your own Sent Items.
    fetched = [m for m, _ in accepted] + [m for m, _ in review] + [m for m, _ in rejected]
    exchange = doi_filter.last_exchange_email(fetched, accepted, seen_state.known_threads(state))
    own, sent = {}, []
    if doi_filter.TRACK_SENT:
        try:
            sent = doi_filter.match_sent(doi_filter.fetch_sent(since=since), accepted)
            own = doi_filter.last_sent_by_request([x for x, _ in sent], accepted)
        except Exception as e:                      # never let Sent Items stop the main update
            print(f"Sent Items not read: {e}")
            log.append(run_row(run_at, "error", f"Sent Items not read: {e}"))
    last_sent = doi_filter.merge_last_sent(exchange, own)
    written = True
    try:
        actions = upsert(doi_filter.TRACKER_PATH, accepted, dry_run=args.dry_run,
                         last_sent=last_sent, sent=sent)
    except TrackerError as e:
        print(f"NOT WRITTEN: {e} (these emails will be tried again next run)")
        log += [log_row(run_at, m, "error", str(e)) for m, _ in accepted]
        actions, written = [], False
    for a in actions:
        if a["mail"] is not None and ref(a["mail"]) not in news and a["action"] in ("unchanged", "skipped"):
            continue                                # nothing new to say about this email
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
        if ref(m) in news:
            print(f"  REVIEW    {m.subject[:60]!r}: {det.reasons[0]}")
    for m, reason in rejected:
        if ref(m) in news:
            print(f"  rejected  {m.subject[:60]!r}: {reason}")

    # Contact directory: record everyone on request emails (accepted, or rejected after the ID
    # check passed). New senders stay untrusted until you type Y in Approved.
    if contacts is not None:
        new, before = [], doi_contacts.thread_third_parties(contacts)
        for m, why in accepted + rejected + review:
            rid = doi_filter.find_id(m.subject) or getattr(m.detection, "request_id", "")
            if rid and ref(m) in news:
                included = examiner_included(m, approved) if doi_filter.THIRD_PARTY_FROM_THREAD else set()
                new += doi_contacts.observe(contacts, m, rid, m.received or run_at, included)
        joined = sorted(doi_contacts.thread_third_parties(contacts) - before)
        if joined:                      # takes effect from the next run (review items are re-judged then)
            print(f"Contacts: {len(joined)} person(s) now count as third party - the examiner side put them on "
                  f"a request email (column Examiner_Thread; type N in Approved to undo)")
            log.append(run_row(run_at, "contacts", f"third party via examiner thread: {', '.join(joined)}"))
        pending = doi_contacts.pending_senders(contacts)
        print(f"Contacts: {len(new)} new, {len(pending)} sender(s) awaiting your approval"
              + (f" - open {Path(doi_filter.CONTACTS_PATH).name} and type Y in Approved" if pending else ""))
        if not args.dry_run:
            result = doi_contacts.save(doi_filter.CONTACTS_PATH, contacts)
            log.append(run_row(run_at, "contacts", f"{len(new)} new: {', '.join(new)}; {result}" if new else result))

    if args.dry_run:
        print("Dry run - tracker, shared copy, contacts, memory and log not written.")
        return 0

    # Remember what was handled - but only once the tracker was saved. A run that found the
    # tracker open in Excel remembers nothing, so the next run does the whole job again
    # (including Last Email Sent dates that came from review/rejected emails in the thread).
    if written:
        outcome = {a["ref"]: a["action"] for a in actions if a.get("mail") is not None}
        for m, rid in accepted:
            seen_state.record(state, ref(m), "skipped" if outcome.get(ref(m)) == "skipped" else "done",
                              rid, m.conversation_id, run_at)
        for m, det in review:
            seen_state.record(state, ref(m), "review", det.request_id, m.conversation_id, run_at)
        for m, _ in rejected:
            seen_state.record(state, ref(m), "reject", getattr(m.detection, "request_id", "") or "",
                              m.conversation_id, run_at)
        seen_state.prune(state, keep_days=args.days + 30, now=run_at)
        result = seen_state.save(seen_path(), state)
        if result != "saved":
            print(result)
            log.append(run_row(run_at, "error", result))

    # Publish even if the upsert was refused (e.g. tracker open in Excel): the copy is made
    # from your last saved version, and it also carries edits you typed yourself.
    if doi_filter.PUBLISH_PATH:
        result = publish(doi_filter.TRACKER_PATH, doi_filter.PUBLISH_PATH)
        print(result)
        log.append(run_row(run_at, "publish", result))
    write_log(doi_filter.LOG_PATH, log)
    return 0


if __name__ == "__main__":
    sys.exit(main())
