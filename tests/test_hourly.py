"""The hourly run: whole-Inbox scan, the memory of handled emails, "nothing new -> do nothing",
and the Task Scheduler definition. Outlook is faked with plain objects shaped like its COM items."""
import io
import os
import shutil
import tempfile
import unittest
import xml.etree.ElementTree as ET
from contextlib import redirect_stdout
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import doi_filter
import run
import schedule_hourly
import seen_state
from tests import fake_doi_emails as E
from upsert_tracker import TrackerError

NOW = datetime.now().replace(microsecond=0)


# --- a fake Outlook --------------------------------------------------------
def person(addr, kind=None):
    return SimpleNamespace(Address=addr, Type=kind, Name="",
                           PropertyAccessor=SimpleNamespace(GetProperty=lambda prop: addr))


def fake_msg(subject, body, sender, entry_id, conversation="", received=None):
    return SimpleNamespace(Class=43, Subject=subject, Body=body, Sender=person(sender), SenderEmailAddress=sender,
                           SenderName="", Recipients=[person("ana.accountant@corp.test", 1), person(E.INTERN, 2)],
                           Attachments=[], ReceivedTime=received or NOW - timedelta(hours=2),
                           EntryID=entry_id, ConversationID=conversation)


class FakeItems(list):
    def Restrict(self, _filter):            # Outlook filters; the fake hands everything back
        return self


class FakeFolder:
    def __init__(self, name, messages=(), children=(), broken=False):
        self.Name, self._messages, self.Folders, self.broken = name, list(messages), list(children), broken

    @property
    def Items(self):
        if self.broken:
            raise RuntimeError("folder not available offline")
        return FakeItems(self._messages)


def fake_outlook(inbox):
    return SimpleNamespace(GetDefaultFolder=lambda n: inbox)


REQUEST = ("FIN-301 Request - Premium Tax",
           "Please provide the DOI examination support for premium tax by October 30, 2027.",
           "lead.examiner@doi-state.test")
UNKNOWN = ("FIN-302 Request - Agency Review",
           "Please provide the DOI examination items for the agency review.", "someone@elsewhere.test")


class RolesMixin:
    def setUp(self):
        patcher = mock.patch.multiple(doi_filter, **E.ROLES)
        patcher.start()
        self.addCleanup(patcher.stop)


# --- whole-Inbox scan ------------------------------------------------------
class InboxScanTests(RolesMixin, unittest.TestCase):
    def fetch(self, inbox, **kw):
        with mock.patch.object(doi_filter, "_outlook", return_value=fake_outlook(inbox)):
            return doi_filter.fetch_candidates(NOW - timedelta(days=14), **kw)

    def test_subfolders_at_any_depth_are_scanned(self):
        deep = FakeFolder("2027", [fake_msg(*REQUEST, "E1")])
        inbox = FakeFolder("Inbox", children=[FakeFolder("DOI Exam", children=[deep])])
        accepted, rejected, review = self.fetch(inbox)
        self.assertEqual([rid for _, rid in accepted], ["FIN-301"])

    def test_same_email_in_two_folders_is_judged_once(self):
        msg = fake_msg(*REQUEST, "E1")
        inbox = FakeFolder("Inbox", [msg], children=[FakeFolder("DOI Exam", [msg])])
        accepted, _, _ = self.fetch(inbox)
        self.assertEqual(len(accepted), 1)

    def test_unreadable_folder_and_non_mail_items_are_skipped(self):
        meeting = SimpleNamespace(Class=53, EntryID="M1")                 # a meeting request, not mail
        inbox = FakeFolder("Inbox", [meeting, fake_msg(*REQUEST, "E1")],
                           children=[FakeFolder("Archive", broken=True)])
        accepted, rejected, review = self.fetch(inbox)
        self.assertEqual((len(accepted), len(rejected), len(review)), (1, 0, 0))

    def test_no_doi_exam_folder_needed(self):
        accepted, _, _ = self.fetch(FakeFolder("Inbox", [fake_msg(*REQUEST, "E1")]))
        self.assertEqual(len(accepted), 1)

    def test_done_emails_are_not_opened_again(self):
        inbox = FakeFolder("Inbox", [fake_msg(*REQUEST, "E1")])
        accepted, _, _ = self.fetch(inbox, skip_refs={doi_filter.message_ref("E1")})
        self.assertEqual(accepted, [])


# --- the memory ------------------------------------------------------------
class SeenStateTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir)

    def test_news_rules(self):
        state = {}
        self.assertTrue(seen_state.is_news(state, "ref 1", "accept"))           # never seen
        seen_state.record(state, "ref 1", "done", "FIN-1", "T1")
        self.assertFalse(seen_state.is_news(state, "ref 1", "accept"))          # already written
        seen_state.record(state, "ref 2", "review", "FIN-2")
        self.assertFalse(seen_state.is_news(state, "ref 2", "review"))          # still waiting: quiet
        self.assertTrue(seen_state.is_news(state, "ref 2", "accept"))           # sender since approved
        seen_state.record(state, "ref 3", "reject")
        self.assertFalse(seen_state.is_news(state, "ref 3", "reject"))
        self.assertEqual(seen_state.done_refs(state), {"ref 1"})
        self.assertEqual(seen_state.known_threads(state), {"T1": "FIN-1"})

    def test_save_load_round_trip_and_first_seen_kept(self):
        path = self.dir / "DOI_Seen_Messages.csv"
        state = {}
        seen_state.record(state, "ref 1", "review", "FIN-1", when=datetime(2027, 10, 1, 9))
        seen_state.record(state, "ref 1", "done", when=datetime(2027, 10, 2, 9))
        self.assertEqual(seen_state.save(path, state), "saved")
        row = seen_state.load(path)["ref 1"]
        self.assertEqual((row["result"], row["request_id"], row["first_seen"], row["last_seen"]),
                         ("done", "FIN-1", "2027-10-01 09:00", "2027-10-02 09:00"))

    def test_missing_or_broken_file_means_start_fresh(self):
        self.assertEqual(seen_state.load(self.dir / "nope.csv"), {})
        broken = self.dir / "broken.csv"
        broken.write_text("something,else\n1,2\n", encoding="utf-8")
        self.assertEqual(seen_state.load(broken), {})

    def test_prune_forgets_old_emails(self):
        state = {}
        seen_state.record(state, "ref old", "done", when=datetime(2027, 1, 1))
        seen_state.record(state, "ref new", "done", when=datetime(2027, 10, 1))
        seen_state.prune(state, keep_days=44, now=datetime(2027, 10, 2))
        self.assertEqual(list(state), ["ref new"])

    def test_locked_file_is_reported_not_raised(self):
        with mock.patch("builtins.open", side_effect=PermissionError("open in Excel")):
            self.assertIn("NOT saved", seen_state.save(self.dir / "DOI_Seen_Messages.csv", {}))


# --- run.py hourly behaviour -----------------------------------------------
def fake_upsert(path, accepted, **kw):
    return [{"request_id": rid, "row": 4, "action": "added", "ref": doi_filter.message_ref(m.entry_id),
             "mail": m, "changes": "new row"} for m, rid in accepted]


class HourlyRunTests(RolesMixin, unittest.TestCase):
    def setUp(self):
        super().setUp()
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir)
        self.log = self.dir / "Logs" / "DOI_Run_Log.csv"
        self.inbox = FakeFolder("Inbox", [fake_msg(*REQUEST, "E1", "T1")])
        for name, value in dict(LOG_PATH=str(self.log), TRACKER_PATH=str(self.dir / "Tracker.xlsx"),
                                CONTACTS_PATH=None, PUBLISH_PATH=None, ATTACH_DIR=None, TRACK_SENT=False).items():
            p = mock.patch.object(doi_filter, name, value)
            p.start()
            self.addCleanup(p.stop)
        p = mock.patch.object(doi_filter, "_outlook", side_effect=lambda: fake_outlook(self.inbox))
        p.start()
        self.addCleanup(p.stop)

    def run_once(self, *args, upsert=fake_upsert):
        out = io.StringIO()
        with mock.patch.object(run, "upsert", side_effect=upsert) as up, redirect_stdout(out):
            code = run.main(list(args))
        return code, out.getvalue(), up.call_count

    def log_lines(self):
        return self.log.read_text(encoding="utf-8").splitlines() if self.log.exists() else []

    def test_second_run_with_nothing_new_does_nothing(self):
        code, out, calls = self.run_once()
        self.assertEqual((code, calls), (0, 1))
        lines = self.log_lines()
        self.assertTrue((self.dir / "Logs" / "DOI_Seen_Messages.csv").exists())
        code, out, calls = self.run_once()
        self.assertEqual((code, calls), (0, 0))                          # tracker not even opened
        self.assertIn("nothing to do", out)
        self.assertEqual(self.log_lines(), lines)                        # no new log lines

    def test_a_new_email_later_runs_the_workflow_again(self):
        self.run_once()
        self.inbox._messages.append(fake_msg("RE: FIN-301 Request - Premium Tax",
                                             "We received the DOI exam request.", "ana.accountant@corp.test",
                                             "E2", "T1"))
        code, out, calls = self.run_once()
        self.assertEqual(calls, 1)
        self.assertNotIn("nothing to do", out)

    def test_review_item_is_quiet_until_sender_is_approved(self):
        self.inbox._messages = [fake_msg(*UNKNOWN, "E9")]
        code, out, calls = self.run_once()
        self.assertIn("REVIEW", out)
        code, out, calls = self.run_once()
        self.assertIn("nothing to do", out)                              # still waiting: not repeated
        with mock.patch.object(doi_filter, "INTERNAL_DOMAINS", {"corp.test", "elsewhere.test"}):
            code, out, calls = self.run_once()                           # sender now known
        self.assertEqual(calls, 1)
        self.assertIn("FIN-302", out)

    def test_tracker_open_in_excel_means_retry_next_run(self):
        def locked(*a, **kw):
            raise TrackerError("Tracker.xlsx is open in Excel")
        code, out, calls = self.run_once(upsert=locked)
        self.assertIn("NOT WRITTEN", out)
        code, out, calls = self.run_once()
        self.assertEqual(calls, 1)                                       # tried again, now written
        code, out, calls = self.run_once()
        self.assertEqual(calls, 0)

    def test_dry_run_remembers_nothing(self):
        self.run_once("--dry-run")
        self.assertFalse((self.dir / "Logs" / "DOI_Seen_Messages.csv").exists())
        code, out, calls = self.run_once("--dry-run")
        self.assertEqual(calls, 1)

    def test_all_reports_everything_again(self):
        self.run_once()
        code, out, calls = self.run_once("--all")
        self.assertEqual(calls, 1)

    def test_scheduled_crash_goes_to_the_error_file(self):
        with mock.patch.object(doi_filter, "fetch_candidates", side_effect=RuntimeError("Outlook not running")):
            code, out, calls = self.run_once("--scheduled")
        self.assertEqual(code, 1)
        errors = (self.dir / "Logs" / "DOI_Run_Errors.txt").read_text(encoding="utf-8")
        self.assertIn("Outlook not running", errors)

    def test_unscheduled_crash_still_shows(self):
        with mock.patch.object(doi_filter, "fetch_candidates", side_effect=RuntimeError("boom")):
            with self.assertRaises(RuntimeError):
                self.run_once()


FIXTURE = Path(os.environ.get("DOI_TRACKER_FIXTURE", Path(__file__).parent / "fixtures" / "DOI_Exam_Test_Tracker_v2.xlsx"))


@unittest.skipUnless(FIXTURE.exists(), "needs a fake-data tracker (DOI_TRACKER_FIXTURE)")
class HourlyRunWithTrackerTests(RolesMixin, unittest.TestCase):
    """Same scenario against a real (fake-data) workbook: the second run leaves the file alone."""

    def test_second_run_leaves_the_tracker_untouched(self):
        d = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, d)
        tracker = d / "Tracker.xlsx"
        shutil.copy(FIXTURE, tracker)
        inbox = FakeFolder("Inbox", [fake_msg(*REQUEST, "E1", "T1")])
        with mock.patch.multiple(doi_filter, LOG_PATH=str(d / "DOI_Run_Log.csv"), TRACKER_PATH=str(tracker),
                                 CONTACTS_PATH=None, PUBLISH_PATH=None, ATTACH_DIR=None, TRACK_SENT=False), \
                mock.patch.object(doi_filter, "_outlook", return_value=fake_outlook(inbox)), \
                redirect_stdout(io.StringIO()) as out:
            run.main([])
            first = tracker.stat().st_mtime_ns
            self.assertIn("FIN-301", out.getvalue())
            run.main([])
        self.assertEqual(tracker.stat().st_mtime_ns, first)
        self.assertEqual(len(list((d / "Backups").glob("*.xlsx"))), 1)    # one backup, from the first save
        self.assertIn("nothing to do", out.getvalue())


# --- Task Scheduler definition ---------------------------------------------
class ScheduleTests(unittest.TestCase):
    NS = {"t": "http://schemas.microsoft.com/windows/2004/02/mit/task"}

    def parse(self, **kw):
        xml = schedule_hourly.task_xml(r"C:\Python\pythonw.exe", r"C:\Users\me\R&D Tools\run.py",
                                       r"C:\Users\me\R&D Tools", start=datetime(2027, 10, 8, 9, 15, 42), **kw)
        return ET.fromstring(xml.encode("utf-16"))

    def text(self, root, path):
        return root.find(path, self.NS).text

    def test_hourly_while_signed_in_and_catches_up_after_sleep(self):
        root = self.parse()
        self.assertEqual(self.text(root, ".//t:Repetition/t:Interval"), "PT1H")
        self.assertEqual(self.text(root, ".//t:StartBoundary"), "2027-10-08T09:15:00")
        self.assertEqual(self.text(root, ".//t:LogonType"), "InteractiveToken")      # no stored password
        self.assertEqual(self.text(root, ".//t:RunLevel"), "LeastPrivilege")         # no admin
        self.assertEqual(self.text(root, ".//t:StartWhenAvailable"), "true")
        self.assertEqual(self.text(root, ".//t:WakeToRun"), "false")                 # only "if the laptop is on"
        self.assertEqual(self.text(root, ".//t:DisallowStartIfOnBatteries"), "false")
        self.assertEqual(self.text(root, ".//t:MultipleInstancesPolicy"), "IgnoreNew")

    def test_runs_the_script_windowless_in_scheduled_mode(self):
        root = self.parse()
        self.assertEqual(self.text(root, ".//t:Command"), r"C:\Python\pythonw.exe")
        self.assertEqual(self.text(root, ".//t:Arguments"), r'"C:\Users\me\R&D Tools\run.py" --scheduled')
        self.assertEqual(self.text(root, ".//t:WorkingDirectory"), r"C:\Users\me\R&D Tools")

    def test_pythonw_falls_back_to_python(self):
        d = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, d)
        (d / "python.exe").touch()
        self.assertEqual(schedule_hourly.pythonw(d / "python.exe"), d / "python.exe")
        (d / "pythonw.exe").touch()
        self.assertEqual(schedule_hourly.pythonw(d / "python.exe"), d / "pythonw.exe")

    def test_show_changes_nothing(self):
        with mock.patch.object(schedule_hourly, "schtasks") as st, redirect_stdout(io.StringIO()) as out:
            self.assertEqual(schedule_hourly.main(["--show"]), 0)
        st.assert_not_called()
        self.assertIn("PT1H", out.getvalue())


if __name__ == "__main__":
    unittest.main()
