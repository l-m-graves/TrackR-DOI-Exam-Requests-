"""Tests for run.py's CSV audit log."""
import csv
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import run

ROW = {"run_at": "2027-10-08 09:00:00", "received": "2027-10-08 08:00", "ref": "ref 0000aaaa",
       "current_sender": "x@example.test", "subject": "FIN-53 Request", "result": "added", "detail": "row 5"}


class WriteLogTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir)

    def read(self, path):
        with open(path, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))

    def test_creates_folder_and_file_with_header_then_appends(self):
        path = self.dir / "Logs" / "DOI_Run_Log.csv"                  # folder doesn't exist yet
        self.assertEqual(run.write_log(path, [ROW]), path)
        run.write_log(path, [dict(ROW, result="unchanged")])
        rows = self.read(path)
        self.assertEqual([r["result"] for r in rows], ["added", "unchanged"])   # one header, two rows
        self.assertEqual(list(rows[0]), run.LOG_FIELDS)

    def test_old_format_log_is_kept_not_mixed(self):
        path = self.dir / "DOI_Run_Log.csv"
        path.write_text("run_at,received,ref,sender,subject,result,detail\n2027-01-01,,,,,added,\n", encoding="utf-8")
        run.write_log(path, [ROW])
        olds = list(self.dir.glob("DOI_Run_Log_Old_Format_*.csv"))
        self.assertEqual(len(olds), 1)
        self.assertIn("sender,subject", olds[0].read_text(encoding="utf-8"))
        self.assertEqual(list(self.read(path)[0]), run.LOG_FIELDS)

    def test_log_row_carries_detection_and_dates(self):
        from datetime import datetime
        from doi_analysis import analyze, detect
        from doi_filter import Mail
        m = Mail("FIN-53 Request", "Please provide the schedule by 10/20/2027.", "examiner@example.test",
                 received=datetime(2027, 10, 8), entry_id="E53", conversation_id="T53")
        m.detection, m.analysis = detect(m), analyze(m, "FIN-53")
        row = run.log_row(datetime(2027, 10, 8, 9), m, "added")
        self.assertEqual((row["request_id"], row["message_id"], row["conversation_id"]), ("FIN-53", "E53", "T53"))
        self.assertEqual((row["exam_due_date"], row["tracking_due_date"], row["date_source"]),
                         ("2027-10-20", "2027-10-20", "examiner"))
        self.assertIn("by 10/20/2027", row["date_evidence"])
        self.assertEqual(set(row) - set(run.LOG_FIELDS), set())

    def test_locked_log_goes_to_a_dated_fallback(self):
        path = self.dir / "DOI_Run_Log.csv"
        real_append = run._append_csv
        def locked_main(p, rows):
            if p == path:
                raise PermissionError("open in Excel")
            real_append(p, rows)
        with mock.patch.object(run, "_append_csv", side_effect=locked_main):
            used = run.write_log(path, [ROW])
        self.assertNotEqual(used, path)
        self.assertTrue(used.name.startswith("DOI_Run_Log_") and used.suffix == ".csv")
        self.assertEqual(self.read(used)[0]["subject"], "FIN-53 Request")


if __name__ == "__main__":
    unittest.main()
