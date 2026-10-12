"""Tests for run.py's CSV audit log."""
import csv
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import run

ROW = {"run_at": "2027-10-08 09:00:00", "received": "2027-10-08 08:00", "ref": "ref 0000aaaa",
       "sender": "x@example.test", "subject": "FIN-53 Request", "result": "added", "detail": "row 5"}


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
