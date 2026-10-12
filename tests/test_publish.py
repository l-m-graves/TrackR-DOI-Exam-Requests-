"""Tests for publish.py. Excel isn't available here, so a fake recalc stands in for
excel_recalc(); the real one is tested by running run.py on Windows."""
import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from publish import needs_publish, publish


class PublishTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp)
        self.master = self.tmp / "mine" / "tracker.xlsx"
        self.dest = self.tmp / "shared" / "DOI_Exam_Tracker_View_Only.xlsx"
        self.master.parent.mkdir()
        self.dest.parent.mkdir()
        self.master.write_bytes(b"version 1")
        self.recalc_calls = []

    def fake_recalc(self, path):
        """Stands in for Excel: records which file it was given and 'calculates' it."""
        self.recalc_calls.append(Path(path))
        Path(path).write_bytes(Path(path).read_bytes() + b" +calculated")

    def edit_master(self, content):
        self.master.write_bytes(content)
        later = time.time() + 5
        os.utime(self.master, (later, later))          # make "newer" unambiguous

    def leftovers(self):
        return [p.name for p in self.dest.parent.iterdir() if p != self.dest]

    def test_first_publish_writes_calculated_copy(self):
        result = publish(self.master, self.dest, recalc=self.fake_recalc)
        self.assertTrue(result.startswith("published"))
        self.assertEqual(self.dest.read_bytes(), b"version 1 +calculated")
        self.assertEqual(self.master.read_bytes(), b"version 1")        # master untouched
        [called] = self.recalc_calls
        self.assertNotIn(called.parent, (self.master.parent, self.dest.parent))  # Excel works on a temp copy
        self.assertFalse(called.exists())                                # temp cleaned up
        self.assertEqual(self.leftovers(), [])

    def test_unchanged_master_is_not_republished(self):
        publish(self.master, self.dest, recalc=self.fake_recalc)
        result = publish(self.master, self.dest, recalc=self.fake_recalc)
        self.assertIn("up to date", result)
        self.assertEqual(len(self.recalc_calls), 1)

    def test_edited_master_is_republished(self):
        publish(self.master, self.dest, recalc=self.fake_recalc)
        self.edit_master(b"version 2")
        self.assertTrue(needs_publish(self.master, self.dest))
        publish(self.master, self.dest, recalc=self.fake_recalc)
        self.assertEqual(self.dest.read_bytes(), b"version 2 +calculated")

    def test_shared_copy_open_leaves_old_copy_and_no_debris(self):
        publish(self.master, self.dest, recalc=self.fake_recalc)
        self.edit_master(b"version 2")
        with mock.patch("publish.os.replace", side_effect=PermissionError):
            result = publish(self.master, self.dest, recalc=self.fake_recalc)
        self.assertIn("open by someone", result)
        self.assertEqual(self.dest.read_bytes(), b"version 1 +calculated")
        self.assertEqual(self.leftovers(), [])
        self.assertTrue(publish(self.master, self.dest, recalc=self.fake_recalc).startswith("published"))

    def test_excel_failure_publishes_nothing(self):
        def broken(path):
            raise RuntimeError("Excel not available")
        result = publish(self.master, self.dest, recalc=broken)
        self.assertIn("FAILED", result)
        self.assertFalse(self.dest.exists())
        self.assertEqual(self.leftovers(), [])

    def test_refuses_to_publish_over_the_master(self):
        with self.assertRaises(ValueError):
            publish(self.master, self.master, recalc=self.fake_recalc)


if __name__ == "__main__":
    unittest.main()
