"""Publish a view-only copy of the tracker for coworkers.

Your tracker (TRACKER_PATH) stays the single source of truth. After each run, if it has
changed since the last publish, a copy goes to PUBLISH_PATH:

  1. copy the tracker into a private temp folder
  2. have Excel open that copy, calculate every formula, and save it      <- excel_recalc()
     (openpyxl saves formulas without results, so a copy opened in Protected View or a
     browser preview would otherwise show blank Request IDs, flags and Summary figures)
  3. swap it into place in one step, so coworkers never see a half-written file

If a coworker has the shared copy open, or Excel can't calculate it, nothing is published
and the old copy stays put - the next run tries again.
"""
import os
import shutil
import tempfile
from pathlib import Path


def excel_recalc(path: Path):
    """The only function in this module that touches Excel (COM). Tested on Windows only."""
    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")   # a separate, hidden Excel -
    excel.Visible = False                                     # leaves your open windows alone
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0)
        excel.CalculateFull()
        # ReadOnlyRecommended: Excel asks coworkers "open as read-only?" - a reminder that
        # edits belong in the master, not this copy. 51 = .xlsx format.
        wb.SaveAs(str(path), FileFormat=51, ReadOnlyRecommended=True)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()


def needs_publish(master: Path, dest: Path) -> bool:
    """True if the shared copy is missing or older than the master. This also picks up
    edits you typed into the master yourself, not only the script's changes."""
    return not dest.exists() or master.stat().st_mtime > dest.stat().st_mtime


def publish(master, dest, recalc=excel_recalc) -> str:
    """Returns a one-line result for the log. Never raises for a locked or failed publish."""
    master, dest = Path(master), Path(dest)
    if master.resolve() == dest.resolve():
        raise ValueError("PUBLISH_PATH must be a different file from TRACKER_PATH")
    if not needs_publish(master, dest):
        return "publish skipped: shared copy already up to date"

    work = Path(tempfile.mkdtemp(prefix="doi_publish_"))
    staged = dest.with_name(dest.name + ".publishing")        # same folder -> the swap is atomic
    try:
        snapshot = work / "publish_snapshot.xlsx"                # different name from the master, so
        shutil.copy2(master, snapshot)                           # Excel never sees two files named alike
        try:
            recalc(snapshot)
        except Exception as e:
            return f"publish FAILED (Excel could not calculate the copy): {e}"
        shutil.copy(snapshot, staged)                            # copy (not copy2): stamp it 'now'
        try:
            os.replace(staged, dest)
        except PermissionError:
            return f"publish skipped: {dest.name} is open by someone - will retry next run"
        return f"published to {dest}"
    finally:
        shutil.rmtree(work, ignore_errors=True)
        if staged.exists():
            staged.unlink()
