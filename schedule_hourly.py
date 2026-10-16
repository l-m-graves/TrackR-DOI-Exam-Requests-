"""Run the tracker update every hour while the laptop is on (Windows Task Scheduler).

    python schedule_hourly.py            # create (or replace) the hourly task
    python schedule_hourly.py --show     # print the task definition, change nothing
    python schedule_hourly.py --remove   # delete the task

What the task does:
  - every 60 minutes, only while you are signed in (it runs as you, in your session, so it can
    use your open Outlook - no password is stored and no admin rights are needed)
  - if the laptop was asleep or off at run time, it runs once soon after it wakes up / you sign
    in, instead of piling up the missed runs
  - runs `pythonw run.py --scheduled`: no console window pops up. Each run checks the Inbox; if
    no new DOI-related email is there it stops without touching anything. Crashes are written to
    DOI_Run_Errors.txt next to the run log.
  - never wakes the laptop, and runs on battery too
  - if a run is still going after 30 minutes it is stopped; a new run never starts on top of one

Check it in Task Scheduler (Start menu -> "Task Scheduler" -> Task Scheduler Library ->
"DOI Exam Tracker"): the "Last Run Result" column shows 0x0 when a run finished normally.
"""
import argparse
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

TASK_NAME = "DOI Exam Tracker"
HERE = Path(__file__).resolve().parent


def pythonw(executable=sys.executable) -> Path:
    """pythonw.exe (no console window) next to the running python.exe; python.exe if missing."""
    exe = Path(executable)
    windowless = exe.with_name("pythonw.exe")
    return windowless if windowless.exists() else exe


def task_xml(python, script, workdir, start=None) -> str:
    """The Task Scheduler definition. Pure text - tested on Linux, registered on Windows."""
    start = f"{(start or datetime.now()).replace(second=0, microsecond=0):%Y-%m-%dT%H:%M:%S}"
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Reads DOI exam-request emails from Outlook (read-only) and updates the tracker. Does nothing when no new DOI email arrived.</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <StartBoundary>{start}</StartBoundary>
      <Enabled>true</Enabled>
      <Repetition>
        <Interval>PT1H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <WakeToRun>false</WakeToRun>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{escape(str(python))}</Command>
      <Arguments>"{escape(str(script))}" --scheduled</Arguments>
      <WorkingDirectory>{escape(str(workdir))}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def schtasks(*args):
    """Windows' own command-line front end to Task Scheduler."""
    done = subprocess.run(["schtasks", *args], capture_output=True, text=True)
    print((done.stdout or done.stderr).strip())
    return done.returncode


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--show", action="store_true", help="print the task definition only")
    ap.add_argument("--remove", action="store_true", help="delete the hourly task")
    args = ap.parse_args(argv)
    if args.remove:
        return schtasks("/Delete", "/TN", TASK_NAME, "/F")
    xml = task_xml(pythonw(), HERE / "run.py", HERE)
    if args.show:
        print(xml)
        return 0
    if sys.platform != "win32":
        print("Task Scheduler is Windows-only - run this on your work laptop.")
        return 1
    # schtasks reads the definition from a file, and Task Scheduler expects it in UTF-16.
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "DOI_Exam_Tracker_Task.xml"
        path.write_text(xml, encoding="utf-16")
        code = schtasks("/Create", "/TN", TASK_NAME, "/XML", str(path), "/F")
    if code == 0:
        print(f"Runs every hour while you're signed in. Try it now: schtasks /Run /TN \"{TASK_NAME}\"")
    return code


if __name__ == "__main__":
    sys.exit(main())
