# Copy this file to local_settings.py and fill in real values.
# local_settings.py is git-ignored, so real paths and addresses stay on your laptop.
# Anything set here overrides the placeholder of the same name in doi_filter.py.

TRACKER_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Exam_Test_Tracker_v2.xlsx"
ATTACH_DIR   = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI Attachments"
LOG_PATH     = r"C:\Users\<you>\OneDrive - <Company>\Audit\doi_run_log.csv"
# View-only copy for coworkers (a shared folder they can open). Must differ from TRACKER_PATH.
PUBLISH_PATH = r"C:\Users\<you>\<Company>\<Shared folder>\DOI Exam Tracker (view only).xlsx"

SENDER_DOMAINS  = {"<real DOI domain>"}
MUST_BE_SENT_TO = {"<audit mailbox address>"}
