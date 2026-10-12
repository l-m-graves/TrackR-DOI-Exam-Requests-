# Copy this file to local_settings.py and fill in real values.
# local_settings.py is git-ignored, so real paths and addresses stay on your laptop.
# Anything set here overrides the placeholder of the same name in doi_filter.py.

TRACKER_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Exam_Test_Tracker_v2.xlsx"
# Folders may be empty - the script creates the log file and one sub-folder per request ID.
# Request attachments are saved as ATTACH_DIR\FIN-53\2027-10-08_<file name>. Delete the line to turn off.
ATTACH_DIR   = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Attachments"
# Saved types (default): ATTACH_TYPES = (".pdf", ".doc", ".docx", ".png", ".xlsx") - add e.g. ".xls" if needed
# The log is a FILE path: folder + a file name ending in .csv
LOG_PATH     = r"C:\Users\<you>\OneDrive - <Company>\Audit\Logs\DOI_Run_Log.csv"
# Contact directory: every sender/To/CC on request emails, updated each run. New senders are
# only trusted after you type Y in its Approved column. Add team-meeting contacts there by hand.
CONTACTS_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Contacts.csv"
# View-only copy for coworkers (a shared folder they can open). Must differ from TRACKER_PATH.
PUBLISH_PATH = r"C:\Users\<you>\<Company>\<Shared folder>\DOI_Exam_Tracker_View_Only.xlsx"

# Whole domains you trust (e.g. the DOI's): the part AFTER the @, e.g. "insurance.ca.gov".
# Never a public provider (gmail.com, icloud.com...) - that trusts everyone on it; approve
# individual people in DOI_Contacts.csv instead.
SENDER_DOMAINS  = {"<real DOI domain>"}
# Request IDs default to FIN-##, IT-## and DOI-TEST-##. Only if a new prefix appears (e.g. ABC-12),
# override both - keep them in step:
# REQUEST_ID_PATTERN = r"\b(?:FIN|IT|ABC|DOI-[A-Z]+)-\d+\b"
# SUBJECT_LIKES      = ("%FIN-%", "%IT-%", "%ABC-%", "%DOI-%")
# Optional - leave out (or set()) while the To line varies between examiner and assigned contact.
# MUST_BE_SENT_TO = {"<one address every request email must reach>"}
