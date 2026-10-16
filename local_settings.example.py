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
# Last Email Sent (column M) = the newest email in each request's exchange, whoever sent it.
# Set True to also count your own Sent Items (newest date wins).
# TRACK_SENT = True
# View-only copy for coworkers (a shared folder they can open). Must differ from TRACKER_PATH.
PUBLISH_PATH = r"C:\Users\<you>\<Company>\<Shared folder>\DOI_Exam_Tracker_View_Only.xlsx"

# --- Who is who (the ONLY place people/domains are named). Addresses and domains, lower-case. ---
# These decide how an email is judged (accept / review / reject) and whose dates are deadlines:
# a date from the examiner side is the exam due date; a date from an internal person is only an
# internal expected date. Leave a set empty if you have nobody in that group.
LEAD_EXAMINERS       = {"<lead examiner address>"}
EXAMINER_CONTACTS    = set()                        # other examiner addresses
EXAMINER_DOMAINS     = {"<regulator domain>"}
THIRD_PARTY_CONTACTS = set()                        # examination-firm people working for the DOI
THIRD_PARTY_DOMAINS  = {"<exam firm domain>"}       # e.g. the firm(s) the DOI uses
ACCOUNTANT_CONTACTS  = {"<accountant address>"}     # internal accountants coordinating requests
INTERNAL_DOMAINS     = {"<company domain>"}         # any other employee
# ENTITY_NAMES  = ("<legal entity name>",)          # optional: spotted in request scope
# PRODUCT_NAMES = ("<product name>",)

# Whole domains you trust (e.g. the DOI's): the part AFTER the @, e.g. "insurance.ca.gov".
# (Counted as examiner domains. Kept for the original five qualification rules.)
# Never a public provider (gmail.com, icloud.com...) - that trusts everyone on it; approve
# individual people in DOI_Contacts.csv instead.
SENDER_DOMAINS  = {"<real DOI domain>"}
# Request IDs default to FIN-##, IT-## and DOI-TEST-##. Only if a new prefix appears (e.g. ABC-12),
# override both - keep them in step:
# REQUEST_ID_PATTERN = r"\b(?:FIN|IT|ABC|DOI-[A-Z]+)-\d+\b"
# SUBJECT_LIKES      = ("%FIN-%", "%IT-%", "%ABC-%", "%DOI-%")
# Optional - leave out (or set()) while the To line varies between examiner and assigned contact.
# MUST_BE_SENT_TO = {"<one address every request email must reach>"}
