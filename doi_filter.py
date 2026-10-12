"""Qualify Outlook mail items as DOI exam requests.

Two layers:
  1. Items.Restrict()  - coarse, runs inside Outlook (fast, no Python loop over 10k items)
  2. qualify()         - precise rules in plain Python, testable without Outlook
"""
import re
from dataclasses import dataclass, field
from datetime import datetime

# --- settings --------------------------------------------------------------
# Placeholders only. Put your real paths/addresses in local_settings.py (git-ignored,
# see local_settings.example.py) so they never get committed to GitHub.
TRACKER_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Exam_Test_Tracker_v2.xlsx"
ATTACH_DIR   = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI Attachments"
LOG_PATH     = r"C:\Users\<you>\OneDrive - <Company>\Audit\doi_run_log.csv"
PUBLISH_PATH = None    # view-only copy for coworkers, e.g. a shared folder; None = don't publish

# Request ID format. If your real IDs use another prefix (e.g. ABC-12), set both of these in
# local_settings.py - the Outlook pre-filter has to agree with the pattern or emails never arrive.
REQUEST_ID_PATTERN = r"\bDOI-[A-Z]+-\d+\b"
SUBJECT_LIKE    = "%DOI-%"                                    # Outlook subject pre-filter (DASL LIKE)
SENDER_DOMAINS  = {"insurance.ca.gov", "example.test"}        # adjust to the real DOI domain
MUST_BE_SENT_TO = {"audit@example.test"}                      # the audit mailbox, or your own address
BODY_PHRASES    = ("request", "please provide")               # any one must appear
IMPORTANT_WORDS = ("DOI", "Request", "Follow-up")             # attachment file names containing any of these
IGNORE_EXT      = (".png", ".gif", ".jpg", ".jpeg")           # signature logos are attachments too
SMTP_PROP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"   # PR_SMTP_ADDRESS

try:
    from local_settings import *    # noqa: F401,F403  - real values override the placeholders above
except ImportError:
    pass

REQUEST_ID = re.compile(REQUEST_ID_PATTERN)     # compiled after local_settings so an override applies


# --- plain-Python view of a message (what the rules operate on) -----------
@dataclass
class Mail:
    subject: str
    body: str
    sender: str                     # SMTP address, lower-case
    to: list = field(default_factory=list)      # SMTP addresses on the To line
    cc: list = field(default_factory=list)
    attachments: list = field(default_factory=list)   # file names
    received: datetime = None       # naive local time (see extract())
    entry_id: str = ""              # Outlook's stable ID for the item
    sender_name: str = ""           # display name, for the tracker's Contact column


def _norm(text: str) -> str:
    """Case-, space- and hyphen-insensitive: 'Follow-up', 'follow up', 'FollowUp' all match."""
    return text.lower().replace("-", "").replace(" ", "").replace("_", "")


def important_attachments(names) -> list:
    """File names that contain any IMPORTANT_WORDS (after normalising), images excluded."""
    keys = [_norm(w) for w in IMPORTANT_WORDS]
    return [n for n in names
            if not n.lower().endswith(IGNORE_EXT) and any(k in _norm(n) for k in keys)]


def smtp_address(recipient_or_sender):
    """Exchange returns X500 strings for internal people; PR_SMTP_ADDRESS gives the real one."""
    try:
        return recipient_or_sender.PropertyAccessor.GetProperty(SMTP_PROP).lower()
    except Exception:
        return (getattr(recipient_or_sender, "Address", "") or "").lower()


def _naive(dt):
    """pywin32 hands back a timezone-tagged datetime; openpyxl refuses those.
    Outlook's value is already local time, so rebuild it without the tag."""
    if dt is None:
        return None
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


def extract(msg) -> Mail:
    """COM -> Mail. Only this function touches Outlook."""
    to, cc = [], []
    for r in msg.Recipients:                      # r.Type: 1 = To, 2 = CC, 3 = BCC
        if r.Type == 1:
            to.append(smtp_address(r))
        elif r.Type == 2:                         # BCC is not "on To or CC" - leave it out
            cc.append(smtp_address(r))
    return Mail(
        subject=msg.Subject or "",
        body=msg.Body or "",
        sender=smtp_address(msg.Sender) if msg.Sender else (msg.SenderEmailAddress or "").lower(),
        to=to, cc=cc,
        attachments=[a.FileName for a in msg.Attachments],
        received=_naive(msg.ReceivedTime),
        entry_id=msg.EntryID,
        sender_name=msg.SenderName or "",
    )


# --- the rules -------------------------------------------------------------
def qualify(m: Mail) -> tuple[bool, str]:
    """Return (ok, reason). Every rejection names the rule that failed - that is your audit trail."""
    rid = REQUEST_ID.search(m.subject)
    if not rid:
        return False, "subject has no request ID"
    if m.sender.split("@")[-1] not in SENDER_DOMAINS:
        return False, f"sender domain not allowed: {m.sender}"
    if not MUST_BE_SENT_TO & set(m.to + m.cc):
        return False, "not addressed to the audit mailbox"
    body = m.body.lower()
    if not any(p in body for p in BODY_PHRASES):
        return False, "body lacks request wording"
    if not important_attachments(m.attachments):
        return False, "no attachment named with DOI / Request / Follow-up"
    return True, rid.group(0)


# --- Outlook side ----------------------------------------------------------
def fetch_candidates(since: datetime, folder_path=("Inbox", "DOI Exam")):
    """Coarse filter inside Outlook, then precise rules in Python."""
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = ns.GetDefaultFolder(6)                          # Inbox
    for name in folder_path[1:]:                             # walk into sub-folders by name
        folder = folder.Folders(name)
    items = folder.Items
    items = items.Restrict(f"[ReceivedTime] >= '{since:%m/%d/%Y} 12:00 AM'")     # Jet syntax
    items = items.Restrict(f'@SQL="urn:schemas:httpmail:subject" LIKE \'{SUBJECT_LIKE}\' '
                           'AND "urn:schemas:httpmail:hasattachment" = 1')        # DASL syntax
    accepted, rejected = [], []
    for msg in items:
        if msg.Class != 43:                                  # olMail only
            continue
        m = extract(msg)
        ok, reason = qualify(m)
        (accepted if ok else rejected).append((m, reason))
    return accepted, rejected
