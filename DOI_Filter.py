"""Qualify Outlook mail items as DOI exam requests.

Two layers:
  1. Items.Restrict()  - coarse, runs inside Outlook (fast, no Python loop over 10k items)
  2. qualify()         - precise rules in plain Python, testable without Outlook
"""
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

# --- settings --------------------------------------------------------------
# Placeholders only. Put your real paths/addresses in Local_Settings.py (git-ignored,
# see Local_Settings_Example.py) so they never get committed to GitHub.
TRACKER_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Exam_Test_Tracker_v2.xlsx"
ATTACH_DIR   = None    # folder for request attachments, e.g. r"C:\...\DOI Attachments"; None = don't save
LOG_PATH     = r"C:\Users\<you>\OneDrive - <Company>\Audit\Logs\DOI_Run_Log.csv"
PUBLISH_PATH = None    # view-only copy for coworkers, e.g. a shared folder; None = don't publish
CONTACTS_PATH = None   # contact directory the script keeps up to date (DOI_Contacts.csv); None = off

# Request IDs: FIN-53, IT-39 (real); DOI-TEST-41 (test data). If a new prefix appears, add it to
# BOTH lines - the Outlook pre-filter has to agree with the pattern or those emails never arrive.
REQUEST_ID_PATTERN = r"\b(?:FIN|IT|DOI-[A-Z]+)-\d+\b"
SUBJECT_LIKES   = ("%FIN-%", "%IT-%", "%DOI-%")               # Outlook subject pre-filter (DASL LIKE, any one)
SENDER_DOMAINS  = {"insurance.ca.gov", "example.test"}        # adjust to the real DOI domain
# Optional. Leave empty to skip the rule - the To line changes as a request moves between the
# examiner and the assigned contact. Set it only if every request email must reach one address.
MUST_BE_SENT_TO = set()
BODY_PHRASES    = ("request", "please provide")               # any one must appear
IMPORTANT_WORDS = ("DOI", "Request", "Follow-up")             # attachment file names containing any of these
IGNORE_EXT      = (".png", ".gif", ".jpg", ".jpeg")           # rule 5 ignores images (signature logos)
ATTACH_TYPES    = (".pdf", ".doc", ".docx", ".png", ".xlsx")  # what ATTACH_DIR saves from qualifying emails
SMTP_PROP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"   # PR_SMTP_ADDRESS

try:
    from Local_Settings import *    # noqa: F401,F403  - real values override the placeholders above
except ImportError:
    try:                            # older name, from before the Title_Case renaming
        from local_settings import *    # noqa: F401,F403
        print("Note: rename local_settings.py to Local_Settings.py (the old name still works for now).")
    except ImportError:
        pass

REQUEST_ID = re.compile(REQUEST_ID_PATTERN)     # compiled after Local_Settings so an override applies


def normalize_domains(domains) -> set:
    """Forgive '@insurance.ca.gov' / 'Insurance.CA.gov': the check compares lower-case text after the '@'."""
    return {d.strip().lstrip("@").lower() for d in domains}


SENDER_DOMAINS = normalize_domains(SENDER_DOMAINS)


def subject_filter() -> str:
    """DASL filter for Items.Restrict(): subject matches any prefix AND the email has attachments.
    Deliberately loose ('%IT-%' also matches 'AUDIT-'); qualify() does the precise check."""
    likes = " OR ".join(f'"urn:schemas:httpmail:subject" LIKE \'{p}\'' for p in SUBJECT_LIKES)
    return f'@SQL=({likes}) AND "urn:schemas:httpmail:hasattachment" = 1'


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
    saved: list = field(default_factory=list)   # attachment results ("saved <path>" / "FAILED ...")
    names: dict = field(default_factory=dict)   # address -> display name, for the contact directory


def _norm(text: str) -> str:
    """Case-, space- and hyphen-insensitive: 'Follow-up', 'follow up', 'FollowUp' all match."""
    return text.lower().replace("-", "").replace(" ", "").replace("_", "")


def important_attachments(names) -> list:
    """File names that contain any IMPORTANT_WORDS (after normalising), images excluded."""
    keys = [_norm(w) for w in IMPORTANT_WORDS]
    return [n for n in names
            if not n.lower().endswith(IGNORE_EXT) and any(k in _norm(n) for k in keys)]


INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
INLINE_IMAGE = re.compile(r"^image\d{3}\.(png|jpe?g|gif)$", re.I)   # Outlook's names for logos pasted in a body


def attachment_targets(rid: str, names, received: datetime, folder) -> list:
    """Which attachments to save, and where: [(outlook_index, destination), ...].

    Every attachment of an ATTACH_TYPES type is kept (the whole email is request evidence),
    except Outlook's auto-named inline images (image001.png = signature logos).

    ATTACH_DIR / <RequestID> / <YYYY-MM-DD>_<file name> - one folder per request, the received date
    in front so a revised 'Request.pdf' on a later email doesn't collide with the first one.
    Already-saved files are skipped, so re-runs don't duplicate. Pure Python - no Outlook."""
    targets = []
    for index, name in enumerate(names, start=1):           # Outlook's Attachments are 1-based
        if not name.lower().endswith(ATTACH_TYPES) or INLINE_IMAGE.match(name):
            continue
        safe = INVALID_CHARS.sub("_", name).strip(" .")
        stem, dot, ext = safe.rpartition(".")
        if dot and len(stem) > 100:                          # keep under Windows' 260-char path limit
            safe = f"{stem[:100]}.{ext}"
        dest = Path(folder) / rid / f"{received:%Y-%m-%d}_{safe}"
        if not dest.exists() and dest not in [d for _, d in targets]:
            targets.append((index, dest))
    return targets


def save_attachments(msg, m: Mail, rid: str, folder) -> list:
    """COM: write the planned attachments to disk. One failure doesn't stop the others."""
    results = []
    for index, dest in attachment_targets(rid, m.attachments, m.received, folder):
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            msg.Attachments.Item(index).SaveAsFile(str(dest))
            results.append(f"saved {dest}")
        except Exception as e:
            results.append(f"FAILED {dest.name}: {e}")
    return results


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
    to, cc, names = [], [], {}
    for r in msg.Recipients:                      # r.Type: 1 = To, 2 = CC, 3 = BCC
        if r.Type in (1, 2):                      # BCC is not "on To or CC" - leave it out
            address = smtp_address(r)
            (to if r.Type == 1 else cc).append(address)
            names[address] = getattr(r, "Name", "") or ""
    return Mail(
        subject=msg.Subject or "",
        body=msg.Body or "",
        sender=smtp_address(msg.Sender) if msg.Sender else (msg.SenderEmailAddress or "").lower(),
        to=to, cc=cc,
        attachments=[a.FileName for a in msg.Attachments],
        received=_naive(msg.ReceivedTime),
        entry_id=msg.EntryID,
        sender_name=msg.SenderName or "",
        names=names,
    )


# --- the rules -------------------------------------------------------------
def qualify(m: Mail, approved_senders=()) -> tuple[bool, str]:
    """Return (ok, reason). Every rejection names the rule that failed - that is your audit trail.
    approved_senders: individual addresses you approved in the contact directory."""
    rid = REQUEST_ID.search(m.subject)
    if not rid:
        return False, "subject has no request ID"
    if m.sender.split("@")[-1].lower() not in SENDER_DOMAINS and m.sender not in approved_senders:
        return False, f"sender not approved: {m.sender}"
    if MUST_BE_SENT_TO and not MUST_BE_SENT_TO & set(m.to + m.cc):
        return False, "not addressed to MUST_BE_SENT_TO"
    body = m.body.lower()
    if not any(p in body for p in BODY_PHRASES):
        return False, "body lacks request wording"
    if not important_attachments(m.attachments):
        return False, "no attachment named with DOI / Request / Follow-up"
    return True, rid.group(0)


# --- Outlook side ----------------------------------------------------------
def fetch_candidates(since: datetime, folder_path=("Inbox", "DOI Exam"), attach_dir=None, approved_senders=()):
    """Coarse filter inside Outlook, then precise rules in Python. With attach_dir, qualifying
    emails' request attachments are saved there (the Outlook item is only reachable here)."""
    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = ns.GetDefaultFolder(6)                          # Inbox
    for name in folder_path[1:]:                             # walk into sub-folders by name
        folder = folder.Folders(name)
    items = folder.Items
    items = items.Restrict(f"[ReceivedTime] >= '{since:%m/%d/%Y} 12:00 AM'")     # Jet syntax
    items = items.Restrict(subject_filter())                                     # DASL syntax
    accepted, rejected = [], []
    for msg in items:
        if msg.Class != 43:                                  # olMail only
            continue
        m = extract(msg)
        ok, reason = qualify(m, approved_senders)
        if ok and attach_dir:
            m.saved = save_attachments(msg, m, reason, attach_dir)
        (accepted if ok else rejected).append((m, reason))
    return accepted, rejected
