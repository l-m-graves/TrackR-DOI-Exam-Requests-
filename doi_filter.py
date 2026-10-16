"""Qualify Outlook mail items as DOI exam requests.

Two layers:
  1. Items.Restrict()  - coarse, runs inside Outlook (fast, no Python loop over 10k items)
  2. qualify()         - precise rules in plain Python, testable without Outlook
"""
import hashlib
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

# --- settings --------------------------------------------------------------
# Placeholders only. Put your real paths/addresses in local_settings.py (git-ignored,
# see local_settings.example.py) so they never get committed to GitHub.
TRACKER_PATH = r"C:\Users\<you>\OneDrive - <Company>\Audit\DOI_Exam_Test_Tracker_v2.xlsx"
ATTACH_DIR   = None    # folder for request attachments, e.g. r"C:\...\DOI_Attachments"; None = don't save
LOG_PATH     = r"C:\Users\<you>\OneDrive - <Company>\Audit\Logs\DOI_Run_Log.csv"
PUBLISH_PATH = None    # view-only copy for coworkers, e.g. a shared folder; None = don't publish
CONTACTS_PATH = None   # contact directory the script keeps up to date (DOI_Contacts.csv); None = off
# Last Email Sent (column M) = the newest email in the request's exchange, whoever sent it, among
# the emails in your mailbox. Your own Sent Items are extra and off by default - you rarely send.
# True = count them too (newest wins).
TRACK_SENT   = False

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

# Who is who. The real values live ONLY in local_settings.py (addresses / domains, lower-case);
# nothing else in the program names people. SENDER_DOMAINS above counts as examiner domains.
LEAD_EXAMINERS       = set()   # the lead examiner's address(es) - kept separate on purpose
EXAMINER_CONTACTS    = set()   # other examiner addresses
EXAMINER_DOMAINS     = set()   # regulator domain(s)
THIRD_PARTY_CONTACTS = set()   # examination-firm people working for the DOI
THIRD_PARTY_DOMAINS  = set()   # examination-firm domain(s)
ACCOUNTANT_CONTACTS  = set()   # internal accountants who coordinate requests
INTERNAL_DOMAINS     = set()   # the company's own domain(s): any other employee
ENTITY_NAMES         = ()      # legal-entity names to spot in request scope (optional)
PRODUCT_NAMES        = ()      # product / line-of-business names to spot in scope (optional)

try:
    from local_settings import *    # noqa: F401,F403  - real values override the placeholders above
except ImportError:
    try:                            # in case it was renamed during the brief Title_Case period
        from Local_Settings import *    # noqa: F401,F403
        print("Note: rename Local_Settings.py back to local_settings.py (this name still works for now).")
    except ImportError:
        pass

REQUEST_ID = re.compile(REQUEST_ID_PATTERN)     # compiled after local_settings so an override applies
# Case-sensitive on purpose: with a pattern that allows 'FIN 45' / 'IT39', ignoring case would
# read ordinary words ('send it 5 days early') as request IDs.


def canonical_id(raw: str) -> str:
    """One spelling per request, whatever the email wrote: 'FIN 45' / 'FIN45' -> 'FIN-45'.
    IDs already hyphenated ('FIN-45', 'DOI-TEST-01') are unchanged. The tracker, the log and the
    memory all key on this form, so 'FIN 45' and 'FIN-45' are the same request."""
    return re.sub(r"^([A-Z]+) ?(?=\d)", r"\1-", (raw or "").strip().upper())


def find_id(text) -> str:
    """The first request ID in a text, in canonical form, or ''."""
    hit = REQUEST_ID.search(text or "")
    return canonical_id(hit.group(0)) if hit else ""


def log_file(path):
    """LOG_PATH may be the log FILE or just its folder - a folder gets DOI_Run_Log.csv inside it."""
    if not path or Path(str(path)).suffix.lower() == ".csv":
        return path
    return str(Path(str(path)) / "DOI_Run_Log.csv")


LOG_PATH = log_file(LOG_PATH)


def normalize_domains(domains) -> set:
    """Forgive '@insurance.ca.gov', 'Insurance.CA.gov' or a whole address ('someone@insurance.ca.gov'):
    the check compares the lower-case text after the '@'."""
    return {d.strip().lower().rsplit("@", 1)[-1] for d in domains if d and d.strip()}


SENDER_DOMAINS = normalize_domains(SENDER_DOMAINS)
EXAMINER_DOMAINS = normalize_domains(EXAMINER_DOMAINS)
THIRD_PARTY_DOMAINS = normalize_domains(THIRD_PARTY_DOMAINS)
INTERNAL_DOMAINS = normalize_domains(INTERNAL_DOMAINS)
LEAD_EXAMINERS, EXAMINER_CONTACTS, THIRD_PARTY_CONTACTS, ACCOUNTANT_CONTACTS = (
    {a.strip().lower() for a in group}
    for group in (LEAD_EXAMINERS, EXAMINER_CONTACTS, THIRD_PARTY_CONTACTS, ACCOUNTANT_CONTACTS))
MUST_BE_SENT_TO = {a.strip().lower() for a in MUST_BE_SENT_TO if a and a.strip()}

EXAMINER_SIDE = {"lead_examiner", "examiner", "third_party"}
INTERNAL_SIDE = {"accountant", "internal"}


def role_of(address, approved=()) -> str:
    """Which configured group an address belongs to (see the settings above). Pure."""
    a = (address or "").strip().lower()
    if "@" not in a:
        return "unknown"
    d = a.split("@")[-1]
    if a in LEAD_EXAMINERS:
        return "lead_examiner"
    if a in EXAMINER_CONTACTS or d in EXAMINER_DOMAINS:
        return "examiner"
    if a in THIRD_PARTY_CONTACTS or d in THIRD_PARTY_DOMAINS:
        return "third_party"
    if a in ACCOUNTANT_CONTACTS:
        return "accountant"
    if d in INTERNAL_DOMAINS:
        return "internal"
    if d in SENDER_DOMAINS:             # legacy list - checked last, so if it also holds your own
        return "examiner"               # company's domain, colleagues still count as internal
    if a in approved:
        return "approved"
    return "unknown"


SUBJECT_TERMS = ("%DOI%", "%Examination%", "%Exam Request%")   # DOI wording in a subject


def subject_filter() -> str:
    """DASL filter for Items.Restrict(): a request-ID prefix in the subject or body, or DOI wording
    in the subject. No attachment requirement - CC'd status emails and forwards often have none.
    Deliberately loose ('%IT-%' also matches 'AUDIT-'); Python makes the real decision."""
    subj, body = '"urn:schemas:httpmail:subject"', '"urn:schemas:httpmail:textdescription"'
    likes = [f"{subj} LIKE '{p}'" for p in SUBJECT_LIKES + SUBJECT_TERMS]
    likes += [f"{body} LIKE '{p}'" for p in SUBJECT_LIKES]
    return "@SQL=" + " OR ".join(likes)


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
    conversation_id: str = ""       # Outlook's thread ID - links your replies in Sent Items
    attachment_texts: dict = field(default_factory=dict)   # file name -> text read from it
    attachment_status: list = field(default_factory=list)  # "<file>: read" / "<file>: not read - why"
    detection: object = None        # doi_analysis.Detection
    analysis: object = None         # doi_analysis.Analysis


@dataclass
class SentMail:
    """An email YOU sent (Sent Items): acknowledgments, status updates, response materials,
    scope questions - and the date that feeds Last Email Sent."""
    subject: str
    sent_on: datetime
    conversation_id: str = ""
    entry_id: str = ""
    body: str = ""


def match_sent(sent, accepted) -> list:
    """[(SentMail, request_id)] for sent emails about a request: the subject names the ID
    (e.g. 'RE: FIN-53 ...') or the email is in the same Outlook conversation as a request
    email (a reply whose subject was changed). Pure Python."""
    by_thread = {m.conversation_id: rid for m, rid in accepted if m.conversation_id}
    matched = []
    for s in sent:
        rid = find_id(s.subject) or by_thread.get(s.conversation_id)
        if rid and s.sent_on:
            matched.append((s, rid))
    return matched


def last_sent_by_request(sent, accepted) -> dict:
    """Request ID -> newest date you emailed about it (see match_sent)."""
    latest = {}
    for s, rid in match_sent(sent, accepted):
        if rid not in latest or s.sent_on > latest[rid]:
            latest[rid] = s.sent_on
    return latest


def last_exchange_email(mails, accepted, known_threads=None) -> dict:
    """Request ID -> (newest datetime, source): the latest email in that request's exchange,
    WHOEVER sent it - examiner, accountant, a coworker, anyone - among everything fetched from
    your mailbox: accepted, review AND rejected (a short 'RE: FIN-53 - thanks' isn't a request,
    but it is still the latest email in the exchange). A message belongs to a request when its
    subject names the ID, its detection found the ID, or it is in the same Outlook conversation
    as an accepted request email (a reply whose subject was changed) - this run's, or one
    remembered from earlier runs (known_threads, from seen_state). Pure Python."""
    by_thread = dict(known_threads or {})
    by_thread.update({m.conversation_id: rid for m, rid in accepted if m.conversation_id})
    latest = {}
    for m in mails:
        if not m.received:
            continue
        det_id = getattr(m.detection, "request_id", "") if m.detection is not None else ""
        rid = find_id(m.subject) or det_id or by_thread.get(m.conversation_id)
        if rid and (rid not in latest or m.received > latest[rid][0]):
            latest[rid] = (m.received, f"latest email in the exchange, from {m.sender}")
    return latest


def merge_last_sent(*sources) -> dict:
    """Combine {rid: datetime or (datetime, source)} dicts, keeping the newest per request."""
    merged = {}
    for source in sources:
        for rid, value in (source or {}).items():
            when, label = value if isinstance(value, tuple) else (value, "your Sent Items")
            if rid not in merged or when > merged[rid][0]:
                merged[rid] = (when, label)
    return merged


def message_ref(entry_id: str) -> str:
    """Short, stable tag for one email ('ref 1af3ac90'): 8 hex characters of a hash of Outlook's
    EntryID. Used in the tracker's Notes, the log and the seen-messages memory."""
    return "ref " + hashlib.sha1((entry_id or "").encode()).hexdigest()[:8]


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


def attachment_paths(rid: str, names, received: datetime, folder) -> list:
    """Every attachment this program would keep, saved or not: [(outlook_index, path)]."""
    plan = []
    for index, name in enumerate(names, start=1):           # Outlook's Attachments are 1-based
        if not name.lower().endswith(ATTACH_TYPES) or INLINE_IMAGE.match(name):
            continue
        safe = INVALID_CHARS.sub("_", name).strip(" .")
        stem, dot, ext = safe.rpartition(".")
        if dot and len(stem) > 100:                          # keep under Windows' 260-char path limit
            safe = f"{stem[:100]}.{ext}"
        dest = Path(folder) / rid / f"{received:%Y-%m-%d}_{safe}"
        if dest not in [d for _, d in plan]:
            plan.append((index, dest))
    return plan


def attachment_targets(rid: str, names, received: datetime, folder) -> list:
    """Which attachments to save, and where: [(outlook_index, destination), ...].

    Every attachment of an ATTACH_TYPES type is kept (the whole email is request evidence),
    except Outlook's auto-named inline images (image001.png = signature logos).

    ATTACH_DIR / <RequestID> / <YYYY-MM-DD>_<file name> - one folder per request, the received date
    in front so a revised 'Request.pdf' on a later email doesn't collide with the first one.
    Already-saved files are skipped, so re-runs don't duplicate. Pure Python - no Outlook."""
    return [(i, dest) for i, dest in attachment_paths(rid, names, received, folder) if not dest.exists()]


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
        conversation_id=getattr(msg, "ConversationID", "") or "",
    )


def extract_sent(msg) -> SentMail:
    """COM -> SentMail (Sent Items)."""
    return SentMail(subject=msg.Subject or "", sent_on=_naive(msg.SentOn),
                    conversation_id=getattr(msg, "ConversationID", "") or "",
                    entry_id=getattr(msg, "EntryID", "") or "", body=getattr(msg, "Body", "") or "")


# --- the rules -------------------------------------------------------------
def qualify(m: Mail, approved_senders=()) -> tuple[bool, str]:
    """Return (ok, reason). Every rejection names the rule that failed - that is your audit trail.
    approved_senders: individual addresses you approved in the contact directory."""
    rid = find_id(m.subject)
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
    return True, rid


# --- Outlook side ----------------------------------------------------------
def _outlook():
    import win32com.client
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def fetch_sent(since: datetime) -> list:
    """Your Sent Items since a date. No subject filter on purpose: a reply in a request's
    thread counts even if its subject no longer shows the ID (matched by ConversationID)."""
    items = _outlook().GetDefaultFolder(5).Items                 # 5 = Sent Items
    items = items.Restrict(f"[SentOn] >= '{since:%m/%d/%Y} 12:00 AM'")
    return [extract_sent(msg) for msg in items if msg.Class == 43]


def read_attachments(msg, m: Mail, rid: str, attach_dir=None) -> list:
    """Save (if attach_dir) and read the email's attachments. Returns save results for the log;
    fills m.attachment_texts / m.attachment_status. Without attach_dir the files go to a temp
    folder just long enough to be read, so a dry run sees the same content as a real run."""
    import tempfile
    from attachment_text import read_text
    work = attach_dir or tempfile.mkdtemp(prefix="doi_attach_")
    saved = save_attachments(msg, m, rid, work)
    for _, path in attachment_paths(rid, m.attachments, m.received, work):
        if path.exists():
            text, status = read_text(path)
            m.attachment_status.append(f"{path.name}: {status}")
            if text:
                m.attachment_texts[path.name] = text
    if not attach_dir:
        import shutil
        shutil.rmtree(work, ignore_errors=True)
        return []
    return saved


def inbox_folders(folder):
    """The Inbox and every folder inside it, however deep - so emails count wherever your
    Outlook rules (or you) put them. A folder that can't be listed is skipped, not fatal."""
    yield folder
    try:
        children = list(folder.Folders)
    except Exception:
        children = []
    for child in children:
        yield from inbox_folders(child)


def fetch_candidates(since: datetime, attach_dir=None, approved_senders=(), skip_refs=()):
    """Coarse filter inside Outlook, then the Python decision (doi_analysis.detect), over the
    whole Inbox tree. Returns (accepted, rejected, review):
      accepted [(Mail, request_id)]  rejected [(Mail, reason)]  review [(Mail, Detection)].
    skip_refs: emails already fully processed on an earlier run (see seen_state) - not
    opened again. The same email in two folders is handled once. Attachments of anything not
    rejected are saved/read here - the Outlook item is only reachable in this function - and
    the email is judged again with their text."""
    from doi_analysis import detect                              # pure module; imports this one
    accepted, rejected, review = [], [], []
    handled = set()
    for folder in inbox_folders(_outlook().GetDefaultFolder(6)):     # 6 = Inbox
        try:
            items = folder.Items.Restrict(f"[ReceivedTime] >= '{since:%m/%d/%Y} 12:00 AM'")   # Jet
            items = items.Restrict(subject_filter())                                        # DASL
            messages = list(items)
        except Exception:                                    # e.g. a folder of calendar items
            continue
        for msg in messages:
            try:
                if msg.Class != 43:                          # olMail only
                    continue
                entry_id = msg.EntryID
            except Exception:
                continue
            if entry_id in handled or message_ref(entry_id) in skip_refs:
                continue
            handled.add(entry_id)
            accepted, rejected, review = _judge(msg, detect, attach_dir, approved_senders,
                                                accepted, rejected, review)
    return accepted, rejected, review


def _judge(msg, detect, attach_dir, approved_senders, accepted, rejected, review):
    """Detect one Outlook email (again once its attachments are read) and file it in a list.
    An email Outlook can't hand over (e.g. still downloading) is skipped this run, not fatal."""
    try:
        m = extract(msg)
    except Exception as e:
        print(f"  skipped an email Outlook could not read: {e}")
        return accepted, rejected, review
    first = detect(m, approved_senders)
    if first.status != "reject" and m.attachments:
        m.saved = read_attachments(msg, m, first.request_id or "_Review", attach_dir)
    m.detection = det = detect(m, approved_senders)          # again, now knowing the attachments
    if det.status == "accept":
        accepted.append((m, det.request_id))
    elif det.status == "review":
        review.append((m, det))
    else:
        rejected.append((m, det.reasons[0]))
    return accepted, rejected, review
