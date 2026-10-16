"""What a DOI email is, who it really comes from, and what it says - pure Python, no Outlook.

detect()  -> Detection: accept / review / reject, a confidence, and the reasons (signals).
analyze() -> Analysis:  dates by who stated them, request items, scope, departments, events,
                        follow-up link, people involved, and anything that needs a human.

Who is who comes only from the role settings in doi_filter / local_settings.py (role_of).
The person running this is usually CC'd: nothing here assumes they sent, or must act on, a message.
"""
import re
from dataclasses import dataclass, field

import doi_filter as F
from due_dates import EXPECT_CUE, find_due
from email_types import classify, topic_from_subject

# --- forwarded / quoted parts -------------------------------------------------------------
HEADER_FROM = re.compile(r"^[ \t>]*From:[ \t]*(.+)$", re.I | re.M)
HEADER_FIELD = re.compile(r"^[ \t>]*(Sent|Date|To|Cc|Subject):", re.I | re.M)
MARKER = re.compile(r"^[ \t>]*(-{2,}\s*(Original|Forwarded) Message\s*-*|_{10,}|Begin forwarded message:)\s*$",
                    re.I | re.M)
ADDRESS = re.compile(r"[\w.+'-]+@[\w-]+(?:\.[\w-]+)+")
X500 = re.compile(r"/O=", re.I)          # Exchange's internal address form - only your own company's people have one


@dataclass
class Segment:
    author: str            # address ('' if the header shows only a name nobody on the email matches)
    author_name: str
    text: str
    exchange: bool = False # header shows an Exchange (X500) address: someone inside the company
    recipients: list = field(default_factory=list)   # addresses on that part's To:/Cc: lines


def _tokens(name) -> frozenset:
    """'Lee, Jordan' and 'Jordan Lee' -> the same key."""
    return frozenset(t for t in re.findall(r"[a-z]+", (name or "").lower()) if len(t) > 1)


def name_book(m) -> dict:
    """Display name -> address for the people on this email (sender, To, CC). Classic Outlook
    often quotes colleagues as 'From: Lastname, Firstname' with no address; this finds the
    address when that person is also on the email. Names shared by two addresses are dropped."""
    book, clash = {}, set()
    pairs = list((getattr(m, "names", None) or {}).items()) + [(m.sender, getattr(m, "sender_name", ""))]
    for addr, name in pairs:
        key = _tokens(name)
        if not key or not addr:
            continue
        if book.get(key, addr) != addr:
            clash.add(key)
        book[key] = addr
    for key in clash:
        del book[key]
    return book


def segments(body: str, sender: str, names=None) -> list:
    """Split a message into who-wrote-what, newest first. Segment 0 is the current sender's
    own words; each later one starts at a quoted/forwarded 'From: ... Sent/Date: ...' header.
    names: name_book() of the email, to find the address of a name-only 'From:' line."""
    body = body or ""
    cuts = []
    for m in HEADER_FROM.finditer(body):
        lookahead = body[m.end(): m.end() + 400]
        if len(HEADER_FIELD.findall(lookahead)) >= 2:          # From: + Sent:/To:/Subject: = a header
            marker = None
            for mk in MARKER.finditer(body, 0, m.start()):
                marker = mk
            start = marker.start() if marker and body[marker.end():m.start()].strip() == "" else m.start()
            cuts.append((start, m))
    parts = [Segment(sender or "", "", body[: cuts[0][0]] if cuts else body)]
    for i, (start, m) in enumerate(cuts):
        end = cuts[i + 1][0] if i + 1 < len(cuts) else len(body)
        chunk = body[m.end(): end]
        fields = list(HEADER_FIELD.finditer(chunk))
        header_end = chunk.find("\n", fields[-1].end()) + 1 if fields else 0
        text = chunk[header_end:] if fields else chunk                                 # after the header
        recipients = [a.lower() for line in chunk[:header_end].splitlines()
                      if re.match(r"[ \t>]*(To|Cc):", line, re.I) for a in ADDRESS.findall(line)]
        addr = ADDRESS.search(m.group(1))
        name = re.sub(r"[<\[].*", "", m.group(1)).strip().strip('"')
        author = addr.group(0).lower() if addr else (names or {}).get(_tokens(name), "")
        parts.append(Segment(author, name, text, exchange=bool(X500.search(m.group(1))), recipients=recipients))
    return parts


def examiner_included(m, approved_senders=()) -> set:
    """Outside addresses the EXAMINER SIDE put on this thread: To/CC of an email an examiner-side
    person sent, and the To/Cc lines of a forwarded part they wrote. Only people the program
    doesn't know yet (not internal, not you, not the team mailbox). Being brought in by the
    examiner is the evidence; being on an email from a colleague is not (that could be the
    company's own outside consultant)."""
    found = set()
    if F.role_of(m.sender, approved_senders) in F.EXAMINER_SIDE:
        found |= set(m.to) | set(m.cc)
    for seg in segments(m.body, m.sender, name_book(m))[1:]:
        if seg.author and F.role_of(seg.author, approved_senders) in F.EXAMINER_SIDE:
            found |= set(seg.recipients)
    return {a for a in found if "@" in a and a not in F.MY_ADDRESSES and a not in F.MUST_BE_SENT_TO
            and F.role_of(a, approved_senders) in ("unknown", "approved")}


def original_sender(segs) -> str:
    """Who started the thread: the oldest quoted/forwarded author with an address, or ''."""
    for seg in reversed(segs[1:]):
        if seg.author:
            return seg.author
    return ""


# --- detection --------------------------------------------------------------------------
DOI_TERMS = re.compile(r"\b(DOI|Department of Insurance|insurance department|examination|examiners?|"
                       r"exam request|market conduct|financial exam|regulators?)\b", re.I)


@dataclass
class Detection:
    status: str                    # accept | review | reject
    confidence: str                # high | medium | low
    request_id: str = ""
    reasons: list = field(default_factory=list)       # signals seen / why
    sender_role: str = "unknown"
    original_sender: str = ""
    original_role: str = ""


def _find_id(m, segs):
    """Request ID and where it was found: subject first, then the message, then attachments."""
    places = [("subject", m.subject)] + [(f"message part {i}", s.text) for i, s in enumerate(segs)]
    places += [(f"attachment {n}", t) for n, t in m.attachment_texts.items()]
    for where, text in places:
        rid = F.find_id(text)
        if rid:
            return rid, where
    return "", ""


def detect(m, approved_senders=()) -> Detection:
    """Accept, review or reject one received email, with the signals behind the decision.

    The original five rules (qualify) still accept on their own. Beyond them, an email is judged
    on several signals - the CURRENT sender's role, the ORIGINAL sender inside a forward, DOI
    wording, a request ID anywhere (subject, body, attachment) - and anything uncertain goes to
    review instead of being silently accepted or dropped."""
    segs = segments(m.body, m.sender, name_book(m))
    orig = original_sender(segs)
    det = Detection("reject", "low", sender_role=F.role_of(m.sender, approved_senders),
                    original_sender=orig, original_role=F.role_of(orig, approved_senders) if orig else "")
    rid, where = _find_id(m, segs)
    det.request_id = rid
    text = " ".join([m.subject or ""] + [s.text for s in segs] + list(m.attachment_texts.values()))
    doi_words = bool(DOI_TERMS.search(text))
    examiner_origin = det.sender_role in F.EXAMINER_SIDE or det.original_role in F.EXAMINER_SIDE
    internal_sender = det.sender_role in F.INTERNAL_SIDE
    to_examiner = any(F.role_of(a, approved_senders) in F.EXAMINER_SIDE for a in list(m.to) + list(m.cc))
    det.reasons = [f"sender {m.sender} = {det.sender_role}"]
    if orig:
        det.reasons.append(f"original sender {orig} = {det.original_role}")
    if rid:
        det.reasons.append(f"request ID {rid} in {where}")
    if doi_words:
        det.reasons.append("DOI/examination wording")

    ok, why = F.qualify(m, approved_senders)
    if ok:
        det.status, det.confidence = "accept", "high"
        det.reasons.insert(0, "passes all five qualification rules")
        return det
    if F.MUST_BE_SENT_TO and not internal_sender and not F.reaches_required(m):
        # An explicit setting is a hard rule for EXTERNAL emails only. Internal forwards and
        # replies (often just to you, CC'd) are judged on the signals below instead.
        det.reasons.insert(0, "external email not addressed to MUST_BE_SENT_TO")
        return det
    if not rid:
        if examiner_origin and doi_words:
            det.status, det.confidence = "review", "medium"
            det.reasons.insert(0, "from the examiner side with DOI wording, but no request ID")
        else:
            det.reasons.insert(0, "no request ID and not from the examiner side - not actionable")
        return det
    if examiner_origin and doi_words:
        det.status, det.confidence = "accept", "high"
        det.reasons.insert(0, "examiner-originated request" + (" (forwarded)" if det.original_role in F.EXAMINER_SIDE and det.sender_role not in F.EXAMINER_SIDE else ""))
    elif internal_sender and (doi_words or to_examiner or (F.find_id(m.subject) and DELIVERED.search(segs[0].text))):
        det.status, det.confidence = "accept", "medium"
        det.reasons.insert(0, "internal request about a DOI exam item" if doi_words
                           else "internal email to the examiner side about a request" if to_examiner
                           else "internal email saying a request's support was delivered")
    elif doi_words:
        det.status, det.confidence = "review", "low"
        det.reasons.insert(0, f"request ID and DOI wording, but sender not in the role settings: {m.sender}")
    else:
        det.reasons.insert(0, f"'{rid}' looks like an ID but nothing else says DOI exam - not actionable")
    return det


# --- analysis ---------------------------------------------------------------------------
EVENTS = {
    "follow_up": r"\bfollow[- ]?up\b",
    "clarification": r"\bclarif(?:y|ication)\b|\bconfirm (?:the )?scope\b",
    "missing_support": r"\b(?:missing|outstanding|still need|not (?:yet )?received|additional (?:support|documentation|information|explanation))\b",
    "extension": r"\bextension\b|\bextend(?:ed)? the (?:due date|deadline)\b",
    "meeting": r"\b(?:interview|meeting|walk-?through)\b",
    "submission_instructions": r"\b(?:SharePoint|SFTP|secure (?:file|portal|site)|upload (?:to|via|through)|file size|too large)\b",
    "files_ready": r"\b(?:files?|materials?|documents?|responses?) (?:are|is) ready\b|\bready for (?:review|submission|upload)\b",
    "submitted": r"\b(?:uploaded|submitted|delivered|provided) (?:to|for) the (?:examiners?|DOI|regulators?|examination team)\b|\bhave been (?:uploaded|submitted)\b",
}
DEPARTMENTS = re.compile(r"\b(Payroll|Investments|Legal|Underwriting|Claims|Actuarial|Tax|Reinsurance|Treasury|"
                         r"Human Resources|HR|Compliance|Internal Audit|Information Technology|IT(?= (?:department|team|group)))\b")
STATES = ("Alabama Alaska Arizona Arkansas California Colorado Connecticut Delaware Florida Georgia Hawaii Idaho "
          "Illinois Indiana Iowa Kansas Kentucky Louisiana Maine Maryland Massachusetts Michigan Minnesota "
          "Mississippi Missouri Montana Nebraska Nevada Ohio Oklahoma Oregon Pennsylvania Tennessee Texas Utah "
          "Vermont Virginia Washington Wisconsin Wyoming").split() + [
          "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota",
          "Rhode Island", "South Carolina", "South Dakota", "West Virginia", "District of Columbia"]
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?"
PERIOD = re.compile(rf"\b(?:FY ?\d{{4}}|Q[1-4] ?\d{{4}}|(?:calendar|fiscal|plan|policy) year \d{{4}}|year[- ]end \d{{4}}|"
                    rf"{MONTH} \d{{1,2}},? \d{{4}} (?:to|through|-|–) {MONTH} \d{{1,2}},? \d{{4}}|"
                    rf"\d{{1,2}}/\d{{1,2}}/\d{{4}} (?:to|through|-|–) \d{{1,2}}/\d{{1,2}}/\d{{4}}|"
                    rf"(?:19|20)\d{{2}} ?(?:-|–|to|through) ?(?:19|20)\d{{2}})\b", re.I)
CONTROL = re.compile(r"\bcontrols? (?:#|no\.? ?|number )?([A-Z]{1,6}[-.]?\d+(?:\.\d+)*)\b|\b((?:ITGC|SOX)[- ]?\d+(?:\.\d+)*)\b")
ITEM = re.compile(r"^[ \t]*(?:\(?(\d{1,2})[.)]|Item (\d{1,2})[:.)]?|Question (\d{1,2})[:.)]?)[ \t]+(\S.{2,})$", re.I | re.M)


@dataclass
class Analysis:
    request_id: str
    kind: str = "other"                    # email_types template type (one signal, not the whole story)
    request_type: str = "initial"          # initial | follow_up | clarification | extension | submission
    subject_area: str = ""
    parent_request_id: str = ""
    events: list = field(default_factory=list)
    request_items: list = field(default_factory=list)
    departments: list = field(default_factory=list)
    states: list = field(default_factory=list)
    periods: list = field(default_factory=list)
    controls: list = field(default_factory=list)
    entities: list = field(default_factory=list)
    products: list = field(default_factory=list)
    current_sender: str = ""
    original_sender: str = ""
    coordinator: str = ""                  # internal accountant managing it
    external_contact: str = ""             # examiner / exam-firm person
    asked: list = field(default_factory=list)            # who is being asked (To) when internal assigns
    exam_due_date: object = None
    internal_expected_date: object = None
    tracking_due_date: object = None
    date_source: str = ""
    date_confidence: str = ""
    date_evidence: str = ""
    date_notes: list = field(default_factory=list)       # ASAP / several dates / unverified dates
    review: list = field(default_factory=list)           # reasons a person must look
    submission_status: str = ""                          # files ready | submitted
    completes: bool = False                # delivered by a SUBMITTER to the lead examiner -> Completed
    submitted_only: bool = False           # delivery wording, but not (yet) from a submitter to the lead -> Submitted
    not_complete_because: str = ""         # why it is Submitted rather than Completed
    partial_submission: bool = False       # only part of it ("remaining items will follow")
    submission_evidence: str = ""          # the sentence that says so
    submitted_to: str = ""                 # the examiner-side recipients of that email
    send_to: str = ""                      # who the response goes to: 'Name <address>' (or a name not found)
    send_to_source: str = ""               # 'stated: <sentence>' | 'lead examiner (default ...)'

    def summary(self) -> str:
        """Short text for the tracker's Notes line."""
        bits = []
        if self.exam_due_date:
            bits.append(f"exam due {self.exam_due_date:%Y-%m-%d} ({self.date_source})")
        if self.internal_expected_date:
            bits.append(f"internal expected {self.internal_expected_date:%Y-%m-%d} - not an examiner deadline")
        bits += self.date_notes
        if self.parent_request_id:
            bits.append(f"follows up on {self.parent_request_id}")
        if self.events:
            bits.append("events: " + ", ".join(self.events))
        if self.request_items:
            bits.append(f"{len(self.request_items)} request items")
        for label, values in (("dept", self.departments), ("states", self.states), ("periods", self.periods),
                              ("controls", self.controls), ("entities", self.entities), ("products", self.products)):
            if values:
                bits.append(f"{label}: {', '.join(values)}")
        if self.send_to and self.send_to_source.startswith("stated"):
            bits.append(f"send to {self.send_to}")             # the default is noted once, on the new row
        if self.review:
            bits.append("REVIEW: " + "; ".join(self.review))
        return "; ".join(bits)


# --- who the response goes to -----------------------------------------------------------
# "Please send to Jordan by 10/20/27", "please give it to me by ...", "submit the responses to
# jordan.lee@... ". Only the EXAMINER SIDE's words (or a request document) say where the DOI wants
# the response; a colleague's "send it to me" is internal collection, not the delivery address.
# Only these shapes: "<verb> [it / them / the responses / your documents ...] [directly] to <who>".
# A fixed list on purpose - "information related to Premium Tax" or "reconcile Schedule D to the
# General Ledger" must not read as "send to ...".
SEND_WHAT = (r"(?:\s+(?:it|them|this|these|those|everything|"
             r"(?:the|your|all|any|our|all\s+the)\s+(?:requested\s+)?(?:responses?|documents?|documentation|files?|"
             r"materials?|information|support(?:ing\s+documents?)?|items?|answers?|data|deliverables?|reports?)))?")
SEND_TO = re.compile(r"\b(?:send|give|provide|submit|forward|return|e-?mail|deliver|direct)" + SEND_WHAT +
                     r"\s+(?:directly\s+|back\s+)?to\s+(?P<who>[^,;:!?\n()]{1,80})", re.I)
WHO_END = re.compile(r"\s+(?:by|no later than|on or before|before|at|via|on|in|within|for|and|so|as|once|when|if)\b.*$", re.I)
NOT_A_PERSON = re.compile(r"^(?:the|our|your|this|these|those|that|all|each|any|every|a|an|following|"
                          r"department|division|team|examiners?|regulators?|doi|sharepoint|portal|site)\b", re.I)
NAME = re.compile(r"^[A-Z][\w'.-]*(?:,?\s+[A-Z][\w'.-]*){0,3}$")


def _resolve_name(name, book):
    """'Jordan', 'Jordan Lee' or 'Lee, Jordan' -> the one address on the email whose name contains
    those words, else '' (none, or more than one fits)."""
    want = _tokens(name)
    hits = {addr for key, addr in book.items() if want and want <= key}
    return hits.pop() if len(hits) == 1 else ""


def _display(addr, m):
    name = (m.names or {}).get(addr) or (m.sender_name if addr == m.sender else "")
    return f"{name} <{addr}>" if name else addr


# --- delivered: "FIN-52 support has been uploaded to OneDrive" -----------------------------
DELIVERED = re.compile(
    r"\b(?:has|have|had)\s+(?:now\s+|just\s+|all\s+)?been\s+(?:uploaded|submitted|provided|sent|posted|shared|delivered|placed|saved|loaded)\b"
    r"|\b(?:I|we)(?:'ve|\s+have|\s+just)?\s+(?:uploaded|submitted|provided|sent|posted|shared|delivered)\b"
    r"|\buploaded\s+(?:to|on|in|into)\s+(?:the\s+|our\s+|your\s+)?(?:OneDrive|SharePoint|portal|site|folder|SFTP|Box|Teams)\b"
    r"|\bplease\s+find\s+(?:attached|enclosed)\b|\battached\s+(?:is|are)\s+(?:the|our)\b"
    r"|\bcompleted\b|\bsubmitted\b|\buploaded\s+to\s+(?:the\s+)?OneDrive\b", re.I)
# "will be submitted", "has not been completed", "once uploaded" - said about the future, not done
NOT_YET = re.compile(r"\b(?:not|never)\s+(?:yet\s+)?(?:been\s+|be\s+)?(?:completed|submitted|uploaded|provided|sent)\b"
                     r"|\b(?:will|to|should|must|can|could|would|need\s+to)\s+(?:be\s+|have\s+been\s+)?"
                     r"(?:completed|submitted|uploaded|provided|sent)\b"
                     r"|\b(?:once|when|until|before|after|if)\s+(?:it\s+is\s+|it's\s+|they\s+are\s+|we\s+have\s+)?"
                     r"(?:completed|submitted|uploaded)\b|\bcompleted\?|\bsubmitted\?", re.I)
FORWARD = re.compile(r"^\s*(?:fw|fwd)\s*:", re.I)
PARTIAL = re.compile(r"\b(?:partial(?:ly)?|remaining|remainder|rest\s+of|still\s+(?:working|outstanding|need|pending|gathering)|"
                     r"(?:will|to)\s+follow|outstanding|part\s+of|portion\s+of|some\s+of|first\s+(?:batch|set|part)|"
                     r"not\s+yet|in\s+progress|by\s+(?:monday|tuesday|wednesday|thursday|friday|end\s+of))\b", re.I)


def _sentence_around(text, start, end):
    s = max(text.rfind("\n", 0, start), text.rfind(". ", 0, start)) + 1
    ends = [i for i in (text.find("\n", end), text.find(". ", end)) if i != -1]
    return " ".join(text[s: min(ends) + 1 if ends else len(text)].split())[:200]


def find_submission(m, segs, roles) -> dict:
    """Did someone inside the company say this request's support was delivered?
      status 'completed'  - a SUBMITTER said so with a lead examiner on To/CC
             'submitted'  - delivery wording from anyone else inside the company (or the lead
                            examiner not on the email): Submitted, not Completed
             'partial'    - only part of it ('the remaining items will follow')
             ''           - nothing of the kind
    Only that person's own new words count - never quoted history - and the subject must name
    the request. You are an observer: being on To/CC makes you neither sender nor recipient.
    For a forward (FW:), the forwarded original is judged as the real email."""
    none = {"status": "", "evidence": "", "to": "", "why": ""}
    if not F.find_id(m.subject):
        return none
    parts = [(m.sender, list(m.to) + list(m.cc), segs[0].text, roles[0])]
    if FORWARD.match(m.subject or ""):
        parts += [(s.author, s.recipients, s.text, r) for s, r in zip(segs[1:], roles[1:]) if s.author]
    for author, recipients, text, role in parts:
        if role not in F.INTERNAL_SIDE or author in F.MY_ADDRESSES:
            continue
        hit = DELIVERED.search(text)
        if not hit:
            continue
        evidence = _sentence_around(text, hit.start(), hit.end())
        if NOT_YET.search(evidence):
            continue
        if PARTIAL.search(text):
            return {**none, "status": "partial", "evidence": evidence}
        people = [a for a in recipients if a not in F.MY_ADDRESSES]
        leads = [a for a in people if (a in F.LEAD_EXAMINERS if F.LEAD_EXAMINERS else F.role_of(a) in F.EXAMINER_SIDE)]
        submitter = author in F.SUBMITTERS if F.SUBMITTERS else True
        to = "; ".join(_display(a, m) for a in (leads or people))
        if submitter and leads:
            return {"status": "completed", "evidence": evidence, "to": to, "why": ""}
        why = ("sent by someone who isn't a designated submitter" if not submitter
               else "the lead examiner isn't on the email")
        return {"status": "submitted", "evidence": evidence, "to": to, "why": why}
    return none


def find_send_to(m, segs, roles, trusted_docs=()):
    """(send_to, source) from the examiner side's own words, newest first, then request documents.
    'me'/'us' = whoever wrote that part. A name is looked up among the people on the email."""
    book = name_book(m)
    texts = [(seg.author or (m.sender if i == 0 else ""), seg.text)
             for i, (seg, role) in enumerate(zip(segs, roles)) if role in F.EXAMINER_SIDE]
    texts += [("", text) for text in trusted_docs]
    for author, text in texts:
        for hit in SEND_TO.finditer(text):
            who = re.split(r"\.(?:\s|$)", hit["who"])[0]          # sentence end (dots inside an address stay)
            who = WHO_END.sub("", who).strip()
            sentence = " ".join(text[max(0, hit.start() - 10): hit.end() + 40].split())[:160]
            address = ADDRESS.search(who)
            if address:
                return _display(address.group(0).lower(), m), f"stated: {sentence}"
            if who.lower() in ("me", "us", "myself"):
                if author:
                    return _display(author, m), f"stated: {sentence}"
                continue
            if NOT_A_PERSON.match(who) or not NAME.match(who):
                continue
            addr = _resolve_name(who, book)
            return (_display(addr, m) if addr else f"{who} (address not on the email)"), f"stated: {sentence}"
    leads = [a for a in F.LEAD_EXAMINERS]
    on_email = [a for a in leads if a in {m.sender, *m.to, *m.cc, *(s.author for s in segs)}]
    if on_email or len(leads) == 1:
        return _display((on_email or leads)[0], m), "lead examiner (default - the email names no one)"
    return "", ""


def _unique(items):
    out = []
    for i in items:
        if i and i not in out:
            out.append(i)
    return out


def analyze(m, request_id, approved_senders=()) -> Analysis:
    segs = segments(m.body, m.sender, name_book(m))
    orig = original_sender(segs)
    an = Analysis(request_id, kind=classify(m.subject, m.body), current_sender=m.sender, original_sender=orig,
                  subject_area=topic_from_subject(m.subject, request_id) or "")
    roles = [F.role_of(s.author or (m.sender if i == 0 else ""), approved_senders) for i, s in enumerate(segs)]
    roles = ["internal" if r == "unknown" and s.exchange else r for s, r in zip(segs, roles)]

    def who(i, seg):                                   # for the log: who stated a date
        if i == 0:
            return m.sender
        return seg.author or (f"forwarded part from {seg.author_name}" if seg.author_name else "a forwarded part")
    examiner_text = [s.text for s, r in zip(segs, roles) if r in F.EXAMINER_SIDE]
    own_and_examiner = [segs[0].text] + examiner_text            # substance, without old internal chatter
    attachments = list(m.attachment_texts.items())
    scope_text = "\n".join(own_and_examiner + [t for _, t in attachments])

    # people
    if roles[0] in F.INTERNAL_SIDE:
        an.coordinator = m.sender
        an.asked = list(m.to)
    for addr, role in [(m.sender, roles[0]), (orig, F.role_of(orig, approved_senders))]:
        if role in F.EXAMINER_SIDE and addr:
            an.external_contact = addr
            break

    # dates - by WHO stated them. Examiner side -> exam due date; internal -> internal expected.
    exam, internal, unverified, evidence = {}, {}, {}, {}
    for i, (seg, role) in enumerate(zip(segs, roles)):
        source = "examiner" if i == 0 else "examiner (forwarded)"
        due = find_due(seg.text, m.received, strip=False)
        exp = find_due(seg.text, m.received, EXPECT_CUE, strip=False)
        if role in F.EXAMINER_SIDE:
            for d in due["dates"]:
                exam.setdefault(d, source)
                evidence.setdefault(d, due["evidence"] or seg.text.strip()[:200])
            if due["note"] and not due["dates"]:
                an.date_notes.append(due["note"] + " (examiner)")
        elif role in F.INTERNAL_SIDE:
            for d in set(due["dates"]) | set(exp["dates"]):
                internal.setdefault(d, f"internal ({who(i, seg)})")
                evidence.setdefault(d, (exp["evidence"] or due["evidence"]))
            if due["note"] and not due["dates"] and i == 0:
                an.date_notes.append(due["note"])
        else:
            for d in due["dates"]:
                unverified.setdefault(d, who(i, seg))
            if due["note"] and not due["dates"] and i == 0:
                an.date_notes.append(due["note"])
    request_doc_trusted = any(r in F.EXAMINER_SIDE for r in roles) or roles[0] in F.INTERNAL_SIDE
    for name, text in attachments:
        due = find_due(text, m.received, strip=False)
        if not due["dates"]:
            continue
        looks_like_request = bool(F.important_attachments([name]))
        target = exam if (request_doc_trusted and looks_like_request) else unverified
        for d in due["dates"]:
            if target is exam:
                exam.setdefault(d, f"attachment {name}")
                evidence.setdefault(d, due["evidence"])
            else:
                unverified.setdefault(d, f"attachment {name}")

    if len(exam) == 1:
        an.exam_due_date = next(iter(exam))
    elif len(exam) > 1:
        an.review.append("conflicting examiner dates " + ", ".join(f"{d:%Y-%m-%d} ({s})" for d, s in sorted(exam.items())))
    if len(internal) == 1:
        an.internal_expected_date = next(iter(internal))
    elif len(internal) > 1:
        an.review.append("several internal dates " + ", ".join(f"{d:%Y-%m-%d}" for d in sorted(internal)))
    if unverified:
        an.review.append("date(s) from a sender not in the role settings: "
                         + ", ".join(f"{d:%Y-%m-%d} ({s})" for d, s in sorted(unverified.items())))
    if an.exam_due_date:
        d = an.exam_due_date
        an.tracking_due_date, an.date_source, an.date_evidence = d, exam[d], evidence.get(d, "")
        an.date_confidence = "high" if exam[d] == "examiner" else "medium"
    elif an.internal_expected_date:
        d = an.internal_expected_date
        an.tracking_due_date, an.date_source, an.date_evidence = d, "internal expected (fallback)", evidence.get(d, "")
        an.date_confidence = "low"

    # what the message is about
    event_text = "\n".join([m.subject or ""] + own_and_examiner)
    an.events = [name for name, pattern in EVENTS.items() if re.search(pattern, event_text, re.I)]
    if "submitted" in an.events:
        an.submission_status = "submitted (confirmed in email)"
    elif "files_ready" in an.events:
        an.submission_status = "files ready - not confirmed as submitted"
    for hit in F.REQUEST_ID.finditer(event_text):
        near = event_text[max(0, hit.start() - 30): hit.end() + 30]
        if F.canonical_id(hit.group(0)) != request_id and re.search(r"follow[- ]?(?:up|s up)", near, re.I):
            an.parent_request_id = F.canonical_id(hit.group(0))
            break
    an.request_type = ("follow_up" if an.parent_request_id or "follow_up" in an.events else
                       "clarification" if "clarification" in an.events else
                       "extension" if "extension" in an.events else
                       "submission" if an.submission_status or "submission_instructions" in an.events else "initial")

    an.request_items = _unique(" ".join((g[3] or "").split())[:80] for g in ITEM.findall(scope_text))
    an.departments = _unique(d if d != "HR" else "Human Resources" for d in DEPARTMENTS.findall(scope_text))
    an.states = [s for s in STATES if re.search(rf"\b{s}\b", scope_text)]
    an.periods = _unique(" ".join(p.split()) for p in PERIOD.findall(scope_text))
    an.controls = _unique(a or b for a, b in CONTROL.findall(scope_text))
    an.entities = [e for e in F.ENTITY_NAMES if e and e.lower() in scope_text.lower()]
    an.products = [p for p in F.PRODUCT_NAMES if p and p.lower() in scope_text.lower()]
    for name in m.attachment_status:
        if "not read" in name:
            an.review.append(f"attachment {name}")
    sub = find_submission(m, segs, roles)
    an.completes, an.submitted_only, an.partial_submission = (sub["status"] == s for s in ("completed", "submitted", "partial"))
    an.submission_evidence, an.submitted_to, an.not_complete_because = sub["evidence"], sub["to"], sub["why"]
    docs = [t for n, t in attachments if request_doc_trusted and F.important_attachments([n])]
    an.send_to, an.send_to_source = find_send_to(m, segs, roles, docs)
    if "address not on the email" in an.send_to:
        an.review.append(f"send to {an.send_to}")
    return an
