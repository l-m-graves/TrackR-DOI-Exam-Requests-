"""Recognise the team's DOI email templates and say what each means for the tracker.

  assignment          'DOI Exam Request FIN-53 – Topic'  "The ... DOI has issued examination request ..."
  acknowledgment      'RE: DOI Exam Request FIN-53 – Topic'  "We received the DOI exam request ..."
  status_update       'DOI Exam Request FIN-53 – Status Update'  "We expect to provide ... by <date>"
  response_materials  'DOI Exam Request FIN-53 – Response Materials'  "Attached are the completed responses ..."
  scope_question      'Question Regarding DOI Exam Request FIN-53'  "... confirm the scope of ..."
  other               anything else (handled the original, generic way)

Only the new part of a reply is read (quoted history cut off), so a reply that quotes the
assignment text is not mistaken for a new assignment.
"""
import re

from due_dates import strip_quoted

REPLY = re.compile(r"^\s*((re|fw|fwd)\s*:\s*)+", re.I)
SUBJECT_TOPIC = r"DOI Exam Request\s+{rid}\s*[-–—:]\s*(.+?)\s*$"     # en/em dash or hyphen
NOT_TOPICS = {"status update", "response materials"}

LABEL = {"assignment": "Assignment", "acknowledgment": "Acknowledgment", "status_update": "Status update",
         "response_materials": "Response materials", "scope_question": "Scope question", "other": "email"}
CREATES_ROW = {"assignment", "other"}          # follow-up emails never invent a request row
DUE_DATE_KINDS = {"assignment", "acknowledgment", "other"}   # where a stated date IS the due date
ADVANCES_TO_IN_PROGRESS = {"acknowledgment", "status_update", "response_materials", "scope_question"}


def classify(subject: str, body: str) -> str:
    subject = subject or ""
    text = strip_quoted(body or "")
    if re.search(r"\bstatus update\b", subject, re.I):
        return "status_update"
    if re.search(r"\bresponse materials\b", subject, re.I):
        return "response_materials"
    if re.match(r"^\s*((re|fw|fwd)\s*:\s*)*question regarding\b", subject, re.I) or \
            re.search(r"\bconfirm the scope\b", text, re.I):
        return "scope_question"
    if re.search(r"\bwe (have )?received the (doi )?(exam )?request\b", text, re.I):
        return "acknowledgment"
    if re.search(r"\bhas issued (an )?examination request\b", text, re.I):
        return "assignment"
    return "other"


def topic_from_subject(subject: str, rid: str):
    """'RE: DOI Exam Request FIN-53 – Premium Tax' -> 'Premium Tax'; None if there's no topic."""
    m = re.search(SUBJECT_TOPIC.format(rid=re.escape(rid)), REPLY.sub("", subject or ""), re.I)
    if not m or m.group(1).lower() in NOT_TOPICS:
        return None
    return m.group(1)
