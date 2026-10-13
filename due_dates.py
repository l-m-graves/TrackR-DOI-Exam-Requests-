"""Find a request's due date in an email body - conservatively.

A date is only returned when the email states exactly ONE due date near due-date wording
("due", "respond by", "no later than", "provide ... by", "deadline" ...). Anything less
clear comes back as a note for a person to read instead of a guess in the Due Date column:
  - several different dates          -> note listing them
  - 'ASAP' / 'TBD' / 'as soon as possible' -> note (keeps column G numeric)
Quoted earlier messages (the 'From: ... Sent: ...' history under a reply) are cut off
first, so an old deadline further down the thread is never picked up.
"""
import re
from datetime import date, datetime, timedelta

QUOTE_START = re.compile(r"^\s*(-{2,}\s*Original Message|_{10,}\s*$|From:\s|On .{5,80} wrote:\s*$)",
                         re.I | re.M)
CUE = re.compile(r"\b(due(?: date)?|deadline|respond|responses?|reply|no later than|on or before|"
                 r"submit(?:ted)?\s+by|provid(?:e|ing)\b[^.\n]{0,60}?\bby|return(?:ed)?\s+by)\b", re.I)
# Status updates promise a delivery date - that is NOT the DOI's due date, so it has its own cue.
EXPECT_CUE = re.compile(r"\bexpect(?:s|ed)?\b[^.\n]{0,80}?\bby\b", re.I)
WINDOW = 100                                     # characters after the cue that are searched
MONTHS = {m: i for i, names in enumerate(
    [("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"), ("may",),
     ("jun", "june"), ("jul", "july"), ("aug", "august"), ("sep", "sept", "september"),
     ("oct", "october"), ("nov", "november"), ("dec", "december")], start=1) for m in names}
NUMERIC = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4}|\d{2})\b")          # 10/15/2027, 10-15-27
ISO = re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b")                            # 2027-10-15
WORDY = re.compile(r"\b(" + "|".join(sorted(MONTHS, key=len, reverse=True)) +
                   r")\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?\b", re.I)  # October 15, 2027
TEXT_DUE = re.compile(r"\b(ASAP|as soon as possible|TBD|to be determined)\b", re.I)


def strip_quoted(body: str) -> str:
    """Only the new part of a reply - everything above the first quoted-history marker."""
    m = QUOTE_START.search(body or "")
    return (body or "")[:m.start()] if m else (body or "")


def _dates_in(text, received: datetime) -> list:
    found = []
    for m in ISO.finditer(text):
        found.append((int(m[1]), int(m[2]), int(m[3])))
    for m in NUMERIC.finditer(ISO.sub(" ", text)):
        year = int(m[3]) + (2000 if len(m[3]) == 2 else 0)
        found.append((year, int(m[1]), int(m[2])))
    for m in WORDY.finditer(text):
        month, day = MONTHS[m[1].lower()], int(m[2])
        year = int(m[3]) if m[3] else received.year
        if not m[3] and date(year, month, min(day, 28)) < received.date() - timedelta(days=1):
            year += 1                                  # 'January 5' written in December = next year
        found.append((year, month, day))
    out = []
    for y, mo, d in found:
        try:
            dt = date(y, mo, d)
        except ValueError:                             # 13/45/2027 etc. - not a date
            continue
        # a due date is after (or the day before) the email, and within a year of it
        if received.date() - timedelta(days=1) <= dt <= received.date() + timedelta(days=365):
            out.append(dt)
    return out


def find_due(body: str, received: datetime, cue_pattern=CUE) -> dict:
    """{'date': date or None, 'note': str or None}. Never both a date and a doubt."""
    text = strip_quoted(body)
    dates, texts = [], []
    for cue in cue_pattern.finditer(text):
        window = text[cue.start(): cue.end() + WINDOW].split("\n\n")[0]
        dates += _dates_in(window, received)
        texts += [t.group(0) for t in TEXT_DUE.finditer(window)]
    distinct = sorted(set(dates))
    if len(distinct) == 1:
        return {"date": distinct[0], "note": None}
    if len(distinct) > 1:
        listed = ", ".join(f"{d:%Y-%m-%d}" for d in distinct)
        return {"date": None, "note": f"several due dates in email ({listed}) - Due Date left for you"}
    if texts:
        return {"date": None, "note": f"due: {texts[0]}"}
    return {"date": None, "note": None}


def find_expected(body: str, sent: datetime):
    """The date a status update promises ('We expect to provide the remaining materials by ...'),
    or None. Goes into Notes, never into Due Date."""
    return find_due(body, sent, EXPECT_CUE)["date"]
