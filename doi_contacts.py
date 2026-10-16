"""The contact directory: everyone seen on DOI request emails, kept up to date by each run.

The script RECORDS people automatically (sender, To, CC, which requests, first/last seen).
It only TRUSTS a new sender after you type Y in the Approved column - automation proposes,
a person approves. Otherwise anyone who put 'FIN-53' in a subject line could make themselves
an approved sender.

Columns you own: Approved (Y = trust, N = never trust), Notes - the script never changes them.
Add contacts from team meetings by hand the same way. Columns the script owns: everything else.
Examiner_Thread = Y: the examiner side put this person on a request email, so they count as
third party (exam-firm staff, say) without being listed in the settings - N in Approved undoes it.
"""
import csv
from pathlib import Path

FIELDS = ["Email", "Name", "Domain", "Roles", "Request_IDs", "First_Seen", "Last_Seen", "Examiner_Thread",
          "Approved", "Notes"]
ROLE_ORDER = ["Sender", "To", "CC", "Forwarded"]     # Forwarded: only on the To/Cc of a quoted message
NO = ("n", "no")


def load(path) -> dict:
    """Email (lower-case) -> row. A missing file is an empty directory - it's created on save."""
    path = Path(path)
    if not path.exists():
        return {}
    with open(path, newline="", encoding="utf-8-sig") as f:       # -sig: Excel adds a BOM on save
        return {row["Email"].strip().lower(): {k: (row.get(k) or "").strip() for k in FIELDS}
                for row in csv.DictReader(f) if (row.get("Email") or "").strip()}


def approved_senders(contacts) -> set:
    return {email for email, row in contacts.items() if row["Approved"].lower() in ("y", "yes")}


def thread_third_parties(contacts) -> set:
    """People the examiner side brought onto a request thread (Examiner_Thread = Y) - treated as
    third party unless you typed N in Approved."""
    return {email for email, row in contacts.items()
            if row["Examiner_Thread"].lower() == "y" and row["Approved"].lower() not in NO}


def _merge(text, item, order=None):
    items = [x for x in text.split("; ") if x]
    if item not in items:
        items.append(item)
    if order:
        items.sort(key=order.index)
    return "; ".join(items)


def observe(contacts, m, rid, when, examiner_included=()) -> list:
    """Record the people on one email. Returns the addresses that were new.
    examiner_included: outside addresses the examiner side put on this thread
    (doi_analysis.examiner_included) - marked Examiner_Thread = Y."""
    new = []
    people = [(m.sender, "Sender")] + [(a, "To") for a in m.to] + [(a, "CC") for a in m.cc]
    on_email = {a.lower() for a, _ in people if a}
    people += [(a, "Forwarded") for a in sorted(examiner_included) if a.lower() not in on_email]
    included = {a.lower() for a in examiner_included}
    for email, role in people:
        email = (email or "").strip().lower()
        if "@" not in email:                      # X500 leftovers etc. are not usable addresses
            continue
        row = contacts.get(email)
        if row is None:
            row = contacts[email] = {k: "" for k in FIELDS}
            row.update(Email=email, Domain=email.split("@")[-1], First_Seen=f"{when:%Y-%m-%d}")
            new.append(email)
        name = m.names.get(email) or (m.sender_name if role == "Sender" else "")
        if name and not row["Name"]:
            row["Name"] = name
        row["Roles"] = _merge(row["Roles"], role, ROLE_ORDER)
        if email in included:
            row["Examiner_Thread"] = "Y"
        row["Request_IDs"] = _merge(row["Request_IDs"], rid)
        row["Last_Seen"] = max(row["Last_Seen"], f"{when:%Y-%m-%d}")
    return new


def pending_senders(contacts) -> list:
    """Senders seen on request emails but not yet approved - your review list."""
    return sorted(e for e, r in contacts.items() if "Sender" in r["Roles"]
                  and r["Approved"].lower() not in ("y", "yes") + NO       # decided either way
                  and r["Examiner_Thread"].lower() != "y")                 # trusted via the examiner


def save(path, contacts) -> str:
    """Write the directory: unapproved senders first (they need you), then by domain. If the CSV is
    open in Excel nothing is lost - the next run sees the same emails and records them again."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    pending = set(pending_senders(contacts))
    rows = sorted(contacts.values(), key=lambda r: (r["Email"] not in pending, r["Domain"], r["Email"]))
    try:
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=FIELDS)
            w.writeheader()
            w.writerows(rows)
    except PermissionError:
        return f"contacts NOT saved: {path.name} is open - close it; the next run records them again"
    return f"contacts saved to {path}"
