"""Tests for DOI_Contacts.py - the contact directory. Invented people only."""
import csv
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import DOI_Contacts as C
from DOI_Filter import Mail

D1, D2 = datetime(2027, 10, 8), datetime(2027, 10, 15)


def mail(sender, to=(), cc=(), names=None, sender_name=""):
    return Mail("FIN-53 Request", "", sender, list(to), list(cc), received=D1,
                names=names or {}, sender_name=sender_name)


class ContactTests(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir)
        self.path = self.dir / "Contacts" / "DOI_Contacts.csv"

    def test_records_sender_to_and_cc_with_request_and_dates(self):
        c = {}
        new = C.observe(c, mail("Examiner@Example.test", ["contact@corp.test"], ["boss@corp.test"],
                                names={"contact@corp.test": "Jo Park"}, sender_name="Sam Lee"), "FIN-53", D1)
        self.assertEqual(new, ["examiner@example.test", "contact@corp.test", "boss@corp.test"])
        self.assertEqual(c["examiner@example.test"]["Roles"], "Sender")
        self.assertEqual(c["examiner@example.test"]["Name"], "Sam Lee")
        self.assertEqual(c["contact@corp.test"]["Name"], "Jo Park")
        self.assertEqual(c["boss@corp.test"]["Roles"], "CC")
        self.assertEqual(c["boss@corp.test"]["Domain"], "corp.test")

    def test_second_email_updates_instead_of_duplicating(self):
        c = {}
        C.observe(c, mail("examiner@example.test", ["contact@corp.test"]), "FIN-53", D1)
        new = C.observe(c, mail("contact@corp.test", ["examiner@example.test"]), "FIN-60", D2)
        self.assertEqual(new, [])
        row = c["contact@corp.test"]
        self.assertEqual((row["Roles"], row["Request_IDs"]), ("Sender; To", "FIN-53; FIN-60"))
        self.assertEqual((row["First_Seen"], row["Last_Seen"]), ("2027-10-08", "2027-10-15"))

    def test_only_approved_senders_are_trusted(self):
        c = {}
        C.observe(c, mail("a@example.test", ["b@corp.test"]), "FIN-53", D1)
        self.assertEqual(C.approved_senders(c), set())
        self.assertEqual(C.pending_senders(c), ["a@example.test"])        # recipients aren't on the review list
        c["a@example.test"]["Approved"] = "Y"
        self.assertEqual(C.approved_senders(c), {"a@example.test"})
        self.assertEqual(C.pending_senders(c), [])

    def test_save_load_round_trip_keeps_your_columns(self):
        c = {}
        C.observe(c, mail("a@example.test", ["b@corp.test"]), "FIN-53", D1)
        C.save(self.path, c)                                              # creates the folder too
        rows = list(csv.DictReader(open(self.path, encoding="utf-8-sig")))
        self.assertEqual(rows[0]["Email"], "a@example.test")              # pending senders listed first
        # you approve one and add a meeting contact by hand, then the script runs again
        rows[0]["Approved"], rows[0]["Notes"] = "y", "Lead examiner - confirmed at team meeting"
        rows.append(dict.fromkeys(C.FIELDS, "") | {"Email": "New.Person@Example.test", "Approved": "Y"})
        with open(self.path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=C.FIELDS); w.writeheader(); w.writerows(rows)
        c2 = C.load(self.path)
        C.observe(c2, mail("a@example.test"), "FIN-60", D2)
        C.save(self.path, c2)
        c3 = C.load(self.path)
        self.assertEqual(c3["a@example.test"]["Notes"], "Lead examiner - confirmed at team meeting")
        self.assertEqual(c3["a@example.test"]["Request_IDs"], "FIN-53; FIN-60")
        self.assertEqual(C.approved_senders(c3), {"a@example.test", "new.person@example.test"})

    def test_missing_file_is_empty_and_non_addresses_are_skipped(self):
        self.assertEqual(C.load(self.path), {})
        c = {}
        self.assertEqual(C.observe(c, mail("/O=EXCHANGE/OU=X500 LEFTOVER"), "FIN-53", D1), [])


if __name__ == "__main__":
    unittest.main()
