"""Tests for the rules in DOI_Filter.py - no Outlook needed.

Run from the repo folder:  python -m unittest -v
"""
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import DOI_Filter
from DOI_Filter import (Mail, attachment_targets, extract, important_attachments, qualify,
                        save_attachments, subject_filter)

# Pin the settings so a real Local_Settings.py on your laptop can't change test results.
TEST_SETTINGS = {
    "SENDER_DOMAINS": {"example.test"},
    "MUST_BE_SENT_TO": {"audit@example.test"},
}


def good_mail(**overrides):
    m = Mail(
        subject="DOI-TEST-60 Request - Claims Handling",
        body="Please provide the documents listed in the attached request.",
        sender="examiner@example.test",
        to=["audit@example.test"],
        attachments=["image001.png", "DOI-TEST-60 Request Letter.pdf"],
        received=datetime(2027, 10, 4, 9, 30),
        entry_id="00000000FAKEENTRYID60",
        sender_name="Jordan Lee",
    )
    for k, v in overrides.items():
        setattr(m, k, v)
    return m


class QualifyTests(unittest.TestCase):
    def setUp(self):
        patcher = mock.patch.multiple(DOI_Filter, **TEST_SETTINGS)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_good_mail_passes_and_returns_request_id(self):
        self.assertEqual(qualify(good_mail()), (True, "DOI-TEST-60"))

    def test_reply_subject_still_finds_id(self):
        self.assertEqual(qualify(good_mail(subject="RE: DOI-TEST-60 Request"))[1], "DOI-TEST-60")

    def test_each_rule_rejects_with_its_own_reason(self):
        cases = {
            "subject has no request ID": dict(subject="Quarterly update"),
            "sender not approved": dict(sender="someone@other.test"),
            "not addressed to MUST_BE_SENT_TO": dict(to=["someone@example.test"]),
            "body lacks request wording": dict(body="FYI only."),
            "no attachment named": dict(attachments=["image001.png", "notes.pdf"]),
        }
        for reason, change in cases.items():
            with self.subTest(reason=reason):
                ok, why = qualify(good_mail(**change))
                self.assertFalse(ok)
                self.assertIn(reason, why)

    def test_individually_approved_sender_passes(self):
        m = good_mail(sender="coordinator@other.test")
        self.assertFalse(qualify(m)[0])
        self.assertEqual(qualify(m, approved_senders={"coordinator@other.test"}), (True, "DOI-TEST-60"))

    def test_recipient_rule_skipped_when_setting_is_empty(self):
        with mock.patch.object(DOI_Filter, "MUST_BE_SENT_TO", set()):
            self.assertTrue(qualify(good_mail(to=["assigned.contact@example.test"], cc=[]))[0])

    def test_cc_counts_as_addressed(self):
        self.assertTrue(qualify(good_mail(to=[], cc=["audit@example.test"]))[0])

    def test_real_id_formats(self):
        for subject, rid in [("FIN-53 Request - Fake Subject", "FIN-53"),
                             ("RE: IT-39 Request - Fake Subject", "IT-39"),
                             ("FIN-52: Fake Subject - FIN 45 Follow Up", "FIN-52")]:
            with self.subTest(subject=subject):
                self.assertEqual(qualify(good_mail(subject=subject)), (True, rid))

    def test_lookalikes_are_not_ids(self):
        for subject in ("AUDIT-2024 planning", "PROFIT-12 memo", "FIN 53 no hyphen", "fin-53 lower case"):
            with self.subTest(subject=subject):
                self.assertEqual(qualify(good_mail(subject=subject)), (False, "subject has no request ID"))

    def test_outlook_prefilter_covers_every_prefix(self):
        f = subject_filter()
        for prefix in ("%FIN-%", "%IT-%", "%DOI-%"):
            self.assertIn(f"LIKE '{prefix}'", f)
        self.assertTrue(f.startswith("@SQL=("))
        self.assertTrue(f.endswith(') AND "urn:schemas:httpmail:hasattachment" = 1'))

    def test_lowercase_id_is_not_a_request_id(self):
        self.assertFalse(qualify(good_mail(subject="doi-test-60 request"))[0])


class AttachmentTests(unittest.TestCase):
    def test_name_matching_ignores_case_spaces_hyphens_underscores(self):
        names = ["follow up letter.pdf", "FOLLOWUP.docx", "doi_items.xlsx", "Data_Request.pdf"]
        self.assertEqual(important_attachments(names), names)

    def test_images_ignored_even_if_named_right(self):
        self.assertEqual(important_attachments(["DOI logo.PNG", "request.jpeg"]), [])

    def test_unrelated_names_ignored(self):
        self.assertEqual(important_attachments(["minutes.pdf", "budget.xlsx"]), [])


class AttachmentSaveTests(unittest.TestCase):
    NAMES = ["image001.png", "FIN-53 Request Letter.pdf", "notes.docx", "Follow-up: items?.xlsx",
             "Portal Screenshot.PNG", "data.csv", "old.doc"]
    DAY = datetime(2027, 10, 8, 9, 30)

    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.dir)

    def fake_msg(self, fail_on=()):
        written = []
        def item(i):
            def save(path):
                if i in fail_on:
                    raise OSError("blocked by policy")
                Path(path).write_text(f"attachment {i}")
                written.append(i)
            return SimpleNamespace(SaveAsFile=save)
        return SimpleNamespace(Attachments=SimpleNamespace(Item=item)), written

    def test_plan_keeps_listed_types_with_outlook_numbering(self):
        plan = attachment_targets("FIN-53", self.NAMES, self.DAY, self.dir)
        self.assertEqual([i for i, _ in plan], [2, 3, 4, 5, 7])         # 1-based; logo + .csv skipped
        self.assertEqual([d.relative_to(self.dir).as_posix() for _, d in plan],
                         ["FIN-53/2027-10-08_FIN-53 Request Letter.pdf",
                          "FIN-53/2027-10-08_notes.docx",
                          "FIN-53/2027-10-08_Follow-up_ items_.xlsx",    # ':' and '?' made safe
                          "FIN-53/2027-10-08_Portal Screenshot.PNG",     # a real screenshot is kept
                          "FIN-53/2027-10-08_old.doc"])

    def test_save_then_rerun_saves_nothing_new(self):
        m = Mail("FIN-53 Request", "", "x@example.test", attachments=self.NAMES, received=self.DAY)
        msg, written = self.fake_msg()
        first = save_attachments(msg, m, "FIN-53", self.dir)
        self.assertEqual(written, [2, 3, 4, 5, 7])
        self.assertTrue(all(r.startswith("saved ") for r in first))
        self.assertEqual(save_attachments(msg, m, "FIN-53", self.dir), [])   # already there
        self.assertEqual(written, [2, 3, 4, 5, 7])

    def test_same_name_on_a_later_email_is_kept_separately(self):
        m1 = Mail("FIN-53", "", "", attachments=["Request.pdf"], received=self.DAY)
        m2 = Mail("RE: FIN-53", "", "", attachments=["Request.pdf"], received=self.DAY + timedelta(days=3))
        save_attachments(self.fake_msg()[0], m1, "FIN-53", self.dir)
        save_attachments(self.fake_msg()[0], m2, "FIN-53", self.dir)
        self.assertEqual(sorted(p.name for p in (self.dir / "FIN-53").iterdir()),
                         ["2027-10-08_Request.pdf", "2027-10-11_Request.pdf"])

    def test_one_failure_does_not_stop_the_rest(self):
        m = Mail("FIN-53", "", "", attachments=self.NAMES, received=self.DAY)
        msg, written = self.fake_msg(fail_on={2})
        results = save_attachments(msg, m, "FIN-53", self.dir)
        self.assertTrue(results[0].startswith("FAILED") and "blocked by policy" in results[0])
        self.assertTrue(all(r.startswith("saved ") for r in results[1:]))
        self.assertEqual(written, [3, 4, 5, 7])


class ExtractTests(unittest.TestCase):
    """extract() with fake objects shaped like Outlook's COM items."""

    def fake_msg(self):
        def person(addr, kind=None):
            p = SimpleNamespace(Address=addr, Type=kind,
                                PropertyAccessor=SimpleNamespace(GetProperty=lambda prop: addr.upper()))
            return p
        return SimpleNamespace(
            Subject="DOI-TEST-61 Request", Body="please provide",
            Sender=person("examiner@example.test"), SenderEmailAddress="x", SenderName="Jordan Lee",
            Recipients=[person("audit@example.test", 1), person("cc@example.test", 2),
                        person("hidden@example.test", 3)],
            Attachments=[SimpleNamespace(FileName="DOI-TEST-61.pdf")],
            # pywin32 returns a timezone-tagged value like this one
            ReceivedTime=datetime(2027, 10, 5, 14, 5, tzinfo=timezone(timedelta(0))),
            EntryID="FAKEID61",
        )

    def test_fields(self):
        m = extract(self.fake_msg())
        self.assertEqual(m.sender, "examiner@example.test")      # lower-cased SMTP
        self.assertEqual(m.to, ["audit@example.test"])
        self.assertEqual(m.cc, ["cc@example.test"])               # BCC left out
        self.assertEqual(m.sender_name, "Jordan Lee")
        self.assertEqual(m.attachments, ["DOI-TEST-61.pdf"])
        self.assertEqual(set(m.names), {"audit@example.test", "cc@example.test"})   # BCC not recorded

    def test_received_time_is_naive_for_openpyxl(self):
        m = extract(self.fake_msg())
        self.assertIsNone(m.received.tzinfo)
        self.assertEqual(m.received, datetime(2027, 10, 5, 14, 5))

    def test_smtp_falls_back_to_address(self):
        msg = self.fake_msg()
        msg.Sender = SimpleNamespace(Address="Fallback@Example.test")    # no PropertyAccessor
        self.assertEqual(extract(msg).sender, "fallback@example.test")


if __name__ == "__main__":
    unittest.main()
