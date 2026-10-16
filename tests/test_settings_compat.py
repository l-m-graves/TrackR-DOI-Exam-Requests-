"""Settings written the way the real local_settings.py writes them: a request-ID pattern that
allows 'FIN 45', the company's own domain in SENDER_DOMAINS, whole addresses where domains are
expected, MUST_BE_SENT_TO for external emails, and LOG_PATH given as a folder. Fake values only."""
import re
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import doi_filter
from attachment_text import read_text
from doi_analysis import detect
from doi_filter import Mail, canonical_id, log_file, normalize_domains
from tests import fake_doi_emails as E

LOOSE_PATTERN = r"\b(?:FIN|IT)[ -]?\d+\b|\bDOI-[A-Z0-9]+-\d+\b"     # the shape used in real settings


class LooseIdTests(unittest.TestCase):
    def setUp(self):
        p = mock.patch.object(doi_filter, "REQUEST_ID", re.compile(LOOSE_PATTERN))
        p.start()
        self.addCleanup(p.stop)

    def test_every_spelling_becomes_one_id(self):
        for text in ("FIN-45 Request", "RE: FIN 45 Request", "FIN45 support", "Request FIN-45:"):
            self.assertEqual(doi_filter.find_id(text), "FIN-45", text)
        self.assertEqual(doi_filter.find_id("IT 39 Access Reviews"), "IT-39")
        self.assertEqual(doi_filter.find_id("DOI-TEST-07 Request"), "DOI-TEST-07")

    def test_ordinary_words_are_not_ids(self):
        self.assertEqual(doi_filter.find_id("Please send it 5 days early"), "")     # case-sensitive
        self.assertEqual(doi_filter.find_id("Final 2027 numbers"), "")

    def test_tracker_text_gets_the_hyphenated_id_first(self):
        import upsert_tracker
        with mock.patch.object(upsert_tracker, "REQUEST_ID", re.compile(LOOSE_PATTERN)):
            self.assertEqual(upsert_tracker.request_text("RE: FIN 45 Request - Payroll", "FIN-45"),
                             "FIN-45 Request - Payroll")
            self.assertEqual(upsert_tracker.first_word_id("FIN-45 Request - Payroll"), "FIN-45")

    def test_canonical_id_leaves_hyphenated_ids_alone(self):
        self.assertEqual(canonical_id("FIN-45"), "FIN-45")
        self.assertEqual(canonical_id("fin 45"), "FIN-45")
        self.assertEqual(canonical_id("DOI-TEST-01"), "DOI-TEST-01")


class DomainSettingTests(unittest.TestCase):
    def test_whole_addresses_count_as_their_domain(self):
        self.assertEqual(normalize_domains({"Someone@Exam-Firm.test", "@doi-state.test", "corp.test", " "}),
                         {"exam-firm.test", "doi-state.test", "corp.test"})

    def test_own_company_in_sender_domains_still_counts_as_internal(self):
        with mock.patch.multiple(doi_filter, **dict(E.ROLES, SENDER_DOMAINS={"corp.test", "doi-state.test"})):
            self.assertEqual(doi_filter.role_of("colleague@corp.test"), "internal")
            self.assertEqual(doi_filter.role_of("ana.accountant@corp.test"), "accountant")
            self.assertEqual(doi_filter.role_of("someone@doi-state.test"), "examiner")
        with mock.patch.multiple(doi_filter, **dict(E.ROLES, SENDER_DOMAINS={"legacy-doi.test"}, EXAMINER_DOMAINS=set())):
            self.assertEqual(doi_filter.role_of("x@legacy-doi.test"), "examiner")     # legacy list still works


class MustBeSentToTests(unittest.TestCase):
    MAILBOX = "audit@corp.test"

    def setUp(self):
        p = mock.patch.multiple(doi_filter, **dict(E.ROLES, MUST_BE_SENT_TO={self.MAILBOX}))
        p.start()
        self.addCleanup(p.stop)

    def judge(self, sender, to, cc=(E.INTERN,)):
        return detect(E.mail("FIN-401 Request - Premium Tax",
                             "Please provide the DOI examination support for premium tax.", sender, to=to, cc=cc))

    def test_external_email_must_reach_the_mailbox(self):
        self.assertEqual(self.judge("lead.examiner@doi-state.test", [self.MAILBOX]).status, "accept")
        det = self.judge("lead.examiner@doi-state.test", ["ana.accountant@corp.test"])
        self.assertEqual(det.status, "reject")
        self.assertIn("MUST_BE_SENT_TO", det.reasons[0])

    def test_internal_forward_to_you_is_not_held_to_it(self):
        det = self.judge("ana.accountant@corp.test", [E.INTERN], cc=())
        self.assertEqual(det.status, "accept")


class LogFolderTests(unittest.TestCase):
    def test_folder_gets_the_standard_file_name(self):
        self.assertEqual(Path(log_file("/share/DOI Requests Tracker/Log")),
                         Path("/share/DOI Requests Tracker/Log/DOI_Run_Log.csv"))
        self.assertEqual(log_file("/share/Logs/My_Log.csv"), "/share/Logs/My_Log.csv")
        self.assertIsNone(log_file(None))


class CsvAttachmentTests(unittest.TestCase):
    def test_csv_attachments_are_read(self):
        d = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, d)
        path = d / "FIN-45_Request_Items.csv"
        path.write_text("Item,Description\n1,Premium tax returns for 2026\n", encoding="utf-8")
        text, status = read_text(path)
        self.assertEqual(status, "read")
        self.assertIn("Premium tax returns", text)


if __name__ == "__main__":
    unittest.main()
