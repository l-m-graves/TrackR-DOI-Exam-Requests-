"""Tests for due_dates.py. Email wording here is invented."""
import unittest
from datetime import date, datetime

from due_dates import find_due, strip_quoted

RECEIVED = datetime(2027, 10, 4, 9, 30)


def due(body):
    return find_due(body, RECEIVED)


class FindDueTests(unittest.TestCase):
    def test_common_wordings(self):
        cases = {
            "Please provide the items listed below by October 15, 2027.": date(2027, 10, 15),
            "Responses are due 10/15/2027.": date(2027, 10, 15),
            "Due Date: 10/15/27": date(2027, 10, 15),
            "Kindly respond no later than Oct. 15th.": date(2027, 10, 15),        # year inferred
            "Deadline - 2027-10-15": date(2027, 10, 15),
            "Please submit by Friday, October 15, 2027 COB.": date(2027, 10, 15),
            "Return by 10-15-2027 please.": date(2027, 10, 15),
        }
        for body, expected in cases.items():
            with self.subTest(body=body):
                self.assertEqual(due(body), {"date": expected, "note": None})

    def test_same_date_twice_is_still_one_date(self):
        self.assertEqual(due("Due 10/15/2027. Again: responses due October 15, 2027.")["date"], date(2027, 10, 15))

    def test_two_different_dates_are_not_guessed(self):
        r = due("Part A is due 10/15/2027 and Part B is due 10/29/2027.")
        self.assertIsNone(r["date"])
        self.assertIn("2027-10-15, 2027-10-29", r["note"])

    def test_text_due_dates_become_a_note(self):
        self.assertEqual(due("Please respond ASAP."), {"date": None, "note": "due: ASAP"})
        self.assertEqual(due("Due date: TBD")["note"], "due: TBD")

    def test_dates_without_due_wording_are_ignored(self):
        self.assertEqual(due("Following our meeting on 10/01/2027, attached is the request."),
                         {"date": None, "note": None})

    def test_quoted_history_is_ignored(self):
        body = ("Thanks - the revised response is due 10/20/2027.\n\n"
                "________________________________\nFrom: Examiner\nSent: Monday\n"
                "The original response was due 10/15/2027.")
        self.assertEqual(due(body)["date"], date(2027, 10, 20))
        self.assertNotIn("10/15/2027", strip_quoted(body))

    def test_implausible_dates_are_ignored(self):
        self.assertIsNone(due("This replaces the request that was due 01/05/2026.")["date"])   # long past
        self.assertIsNone(due("Due 13/45/2027")["date"])                                       # not a date

    def test_month_without_year_rolls_into_next_year(self):
        r = find_due("Please respond by January 10.", datetime(2027, 12, 20))
        self.assertEqual(r["date"], date(2028, 1, 10))

    def test_empty_body(self):
        self.assertEqual(find_due("", RECEIVED), {"date": None, "note": None})
        self.assertEqual(find_due(None, RECEIVED), {"date": None, "note": None})


if __name__ == "__main__":
    unittest.main()
