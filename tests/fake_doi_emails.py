"""Invented DOI-exam emails covering the shapes real correspondence takes. No real data:
people and domains are fictional (.test), request numbers are made up."""
from datetime import datetime

from doi_filter import Mail

# Role settings used by these tests - the same names you fill in local_settings.py.
ROLES = dict(
    SENDER_DOMAINS=set(),
    LEAD_EXAMINERS={"lead.examiner@doi-state.test"},
    EXAMINER_CONTACTS=set(),
    EXAMINER_DOMAINS={"doi-state.test"},
    THIRD_PARTY_CONTACTS=set(),
    THIRD_PARTY_DOMAINS={"examfirm.test"},
    ACCOUNTANT_CONTACTS={"ana.accountant@corp.test"},
    INTERNAL_DOMAINS={"corp.test"},
    ENTITY_NAMES=("Sample Casualty Company",),
    PRODUCT_NAMES=("Personal Auto",),
    MUST_BE_SENT_TO=set(),
)
DAY = datetime(2027, 10, 4, 9, 30)
INTERN = "intern@corp.test"          # the person running the script - usually only CC'd


def mail(subject, body, sender, to=("ana.accountant@corp.test",), cc=(INTERN,), attachments=(), eid="E1", **kw):
    return Mail(subject, body, sender, list(to), list(cc), list(attachments), kw.pop("received", DAY), eid, **kw)


DIRECT = mail("FIN-201 Request - Premium Tax",
              "Good morning,\nPlease provide the 2026 premium tax returns and reconciliations by October 22, 2027.\n"
              "Thank you,\nLead Examiner", "lead.examiner@doi-state.test", attachments=["FIN-201 Request.pdf"])

FORWARDED = mail("FW: FIN-202 Request - Claims Handling",
                 "Team - please see the examiner's request below. Payroll and Claims, please send me your pieces.\n\n"
                 "________________________________\n"
                 "From: Examiner Two <examiner.two@doi-state.test>\n"
                 "Sent: Monday, October 4, 2027 9:30 AM\n"
                 "To: Ana Accountant <ana.accountant@corp.test>\n"
                 "Subject: FIN-202 Request - Claims Handling\n\n"
                 "Please provide the closed claims listing by 10/29/2027.\n",
                 "ana.accountant@corp.test", to=["claims.lead@corp.test"], eid="E2")

INTERNAL_PAYROLL = mail("FIN-203 Payroll records for the DOI exam",
                        "Hi Pat,\nFor DOI examination request FIN-203, please send the Payroll department's 2026 "
                        "payroll register.\nWe expect to have it done by 10/25/2027.\nThanks, Ana",
                        "ana.accountant@corp.test", to=["pat.payroll@corp.test"], eid="E3")

CONFLICTING = mail("FIN-207 Request - Investments",
                   "Please provide Part A of the investment schedules by 10/20/2027 and Part B by 10/29/2027.",
                   "examiner.three@doi-state.test", eid="E7")

ATTACHMENT_ONLY = mail("Examination request letter",
                       "Please see the attached letter from the Department of Insurance.",
                       "lead.examiner@doi-state.test", attachments=["DOI Request Letter.docx"], eid="E8")
ATTACHMENT_TEXT = ("Department of Insurance - Examination Request FIN-204\n"
                   "Please provide the following by November 5, 2027:\n"
                   "1. Board meeting minutes for 2026\n2. The reinsurance contracts in force\n")

MULTI_ITEM = mail("FIN-208 Request - Underwriting Pricing",
                  "Please provide the following by 11/01/2027:\n"
                  "1. Rate filings for Personal Auto in Texas and California\n"
                  "2. The pricing model documentation\n"
                  "3. Explain how control PT-07 and ITGC-04 are tested\n"
                  "Scope: Sample Casualty Company, for the period January 1, 2026 to December 31, 2026.",
                  "lead.examiner@doi-state.test", eid="E9")

FOLLOW_UP_SUBJECT = mail("FIN-205 Request - Reserve Support - FIN-201 Follow-up",
                         "Please provide the support missing from your FIN-201 response by 11/10/2027.",
                         "lead.examiner@doi-state.test", eid="E10")
FOLLOW_UP_BODY = mail("FIN-206 Request - Additional Support",
                      "This is a follow-up to FIN-201. Please provide the reconciliation by 11/12/2027.",
                      "lead.examiner@doi-state.test", eid="E11")

IT_DEPARTMENT = mail("IT-39 Request - Access Reviews",
                     "Please obtain the quarterly user access reviews from the IT department by 10/30/2027.",
                     "ana.accountant@corp.test", to=["it.lead@corp.test"], eid="E12")

EVENTS = mail("RE: FIN-201 Request - Premium Tax",
              "We are missing the support for item 2. Please clarify which entity it covers. We can grant an "
              "extension to 11/12/2027. Please schedule an interview with the Controller. Upload the files to "
              "the SharePoint site.", "lead.examiner@doi-state.test", eid="E13")
FILES_READY = mail("RE: FIN-201 Request - Premium Tax", "The files are ready for review before we send them.",
                   "ana.accountant@corp.test", eid="E14")
SUBMITTED = mail("RE: FIN-201 Request - Premium Tax", "All materials have been uploaded to the examiners' SharePoint site.",
                 "ana.accountant@corp.test", eid="E15")

FALSE_POSITIVE_NEWS = mail("DOI announces new filing guidance",
                           "The Department of Insurance published new guidance for rate filings.",
                           "newsletter@industry-news.test", to=[INTERN], cc=[], eid="E16")
FALSE_POSITIVE_ID = mail("IT-2024 budget planning", "Please send your budget numbers by Friday.",
                         "ana.accountant@corp.test", eid="E17")
UNKNOWN_SENDER = mail("FIN-209 Request - Agency Review", "Please provide the examination items by 11/03/2027.",
                      "someone@unlisted-firm.test", eid="E18")
