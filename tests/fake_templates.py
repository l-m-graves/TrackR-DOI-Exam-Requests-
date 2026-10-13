"""The team's five DOI email templates, filled in with invented values (no real data)."""
D = "–"   # the en dash Outlook puts in the subjects

ASSIGNMENT = (f"DOI Exam Request FIN-990 {D} Premium Tax Support",
              "Hi Jo,\n\nThe California DOI has issued examination request FIN-990 regarding premium tax "
              "support. Please review the attached request and provide the requested evidence and "
              "responses by October 22, 2027.\n\nI have assigned the request items to you in the tracker. "
              "Please let me know if another contact would be better suited for any of the items.\n\n"
              "If you have any questions, please reach out to me directly.\n\nThank you,\nSam")

ACKNOWLEDGMENT = (f"RE: DOI Exam Request FIN-990 {D} Premium Tax Support",
                  "Hi Sam,\n\nWe received the DOI exam request and are reviewing the assigned items. We will "
                  "work toward providing the requested evidence and responses by October 22, 2027.\n\nI will "
                  "let you know if we identify any questions or need to coordinate with another contact.\n\n"
                  "Thank you,\nJo\n\n________________________________\nFrom: Sam\nSent: Monday\n"
                  "The California DOI has issued examination request FIN-990 regarding premium tax support.")

STATUS_UPDATE = (f"DOI Exam Request FIN-990 {D} Status Update",
                 "Hi Sam,\nHere is a status update for request FIN-990:\n\nWe are still gathering the 2026 "
                 "filings for item 3. We expect to provide the remaining materials by 10/29/2027.\nPlease let "
                 "me know if you need anything else in the meantime.\nThank you, Jo")

RESPONSE_MATERIALS = (f"DOI Exam Request FIN-990 {D} Response Materials",
                      "Hi Sam,\nAttached are the completed responses and supporting evidence for request "
                      "FIN-990, premium tax support.\nThe materials include:\nItem 1 response\nItem 2 schedule\n"
                      "Item 3 filings\nPlease let me know if any additional information is needed or if any "
                      "item should be revised before submission.\nThank you, Jo")

SCOPE_QUESTION = ("Question Regarding DOI Exam Request FIN-990",
                  "Hi Sam,\n\nWe are reviewing request FIN-990 and would like to confirm the scope of item 4.\n\n"
                  "Could you please confirm whether the request applies to the period January 1, 2026 to "
                  "December 31, 2026? If the reinsurance condition does not apply to the company, please "
                  "confirm whether the related schedule can be omitted.\n\nWe will proceed with the "
                  "remaining items while we wait for clarification.\n\nThank you,\nJo")
