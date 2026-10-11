# CLAUDE.md — DOI Exam Tracker Automation

Context for Claude Code working in this repo. Read before changing anything.

## What this project is

A Python tool that reads Department of Insurance (DOI) exam-request emails from **classic Outlook** on a Mercury Insurance work laptop, validates them, and writes/updates rows in `DOI_Exam_Test_Tracker_v2.xlsx`. Built by an accounting intern learning Python — explain the *why* behind changes, not just the fix.

## Hard constraints (do not work around these)

- **Runs locally only.** The work mailbox allows no AI or outside access. The script talks to the signed-in Outlook app via COM (`pywin32`) on the user's own machine. No Graph API, no IMAP, no cloud calls, no credentials in code.
- **Never ask for or paste real DOI emails, attachments, names, or addresses.** Use placeholders and fake test data (`example.test` domain, `DOI-TEST-##` IDs).
- **Classic Outlook required.** "New Outlook" (web-based) has no COM object model. If COM fails, check for the New Outlook toggle first.
- **Claude cannot run the Outlook half.** Anything touching `win32com` is tested by the user on Windows. Keep COM code isolated so everything else is testable here.

## Files

| File | Role |
|---|---|
| `DOI_Exam_Test_Tracker_v2.xlsx` | The tracker (target of all writes) |
| `doi_filter.py` | Outlook fetch + qualification rules (exists) |
| `upsert_tracker.py` | Write/update tracker rows (planned) |
| `run.py` | Entry point: fetch → qualify → upsert → log (planned) |

## Architecture

```
Outlook (COM) ──Restrict()──> extract() ──> Mail dataclass ──> qualify() ──> upsert_tracker()
   coarse filter in Outlook     only COM-touching fn    plain Python     returns (ok, reason)
```

- `extract(msg) -> Mail` is the **only** function that touches COM. Everything downstream works on the `Mail` dataclass so it can be unit-tested with fake objects.
- `qualify(m) -> (bool, str)`: every rejection returns a reason string. Log rejections — that is the audit trail.
- `Items.Restrict()` dialects can't be mixed: chain one Jet filter (`[ReceivedTime] >= '...'`) and one DASL filter (`@SQL=...`).

## Qualification rules (in `doi_filter.py` settings block)

An email qualifies only if **all** pass:
1. Subject contains a Request ID matching `\bDOI-[A-Z]+-\d+\b`
2. Sender domain in `SENDER_DOMAINS` (placeholder — user fills in the real DOI domain)
3. `MUST_BE_SENT_TO` appears on To or CC
4. Body contains one of `BODY_PHRASES`
5. At least one attachment whose **file name** contains `DOI`, `Request`, or `Follow-up` — matched case/space/hyphen/underscore-insensitive via `_norm()`; image extensions (`.png .gif .jpg .jpeg`) ignored. Matching is on file name, not document contents.

## Outlook COM gotchas

- Collections are 1-based. `msg.Class == 43` = MailItem; skip everything else.
- `GetDefaultFolder(6)` = Inbox; subfolders by name: `inbox.Folders("DOI Exam")`.
- Internal Exchange addresses come back as X500 strings. Use `PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)` with fallback to `.Address` (see `smtp_address()`).
- `Recipients[i].Type`: 1 = To, 2 = CC, 3 = BCC.
- `msg.EntryID` is the stable item ID — use as secondary dedupe key.
- Attachments: `.FileName`, `.SaveAsFile(path)`. Save only `important_attachments()` to `ATTACH_DIR` as `{RequestID}_{FileName}`.

## Tracker workbook structure

Sheet **DOI Exam Tracker**, Excel Table `tblRequests`, header row 3, data from row 4. Navy headers = input, slate headers = formula.

| Col | Header | Type | Script writes? |
|---|---|---|---|
| A | Request ID | formula (parsed from D) | **no** |
| B | Assigned To | input, dropdown | optional |
| C | Initial Request Received | input, date | yes ← `ReceivedTime` |
| D | Initial Request | input, text | yes ← subject |
| E | Type | formula | **no** |
| F | Follows Up On | formula | **no** |
| G | Due Date | input, date | only if reliably parsed from body |
| H | Days Until Due | formula | **no** |
| I | Status | input, dropdown | yes, `"Not Started"` for new rows only |
| J | Action Flag | formula | **no** |
| K | Completed On | input, date | no |
| L | Contact | input, text | yes ← `Name <smtp>` |
| M | Last Email Sent | input, date | yes ← newest Sent Items match by `ConversationID` |
| N | Days Since Last Email | formula | **no** |
| O | Notes / Latest Update | input, text | append only, never overwrite |

- **Never write into formula columns** (A, E, F, H, J, N). When appending a row, copy formulas down from the row above and extend `tblRequests.ref`.
- Request ID formula = first word of column D, so the subject text must start with the ID.
- **Sheet Summary**: settings `C5` AsOf (`=TODAY()`), `C6` FollowUpDays, `C7` DueSoonDays; KPIs and breakdowns via named ranges (`trkStatus`, `trkDue`, `trkRequest`, … all rows 4:1000).
- **Sheet Lists**: dropdown sources via dynamic names `StatusList` (A4↓) and `AssigneeList` (D4↓).

## Upsert logic (planned)

- Key = Request ID from subject regex. Existing ID → update C/L/M, append to O. New ID → append row.
- Must be idempotent: running twice produces no duplicates.
- Backup the workbook (timestamped copy) before every write.
- openpyxl requires Excel closed; xlwings works with it open. Default: openpyxl, refuse to run if the file is locked.

## Workbook conventions (if regenerating/editing)

- Arial throughout (workbook default font set to Arial 10). Formulas, never hardcoded results.
- Excel-2007-safe functions; post-2007 ones need `_xlfn.` prefix (`_xlfn.MINIFS`). No XLOOKUP/FILTER/UNIQUE/SORT.
- openpyxl drops cached values; LibreOffice recalc drops `tableStyleInfo`/`calculatedColumnFormula` — restore table XML after recalc.

## Environment

- Windows 11, VS Code, **Git Bash** as default terminal — give bash syntax, label any PowerShell/cmd commands explicitly.
- Python + `pip install pywin32 openpyxl`. Tracker path set as raw string `r"C:\Users\<you>\OneDrive - ...\DOI_Exam_Test_Tracker_v2.xlsx"` (OneDrive-synced path, not a SharePoint URL).
- VS Code autocomplete/Copilot deliberately disabled — don't suggest enabling.

## Fallbacks if COM is blocked

1. Power Automate: "When a new email arrives" → "Add a row into a table" targeting `tblRequests`.
2. Outlook VBA (same object model).
3. Export emails as `.msg` and parse with `extract-msg`; `qualify()` works unchanged.

## Working style

- Be direct; state assumptions and proceed rather than stacking questions.
- Run and test everything runnable before handing it over (fake `Mail` objects for `qualify()`; fake rows for upsert).
- Flag when something reflects real audit/D&A practice (audit trail, idempotency, evidence linking) vs. a learning shortcut.
