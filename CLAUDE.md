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
| `DOI_Exam_Tracker_1.0.xlsx` | The live tracker (target of all writes). **Real data — never upload it to Claude**; share structure only |
| `DOI_Exam_Test_Tracker_v2.xlsx` | Fake-data copy (`DOI-TEST-##`, `example.test`) — safe test fixture |
| `doi_filter.py` | Outlook fetch + qualification rules; settings block holds placeholders only |
| `local_settings.py` | Real paths/addresses, overrides `doi_filter` settings. **Git-ignored**; template in `local_settings.example.py` |
| `upsert_tracker.py` | Write/update tracker rows |
| `publish.py` | Publishes a view-only, Excel-recalculated copy to `PUBLISH_PATH` for coworkers |
| `import_original.py` | One-time: rebuild the original request list (sheet `Requests`) in the tracker template's format → new file + `_import_log.csv`. Run locally on real data |
| `run.py` | Entry point: fetch → qualify → upsert → publish → CSV log (`--days N`, `--dry-run`) |
| `tests/` | `python -m unittest -v`. Upsert tests need a fake-data tracker at `tests/fixtures/DOI_Exam_Test_Tracker_v2.xlsx` or `$DOI_TRACKER_FIXTURE`; skipped otherwise. `*.xlsx` is git-ignored — don't commit workbooks |

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
1. Subject contains a Request ID matching `REQUEST_ID_PATTERN` = `\b(?:FIN|IT|DOI-[A-Z]+)-\d+\b` — real IDs are `FIN-##` and `IT-##`; `DOI-TEST-##` is test data only. A new prefix must be added to **both** `REQUEST_ID_PATTERN` and `SUBJECT_LIKES` (the Outlook DASL pre-filter, OR'd by `subject_filter()`), or those emails never reach Python
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

- **Never write into formula columns** (A, E, F, H, J, N). When appending a row, copy formulas down from the first data row and extend `tblRequests.ref`.
- Request ID formula = first word of column D with `:` stripped (`FIN-52:` → `FIN-52`); shows `(no ID)` if that word has no digit. `upsert_tracker.first_word_id()` is its Python twin — keep them in step. Type = contains "Follow". Follows Up On = ID before `Follow-up`/`Follow Up`, handling `FIN-27` and `FIN 45` → `FIN-45`; blank if the follow-up is to a non-ID (e.g. an interview).
- Live data (since 2026-10-12): 18 real requests, two assignees on Lists. Due dates given as text ("ASAP") go to Notes, never into column G (keeps date math numeric).
- **Known workbook bug:** `StatusList` = `OFFSET(Lists!$A$4,…,COUNTA(Lists!$A$4:$A$60))` also counts the help text that starts at Lists!A14, so the Status dropdown shows blanks and help paragraphs. Fix: count `$A$4:$A$12` instead. Code reads Lists columns only down to the first blank (`import_original.list_block`).
- **Sheet Summary**: settings `C5` AsOf (`=TODAY()`), `C6` FollowUpDays, `C7` DueSoonDays; KPIs and breakdowns via named ranges (`trkStatus`, `trkDue`, `trkRequest`, … all rows 4:1000).
- **Sheet Lists**: dropdown sources via dynamic names `StatusList` (A4↓) and `AssigneeList` (D4↓).

## Upsert logic (`upsert_tracker.py`)

- Key = Request ID from subject regex, matched against the first word of column D (same rule as column A's formula). New ID → append row under `tblRequests`: formulas and styles copied from the **first** data row (not the row above — a stray bottom row must not spread), `tbl.ref` + autoFilter extended. D gets RE:/FW: stripped and the ID forced to be the first word.
- Existing ID → C keeps the **earliest** received date; L filled only if blank (never overwrites typed contacts); O gets one appended line per email ending `[ref xxxxxxxx]` (first 8 hex of SHA-1 of EntryID).
- Idempotent via that ref tag: an email whose ref is already in O is skipped. Nothing changed → no backup, no save.
- Timestamped backup to `backups/` next to the tracker before every save.
- openpyxl requires Excel closed: refuses if `~$<name>` exists or the file can't be opened r+b. Also refuses if the file has charts/images/slicers (openpyxl would delete them).
- Not yet done: M (Last Email Sent, needs Sent Items lookup by `ConversationID`) and G (Due Date parsing). Neither is written by the script yet.
- `extract()` strips pywin32's timezone tag from `ReceivedTime` — openpyxl rejects tz-aware datetimes.

## Importing the original request list (`import_original.py`)

- The original list holds real names/addresses: **never upload it to Claude** — the user runs the importer locally. Tests build a fake original mimicking its quirks (colon after ID, en dash, `ABC 27 Follow up`, missing ID, free-text Status, `TBD` due date).
- Maps by header name: Assigned to→B, Initial Request Email Received→C, Request Name→D, Due Date→G, Status→I, Completion Date→K, Contact→L, Last Email Sent→M. Template's example rows are replaced; new assignees appended to Lists!D.
- D is tidied so the formulas parse it; free-text Status and non-date Due Date go to Notes with Status left blank. Every change/undecidable item is printed and logged per row — nothing is silently guessed. Never overwrites an existing output.

## Sharing with coworkers (`publish.py`)

- The user's tracker (`TRACKER_PATH`) is the single source of truth; only the script and the user edit it. Coworkers only **view** a copy at `PUBLISH_PATH` (unset = no publishing). Their edits to the copy are overwritten — don't build two-way sync.
- Publishes when the master is newer than the copy (so the user's own manual edits go out too), even if the upsert was refused because the master was open.
- `excel_recalc()` is the only Excel-COM function (Windows-only, tested by the user): a separate hidden Excel (`DispatchEx`) opens a temp copy, `CalculateFull()`, saves with `ReadOnlyRecommended`. Needed because openpyxl saves no cached results, so Protected View / browser previews would show blank formula columns.
- Swap into place with `os.replace` from a staged file in the same folder (atomic). Copy locked by a coworker or Excel failure → nothing published, logged, retried next run.

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
