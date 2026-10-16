# CLAUDE.md — DOI Exam Tracker Automation

Context for Claude Code working in this repo. Read before changing anything.

## What this project is

A Python tool that reads Department of Insurance (DOI) exam-request emails from **classic Outlook** on a Mercury Insurance work laptop, validates them, and writes/updates rows in `DOI_Exam_Test_Tracker_v2.xlsx`. Built by an accounting intern learning Python — explain the *why* behind changes, not just the fix.

**Who runs it:** a Finance Intern supporting the accountants. They are usually **CC'd or forwarded**, rarely the sender or responder, and may never see the examiner's original email. Never assume a message to them means they must do the work, never write emails in their voice. The tool is **read-only towards Outlook**: it never sends, replies, forwards, moves or deletes.

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
| `local_settings.py` | Real paths/addresses, overrides `doi_filter` settings. Found by `doi_filter.find_settings()`: next to the program, else in a `Settings` folder **beside** the program folder (the user keeps `…\Python Script\Repo` + `…\Python Script\Settings`, so unzipping an update never overwrites settings); two copies → the program-folder one wins and the other is named as ignored; none → `run.py` stops (exit 2) listing where it looked. Loaded by path (`importlib`), `_helper` names skipped. **Git-ignored** (also any `local_settings*.py` variant, and personal `doi_filter_*.py` copies); template in `local_settings.example.py`. Real values are never committed; if asked to upload them, explain the risk (company domains, network paths, staff addresses on GitHub) and get explicit confirmation first |
| `upsert_tracker.py` | Write/update tracker rows |
| `doi_analysis.py` | `detect()` → accept / review / reject + confidence + reasons; `analyze()` → dates by who stated them, items, scope, departments, events, follow-up link, people. Pure Python |
| `attachment_text.py` | `read_text()` for saved attachments: .docx (stdlib), .xlsx (openpyxl), .pdf (optional `pypdf`), .csv (text); anything else / failures → status `not read - …` → review |
| `email_types.py` | Recognises the team's five email templates (`classify()`) and the topic in `DOI Exam Request FIN-53 – Topic` subjects; says which types create rows, carry the due date, or move Status to In Progress |
| `due_dates.py` | `find_due(body, received)`: one clear due date near due wording → date; several dates / ASAP / TBD → note only. Strips quoted reply history first |
| `doi_contacts.py` | Contact directory `DOI_Contacts.csv` (`CONTACTS_PATH`): records every Sender/To/CC on request emails each run; senders are trusted only once the user types `Y` in Approved |
| `publish.py` | Publishes a view-only, Excel-recalculated copy to `PUBLISH_PATH` for coworkers |
| `import_original.py` | One-time: rebuild the original request list (sheet `Requests`) in the tracker template's format → new file + `_Import_Log.csv`. Run locally on real data |
| `run.py` | Entry point: fetch (+ save/read attachments) → detect → analyze → upsert → publish → CSV log (`--days N`, `--dry-run`, `--all` = ignore the memory, `--scheduled` = unattended: crashes → `DOI_Run_Errors.txt` next to the log). **Nothing new → nothing done**: emails already handled (per `DOI_Seen_Messages.csv`) aren't news; if no email is new or changed outcome, the run prints `nothing to do` and writes no tracker/log (it still publishes if the master was hand-edited). Only new/changed emails are printed and logged. Log columns: `run_at, result, request_id, topic, classification, confidence, signals, message_id, conversation_id, current_sender, original_sender, received, ref, subject, exam_due_date, internal_expected_date, tracking_due_date, date_source, date_evidence, send_to, send_to_source, attachments, xlsx_update, detail, review_reason`; an older-format log is renamed `*_Old_Format_<ts>.csv`. `write_log()` creates the CSV + folder on first use; if the log is locked (open in Excel) it writes a dated fallback CSV instead of losing the run |
| `seen_state.py` | Memory of handled emails, `DOI_Seen_Messages.csv` next to `LOG_PATH`: ref, result (`done` / `skipped` / `review` / `reject`), request_id, conversation_id, first/last seen. `done` refs aren't even opened by the fetch; review/reject are re-judged quietly (news only when the outcome changes, e.g. sender approved). Recorded only after a successful tracker save (tracker open in Excel → whole run retried next time); not in `--dry-run`; pruned after `--days`+30. Deleting it is safe (tracker idempotency still prevents duplicates). Its `conversation_id → request_id` map lets later thread replies count toward Last Email Sent |
| `schedule_hourly.py` | Registers Windows Task Scheduler task `DOI Exam Tracker` via `schtasks /XML` (UTF-16): hourly (`PT1H`), only while signed in (`InteractiveToken`, no password, no admin), `StartWhenAvailable` (catch up once after sleep), no wake, runs on battery, `IgnoreNew`, 30 min limit, `pythonw run.py --scheduled`. `--show` prints the XML, `--remove` deletes. `task_xml()` is pure and tested |
| `tests/` | `python -m unittest -v` (test files stay `test_*.py` — unittest discovery looks for that pattern). Upsert tests need a fake-data tracker at `tests/fixtures/DOI_Exam_Test_Tracker_v2.xlsx` or `$DOI_TRACKER_FIXTURE`; skipped otherwise. `*.xlsx` is git-ignored — don't commit workbooks |

## Naming convention

**Saved files** — anything the script writes to the company network / OneDrive — are **Title_Case with underscores**: `DOI_Run_Log.csv` (+ dated fallbacks `DOI_Run_Log_YYYYMMDD_HHMMSS.csv`), `DOI_Contacts.csv`, `DOI_Seen_Messages.csv`, `DOI_Run_Errors.txt`, `Backups/`, `<Tracker>_<timestamp>.xlsx` backups, `*_Import_Log.csv`, `DOI_Exam_Tracker_View_Only.xlsx`, `DOI_Attachments/`. Acronyms stay upper-case (`DOI_`). Saved attachments keep the sender's original file name (evidence fidelity) behind a date prefix, inside a `<RequestID>` folder.

**Python files** follow normal Python style (PEP 8): lower-case `snake_case.py` — `doi_filter.py`, `run.py`, … The convention does not apply to code. `doi_filter` also loads a `Local_Settings.py` (left over from a brief Title_Case rename) with a reminder to rename it back.

## Architecture

```
Outlook (COM) ──Restrict()──> extract() ──> Mail dataclass ──> qualify() ──> upsert_tracker()
   coarse filter in Outlook     only COM-touching fn    plain Python     returns (ok, reason)
```

- COM is confined to small functions in `doi_filter.py`: `extract(msg) -> Mail`, `extract_sent(msg) -> SentMail`, `save_attachments()` / `read_attachments()`, and the `fetch_*` / `_outlook()` wrappers (plus `publish.excel_recalc`). `fetch_candidates()` scans the **Inbox and every subfolder at any depth** (`inbox_folders()`; no folder setting — whatever rules move mail where, it is found), judges each EntryID once, skips `skip_refs` (done emails), ignores folders Outlook can't read, and returns `(accepted, rejected, review)`; attachments of non-rejected emails are saved (or, without `ATTACH_DIR`, put in a temp folder) and read, then the email is detected again with their text. Files saved before an ID is known go to `ATTACH_DIR\_Review\`. Everything downstream works on plain dataclasses so it can be unit-tested with fake objects.
- `qualify(m) -> (bool, str)`: every rejection returns a reason string. Log rejections — that is the audit trail.
- `Items.Restrict()` dialects can't be mixed: chain one Jet filter (`[ReceivedTime] >= '...'`) and one DASL filter (`@SQL=...`).

## Roles (source of truth: `local_settings.py`)

`LEAD_EXAMINERS`, `EXAMINER_CONTACTS`, `EXAMINER_DOMAINS` (+ legacy `SENDER_DOMAINS`), `THIRD_PARTY_CONTACTS`, `THIRD_PARTY_DOMAINS`, `ACCOUNTANT_CONTACTS`, `INTERNAL_DOMAINS`, optional `ENTITY_NAMES`, `PRODUCT_NAMES`. `doi_filter.role_of(address)` → lead_examiner / examiner / third_party / accountant / internal / approved (DOI_Contacts Y) / unknown. **Third party via the thread** (`THIRD_PARTY_FROM_THREAD`, default True): an outside address the **examiner side** put on a request email — To/CC of an email they sent, or the To/Cc lines of a forwarded examiner part (`Segment.recipients`) — is found by `doi_analysis.examiner_included()`, marked `Examiner_Thread = Y` in `DOI_Contacts.csv`, loaded into `THREAD_THIRD_PARTIES` at the start of each run, and then `role_of()` → `third_party`. Per address, never per domain; people a colleague copies are not vouched for (could be the company's own consultant); internal/configured roles always win; `N` in Approved undoes it. Takes effect from the next run — review items are re-judged then. Legacy `SENDER_DOMAINS` is checked **after** accountant/internal, so the company's own domain listed there doesn't turn colleagues into examiners. Domain settings accept whole addresses (`normalize_domains` keeps the part after `@`). `LOG_PATH` may be a folder (`log_file()` adds `DOI_Run_Log.csv`). Examiner side = first three; internal side = accountant, internal. **No names or addresses anywhere else in code.** The user's real `local_settings.py` may name these differently — ask for the variable *names* (never values) and map them.

## Detection (`doi_analysis.detect`)

The five rules below still accept on their own (high). Otherwise signals decide: request ID anywhere (subject → message parts → attachment text), current sender's role, **original sender inside a forward** (`segments()` splits `From:` + `Sent:/To:/Subject:` header blocks; classic Outlook quotes colleagues as `From: Lastname, Firstname` with **no address** — `name_book(m)` resolves such names against the email's own sender/To/CC names (`'Lee, Jordan'` = `'Jordan Lee'`), an Exchange X500 `/O=` address means internal, and an unresolved part is labelled `forwarded part from <name>`, never as the current sender), DOI wording (`DOI_TERMS`). Examiner-side origin + ID + DOI wording → accept high; internal sender + ID + (DOI wording **or** an examiner-side person on To/CC) → accept medium ("FIN-52 support has been uploaded" to the examiner has no DOI wording); ID + DOI wording from an unconfigured sender, or examiner-side with no ID → **review**; no ID and not examiner-side, or an ID with no DOI wording → reject. `MUST_BE_SENT_TO`, if set, is a hard rule for **external** senders only (examiner side / unknown) — a filter, not the response recipient; entries are exact addresses or `@domain` (= anyone at that domain on To/CC, `reaches_required()`; only exact addresses are kept out of Contact); internal forwards and replies are judged on signals (the user is often only CC'd). Review items are logged (result `review`) and never written to the tracker. Outlook pre-filter (`subject_filter`) no longer requires attachments and also matches ID prefixes in the body.

## Dates (`doi_analysis.analyze`)

Separate values: `exam_due_date` (examiner side: own words, forwarded part, or a request-named attachment when the thread is examiner/internal), `internal_expected_date` (anything an internal person states, incl. "we expect … by"), `tracking_due_date` (= exam due, else internal as fallback), `date_source`, `date_confidence` (high = examiner's own words, medium = forwarded/attachment, low = internal fallback), `date_evidence` (the sentence). Several distinct examiner dates, several internal dates, or dates from unconfigured senders → review, nothing chosen. G is written only if empty; a different tracking date → review note, never overwritten; an internal fallback in G is labelled in Notes.

## Send To (`doi_analysis.find_send_to`)

Who a request's **response goes to** — separate from the `MUST_BE_SENT_TO` filter. Read only from the **examiner side's** words (own email or forwarded examiner part, newest first), then request-named attachments: a fixed list of shapes `send|give|provide|submit|forward|return|email|deliver|direct [it|them|the responses|your documents …] [directly|back] to <who>` (fixed on purpose: "information related to Premium Tax" is not an instruction). `<who>` = a written-out address; `me`/`us` = that part's author; a capitalised name looked up among the email's people by `_resolve_name()` (`Jordan` ⊂ `Lee, Jordan`; must be unique) — not found → `Name (address not on the email)` + review. Nothing stated → the lead examiner (`LEAD_EXAMINERS`: the one on the email, or the only one configured). A colleague's "send it to me" is internal collection and ignored. Output: `Analysis.send_to` / `send_to_source`; Notes gets `send to …` when stated, or once on a new row for the default; log columns `send_to`, `send_to_source`.

## Qualification rules (in `doi_filter.py` settings block)

An email qualifies only if **all** pass:
1. Subject contains a Request ID matching `REQUEST_ID_PATTERN` = `\b(?:FIN|IT|DOI-[A-Z]+)-\d+\b` — real IDs are `FIN-##` and `IT-##`; `DOI-TEST-##` is test data only. The user's real pattern also allows `FIN 45` / `FIN45`; every match goes through `doi_filter.find_id()` → `canonical_id()` (`FIN-45`), the only form used for tracker keys, column D, log and memory. Matching is case-sensitive on purpose (`send it 5 days` is not an ID). Never call `REQUEST_ID.search(...).group(0)` directly. The DASL pre-filter only sees hyphenated forms (`%FIN-%`) or the other `SUBJECT_LIKES` phrases. A new prefix must be added to **both** `REQUEST_ID_PATTERN` and `SUBJECT_LIKES` (the Outlook DASL pre-filter, OR'd by `subject_filter()`), or those emails never reach Python
2. Sender's domain in `SENDER_DOMAINS` **or** the sender's address approved (`Y`) in `DOI_Contacts.csv` (passed in as `approved_senders`). Never auto-approve: recording is automatic, trusting is a human decision (else anyone writing `FIN-53` in a subject could whitelist themselves)
3. *Optional:* `MUST_BE_SENT_TO` appears on To or CC — **default empty = rule skipped**, because the To line moves between the examiner and the assigned contact as a request progresses
4. Body contains one of `BODY_PHRASES`
5. At least one attachment whose **file name** contains `DOI`, `Request`, `Follow-up` **or a request ID** (real request documents are named `<Company> Group Exam - FIN-53 …`; the same test decides whether an attachment's dates count as the examiner's) — matched case/space/hyphen/underscore-insensitive via `_norm()`; image extensions (`.png .gif .jpg .jpeg`) ignored. Matching is on file name, not document contents.

## Outlook COM gotchas

- Collections are 1-based. `msg.Class == 43` = MailItem; skip everything else.
- `GetDefaultFolder(6)` = Inbox, `5` = Sent Items; `folder.Folders` = its subfolders (walked recursively, never looked up by name — a missing name would raise).
- Internal Exchange addresses come back as X500 strings. Use `PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)` with fallback to `.Address` (see `smtp_address()`).
- `Recipients[i].Type`: 1 = To, 2 = CC, 3 = BCC.
- `msg.EntryID` is the stable item ID — use as secondary dedupe key.
- Attachments: `.FileName`, `.Item(i)` (1-based), `.SaveAsFile(path)`. With `ATTACH_DIR` set, `fetch_candidates()` saves every `ATTACH_TYPES` attachment (default pdf/doc/docx/png/xlsx; Outlook's inline `image001.png`-style logos skipped) of qualifying emails to `ATTACH_DIR\<RequestID>\<YYYY-MM-DD>_<FileName>` (date prefix keeps revised same-name files apart; existing files skipped = idempotent; Windows-invalid chars → `_`). `attachment_targets()` is the pure, tested plan; `save_attachments()` is the COM part. Each result is logged as `attachment`. `--dry-run` saves nothing.

## Tracker workbook structure

Sheet **DOI Exam Tracker**, Excel Table `tblRequests`, header row 3, data from row 4. Navy headers = input, slate headers = formula.

| Col | Header | Type | Script writes? |
|---|---|---|---|
| A | Request ID | formula (parsed from D) | **no** |
| B | Assigned To | input, dropdown | optional |
| C | Initial Request Received | input, date | yes ← `ReceivedTime` |
| D | Initial Request | input, text | yes ← subject; template subjects become `FIN-53 Request - Topic` |
| E | Type | formula | **no** |
| F | Follows Up On | formula | **no** |
| G | Due Date | input, date | yes, **only into an empty cell** and only one clear date (`due_dates.find_due`); a different stated date is noted in O, never overwritten |
| H | Days Until Due | formula | **no** |
| I | Status | input, dropdown | `"Not Started"` for new rows; **Not Started/blank → In Progress** when an acknowledgment, status update, scope question or response materials email is seen. **→ Completed** (the workbook's own word from Lists!A4↓, `completed_status()`) when an internal person tells the examiner side the support was delivered (`doi_analysis.find_submission()`: current sender internal side, an examiner-side person on To/CC, the sender's **own new words** — not quoted history — say `has been uploaded/submitted/provided…`, `uploaded to OneDrive/SharePoint/portal`, `please find attached`), only from blank / Not Started / In Progress / Awaiting Response; On Hold or any other status a person chose stays, with a note. Partial wording (`remaining`, `will follow`, `part of`…) → note + In Progress, never Completed |
| J | Action Flag | formula | **no** |
| K | Completed On | input, date | yes, **only if empty** ← date of the submission email that completed the request |
| L | Contact | input, text | yes, if blank ← **the person inside the company asked to provide the information**: To line of an internal email that assigns/asks (`upsert_tracker.contact()`), minus the user (`MY_ADDRESSES`, read from Outlook's accounts each run by `own_addresses()`), accountants, the `MUST_BE_SENT_TO` mailbox and the examiner side; CC never counts. Nobody left → stays blank (never the sender, never the examiner). A cell holding only the user's own address (an earlier version's mistake) may be replaced (`only_you()`); anything else typed is never overwritten |
| M | Last Email Sent | input, date | yes ← **the newest email in the request's exchange, whoever sent it** (examiner, accountant, coworker, anyone), among the user's received/CC'd emails about the request (ID in subject, ID found by detection, or same `ConversationID` as an accepted request email) — counts accepted, review **and** rejected messages; plus the user's own Sent Items only if `TRACK_SENT=True` (default False). Newest wins, **only moves forward**, date only, source logged |
| N | Days Since Last Email | formula | **no** |
| O | Notes / Latest Update | input, text | append only, never overwrite |

- **Never write into formula columns** (A, E, F, H, J, N). When appending a row, copy formulas down from the first data row and extend `tblRequests.ref`.
- Request ID formula = first word of column D with `:` stripped (`FIN-52:` → `FIN-52`); shows `(no ID)` if that word has no digit. `upsert_tracker.first_word_id()` is its Python twin — keep them in step. Type = contains "Follow". Follows Up On = ID before `Follow-up`/`Follow Up`, handling `FIN-27` and `FIN 45` → `FIN-45`; blank if the follow-up is to a non-ID (e.g. an interview).
- Live data (since 2026-10-12): 18 real requests, two assignees on Lists. Due dates given as text ("ASAP") go to Notes, never into column G (keeps date math numeric).
- **Known workbook bug:** `StatusList` = `OFFSET(Lists!$A$4,…,COUNTA(Lists!$A$4:$A$60))` also counts the help text that starts at Lists!A14, so the Status dropdown shows blanks and help paragraphs. Fix: count `$A$4:$A$12` instead. Code reads Lists columns only down to the first blank (`import_original.list_block`).
- **Sheet Summary**: settings `C5` AsOf (`=TODAY()`), `C6` FollowUpDays, `C7` DueSoonDays; KPIs and breakdowns via named ranges (`trkStatus`, `trkDue`, `trkRequest`, … all rows 4:1000).
- **Sheet Lists**: dropdown sources via dynamic names `StatusList` (A4↓) and `AssigneeList` (D4↓).

## Team email templates (`email_types.py`)

These templates are **one signal, not the full set** of DOI correspondence; direction is not assumed (the user is usually CC'd). Examples (invented values) live in `tests/fake_templates.py`; broader invented correspondence (direct, forwarded, CC-only, internal payroll request, conflicting dates, attachment-only details, numbered items, scope, follow-ups, events, false positives) in `tests/fake_doi_emails.py`. Classification reads the subject plus the **new** part of the body (quoted history stripped):

| Type | Typical direction | Subject / body marker | Tracker effect |
|---|---|---|---|
| assignment | received | body "has issued examination request" | may create the row; D = `ID Request - Topic`; G from "provide … by"; L = To list |
| acknowledgment | either | body "We received the DOI exam request" | Status → In Progress; restated due date checked against G |
| status_update | sent | subject "Status Update" | Status → In Progress; "expect to provide … by <date>" → note `expects to provide by`, **never G** |
| response_materials | sent | subject "Response Materials" | Status → In Progress; note |
| scope_question | sent | subject "Question Regarding …" / body "confirm the scope" | Status → In Progress; note |
| other | either | — | generic behaviour (may create row, due date from body) |

- Only `assignment` and `other` create rows; any other type for an unknown ID is `skipped` and logged.
- Sent emails are read from Sent Items (`fetch_sent`, now with body + EntryID), matched to a request by ID in subject or ConversationID (`match_sent`), and processed after received emails. Each gets a Notes line `YYYY-MM-DD <Type> sent: <subject> [ref …]`; the ref makes re-runs no-ops. `other` sent emails only count toward Last Email Sent.

## Contact directory (`doi_contacts.py`)

- Each run, for every fetched email whose subject (or detection) has a request ID (accepted or rejected), `observe()` adds/updates Sender/To/CC rows (+ `Forwarded` for addresses only on a quoted examiner message's To/Cc): Email, Name, Domain, Roles, Request_IDs, First_Seen, Last_Seen, Examiner_Thread. Columns **Approved** (Y trust / N never) and **Notes** belong to the user and are never changed — team-meeting contacts are added there by hand (Approved = Y).
- `pending_senders()` = senders not yet approved; listed first in the CSV and counted in the run output. Locked CSV → not saved, logged, re-recorded next run (observation is idempotent). `--dry-run` reports but doesn't save.
- Request progress from meetings goes into the tracker by hand (Status, Notes); the script never overwrites typed values.

## Upsert logic (`upsert_tracker.py`)

- Key = Request ID from subject regex, matched against the first word of column D (same rule as column A's formula). New ID → append row under `tblRequests`: formulas and styles copied from the **first** data row (not the row above — a stray bottom row must not spread), `tbl.ref` + autoFilter extended. D gets RE:/FW: stripped and the ID forced to be the first word.
- Existing ID → C keeps the **earliest** received date; L filled only if blank or only-the-user (never overwrites typed contacts); O gets one appended line per email ending `[ref xxxxxxxx]` (first 8 hex of SHA-1 of EntryID).
- Idempotent via that ref tag: an email whose ref is already in O is skipped. Nothing changed → no backup, no save.
- Timestamped backup to `Backups/` next to the tracker before every save.
- openpyxl requires Excel closed: refuses if `~$<name>` exists or the file can't be opened r+b. Also refuses if the file has charts/images/slicers (openpyxl would delete them).
- Due Date (G): from `find_due()` per email; written only if G is empty; a conflicting stated date, several dates, or ASAP/TBD become text on that email's Notes line.
- Last Email Sent (M) = last email overall in the exchange: `doi_filter.last_exchange_email(fetched, accepted)` over every fetched message, any sender → `{rid: (datetime, source)}`; if `TRACK_SENT`, Sent Items (`fetch_sent` → `last_sent_by_request`) are merged by `merge_last_sent()` (newest wins). → `upsert(last_sent=...)`, whose action text names the source. Updates existing rows only, never creates rows, never moves M backwards. A Sent Items failure is logged and doesn't stop the run. Limitation: an email is only seen if the Outlook pre-filter lets it through (ID prefix in subject/body or DOI wording in subject) and it is in the scanned folder. This is what makes the sheet's FOLLOW UP flag (N ≥ FollowUpDays) meaningful.
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
