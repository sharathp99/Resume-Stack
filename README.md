# Gmail LinkedIn Recruiting Inbox Archiver

Production-ready Gmail Inbox archiver for LinkedIn recruiting workflows.

## Features

- Uses **Gmail API** with OAuth Desktop flow (no IMAP).
- Processes only Inbox messages from `jobs-listings@linkedin.com`.
- Handles only subjects:
  - `New application: <ROLE> from <CANDIDATE>`
  - `Your job is posted: [<ROLE>]`
- Supports date filters with Gmail query generation.
- Archives message metadata, bodies, raw RFC822 emails, and resumes.
- Maintains:
  - SQLite dedupe DB (`state.db`)
  - Excel workbook (`candidates.xlsx`)
- Supports `--dry-run` safe testing mode with full simulation output.

---

## Setup

### 1) Create and configure Google Cloud project

1. Go to [Google Cloud Console](https://console.cloud.google.com/).
2. Create/select a project.
3. Enable **Gmail API**.
4. Configure OAuth consent screen.
5. Create OAuth Client ID of type **Desktop app**.
6. Download JSON and save as `credentials.json` in project root.

### 2) Install dependencies

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3) First run OAuth

On first run, browser login will open and generate an access token. In normal mode token cache is saved to `token.json`.

---

## Usage

### Normal run

```bash
python gmail_archive.py --start-date 2026-01-01 --end-date 2026-02-01
```

### Dry run (safe testing)

```bash
python gmail_archive.py --start-date 2026-02-01 --dry-run
```

### Optional flags

- `--per-role-sheets` : write per-role Excel tabs.
- `--credentials <path>` : OAuth client secrets file path.
- `--token <path>` : token cache file path.
- `--log-level DEBUG` : increase verbosity.

---

## Date Filter Rules

- Both `--start-date` and `--end-date`: inclusive range.
- Only `--start-date`: from start-date to now.
- Neither: default last 30 days.

Gmail query uses:

- `after:YYYY/MM/DD`
- `before:YYYY/MM/DD` (exclusive, internally shifted by +1 day for inclusive behavior)

Example generated query:

```text
in:inbox from:jobs-listings@linkedin.com (subject:"New application:" OR subject:"Your job is posted:") after:2026/01/01 before:2026/02/02
```

---

## Output Structure

Root directory:

```text
RecruitingInboxArchive/
  <ROLE>/<YYYY-MM>/
    meta/
    raw/
    bodies/
    resumes/
```

File naming:

- `raw/<YYYY-MM-DD>__msg-<short_hash>.eml`
- `bodies/<YYYY-MM-DD>__msg-<short_hash>.txt`
- `meta/<YYYY-MM-DD>__msg-<short_hash>.json`
- `resumes/<CANDIDATE>__<YYYY-MM-DD>__msg-<short_hash>__<original_filename>`

All filesystem names are sanitized for Windows-forbidden characters, trailing spaces/dots, and bounded length.

---

## Excel Workbook

File:

- `RecruitingInboxArchive/candidates.xlsx`

Sheets:

1. `_all_candidates`
   - `role`
   - `candidate_name`
   - `applied_date`
   - `resume_filename`
   - `gmail_id`

   One row per resume file for application emails.

2. `_role_map`
   - `original_role`
   - `sheet_name`

Optional per-role sheets are generated when `--per-role-sheets` is enabled.

---

## SQLite Dedupe

File:

- `RecruitingInboxArchive/state.db`

Table:

```sql
processed_messages(
  gmail_id TEXT PRIMARY KEY,
  processed_at TEXT
)
```

Messages already in this table are skipped on future runs.

In `--dry-run`, database is never written.

---

## Dry-Run Guarantees

When `--dry-run` is enabled:

- No archive directories are created.
- No files are written.
- No Excel workbook updates.
- No SQLite DB updates.
- No message processed-state persistence.
- OAuth may still be required to fetch Gmail content.

Per email, the script prints what **would** be saved, including final resume names and target folder.

---

## Notes

- This tool intentionally scopes to LinkedIn recruiting sender and subject patterns.
- Gmail scope is read-only (`gmail.readonly`), ensuring no message modifications.
