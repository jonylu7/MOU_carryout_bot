# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Automates the data carry-out (攜出) approval workflow for a research MOU with Taiwan's Ministry of Finance data center. Each week the system:
1. Scans a Google Drive xlsx for new carry-out requests (`送出申請 = FALSE`)
2. Groups them by `(carry_date × topic)`
3. Optionally uses Claude to format file descriptions
4. Fills a `.docx` template (via `python-docx`)
5. Sends a review email to the coordinator with a HITL approval link
6. Waits (blocking, up to 1 hour) for approve/reject via a local Flask server
7. On approval: creates a carry-out folder, emails all applicants, and marks rows as submitted in Drive

## Running the Project

```bash
pip install -r requirements.txt

# Dry run (read-only, no writes, no email)
python main.py --dry-run

# Live run
python main.py
```

Logs go to `automation.log` in the project directory.

For weekly scheduling via macOS cron (every Monday 9:00 AM):
```
0 9 * * 1 cd "/path/to/carry_out_automation" && python main.py >> automation.log 2>&1
```

## Required Setup Before First Run

1. **Google Service Account**: Place `google_service_account.json` in the project directory. The service account must have **Editor** access to the file on Drive.

2. **`config.py`** – fill in:
   - `DRIVE_FILE_ID` — extracted from either a Google Sheets URL (`spreadsheets/d/{ID}/edit`) or a Drive xlsx URL (`file/d/{ID}/view`); the code auto-detects which type it is
   - `EMAIL_SENDER` / `EMAIL_PASSWORD` (or set env var `EMAIL_PASSWORD`) — use a Gmail App Password, not the account password
   - `ANTHROPIC_API_KEY` (optional) — enables Claude to format file descriptions; without it the system falls back to raw text

3. **Templates** (`.docx`): The four topic-specific Word templates must exist under `BASE_DIR/攜出專區/資料攜出/攜出單/`. Template filenames are configured in `config.TEMPLATES`.

4. **`signature.png`** (optional): Coordinator's signature image (PNG, transparent background, ~300×100 px) placed in the project directory.

## Architecture

All configuration lives in `config.py`. Every module imports `config` directly — there is no dependency injection.

| Module | Responsibility |
|---|---|
| `config.py` | All tuneable parameters (paths, credentials, timeouts, topic maps) |
| `sheets_reader.py` | Downloads xlsx from Drive → `CarryOutRequest` dataclass list; writes back `送出申請=TRUE` after approval |
| `form_filler.py` | Selects the right `.docx` template, does placeholder substitution, fills file rows via `python-docx`, inserts signature image, optionally calls Claude API |
| `hitl_server.py` | Background Flask server (port 8765); `create_review_task()` registers a UUID-keyed task, `wait_for_decision()` **busy-polls** (3 s sleep) until coordinator submits the form |
| `email_sender.py` | Three email types: review request (to coordinator), approval notice (to all filers + coordinator), rejection notice (to filers) |
| `main.py` | Orchestrates all steps; `--dry-run` skips writes/email/server |

## Key Design Decisions

- **HITL is blocking**: `hitl_server.wait_for_decision()` blocks the main thread. The Flask server runs in a daemon thread. Processing one batch at a time is intentional.
- **Claude API is optional**: `form_filler.llm_format_file_rows()` falls back to `_simple_format_file_rows()` if `config.ANTHROPIC_API_KEY` is empty.
- **Date formats**: All dates stored in xlsx use `YYYY/MM/DD`. The system converts to ROC calendar (民國) for `.docx` output using `config.ROC_YEAR_OFFSET = 1911`.
- **Template selection**: `form_filler._select_template()` does substring matching of `topic` against `config.TEMPLATES` keys; unmatched topics fall back to `config.DEFAULT_TEMPLATE`.
- **Placeholder substitution**: `form_filler._replace_in_paragraph()` joins all runs in a paragraph to do the replacement, then writes the result into `runs[0]` and clears the rest — this preserves formatting while handling runs split by Word's internal markup.
- **Credentials**: `EMAIL_PASSWORD` and `ANTHROPIC_API_KEY` should be set as environment variables. The current `config.py` has a hardcoded fallback for `EMAIL_PASSWORD` that should be removed in production.
- **HITL server security**: `HITL_SECRET_TOKEN` is defined but not currently enforced on routes — the review URL is effectively public to anyone with the UUID token on the local network.
- **Google Sheets vs Drive xlsx**: `sheets_reader._is_google_sheet()` calls `files().get(fields="mimeType")` on first access and caches the result in `_file_mimetype_cache`. Google Sheets use `files().export()` to download and Sheets API `batchUpdate` to write back; Drive xlsx files use `get_media` / `files().update()` as before. Both paths require the `spreadsheets` OAuth scope (already in `SCOPES`).
