# FTIR Reply Automation Bot

Reads FTIR case numbers and pre-written replies from an Excel file (or single CLI test mode), opens the SIFT tracking portal, searches each FTIR, selects "Reply individually.", pastes the reply, and saves the entry.

---

## ⚡ Quick Start (Company Laptop / Microsoft Edge)

If you are using **Microsoft Edge** or **Chrome** on a company laptop with SSO / Login:

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Start Browser with Debugging Port Enabled
1. **Close all open Microsoft Edge windows first**.
2. Open **Command Prompt** (CMD) and start Edge on port 9222:
   ```cmd
   "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222
   ```
   *(If Chrome is used: `"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222`)*
3. Log into the **SIFT Portal** in this browser and keep the **SIFT Main Menu** page open on screen.

### Step 3: Run the Bot

#### Mode A: Test Single FTIR (e.g. `AE20250B00111`)

- **Dry-Run Test** *(pastes and verifies reply without clicking Save)*:
  ```cmd
  python bot.py --ftir AE20250B00111 --attach --dry-run
  ```

- **Live Run** *(pastes reply and clicks **Complete** / Save)*:
  ```cmd
  python bot.py --ftir AE20250B00111 --reply "This is a sample reply message." --attach --live
  ```

#### Mode B: Run Batch Mode from Excel (`FTIR_Replies.xlsx`)

- **Dry-Run Batch**:
  ```cmd
  python bot.py --attach --dry-run
  ```

- **Live Batch**:
  ```cmd
  python bot.py --attach --live
  ```

---

## Prerequisites

- **Python 3.8+**
- **Microsoft Edge** or **Google Chrome**
- No admin rights required; no Docker, no databases, no heavy frameworks.

## Setup

### 1. Clone / copy the files

You need these files in one folder:

```
ftir-reply-bot/
├── bot.py
├── config.example.json
├── config.json          ← you create this (step 3)
├── requirements.txt
└── README.md
```

### 2. Install dependencies

Open a terminal in the project folder (or VS Code terminal) and run:

```bash
pip install -r requirements.txt
```

> If `pip` isn't on PATH, try `python -m pip install -r requirements.txt`.

### 3. Create your config

```bash
copy config.example.json config.json
```

Open `config.json` and fill in:

| Key | What to enter |
|---|---|
| `portal_url` | Full URL of the FTIR portal login page |
| `username` | Your portal login username |
| `password` | Your portal login password |
| `excel_path` | Path to your `.xlsx` file (relative or absolute) |
| `log_file` | Path for the log output (default: `bot.log`) |
| `wait_timeout_seconds` | Max seconds to wait for any page element (default: 15) |
| `paste_retry_count` | Number of paste-verify retries before giving up (default: 3) |
| `headless` | `true` to run without a visible browser, `false` to watch it work |
| `dry_run` | `true` to test without clicking Save (see below), `false` for real runs |
| `delay_between_rows_seconds` | Pause between rows to avoid rate-limiting (default: 2) |
| `verbose_logging` | `true` to log full reply text on paste failures (debugging only) |

### 4. Fill in CSS selectors

Inside `config.json` → `"selectors"`, replace every `<TODO: ...>` placeholder with the actual CSS selector from your portal. To find them:

1. Open the portal in Chrome.
2. Right-click the element → **Inspect**.
3. In the Elements panel, right-click the highlighted HTML → **Copy** → **Copy selector**.
4. Paste it as the value in `config.json`.

**Key selectors you need:**

| Selector key | What it targets |
|---|---|
| `username_field` | Login username `<input>` |
| `password_field` | Login password `<input>` |
| `login_button` | Login submit `<button>` |
| `post_login_element` | Any element only visible after successful login (e.g. a nav menu, dashboard header) |
| `quick_search_box` | The Quick Search `<input>` |
| `search_submit_button` | Search submit button (leave as `""` to use Enter key instead) |
| `results_container` | The `<div>` / `<table>` wrapping search results |
| `result_rows` | Individual result rows (e.g. `table tbody tr`) |
| `result_ftir_text` | The FTIR number text within a result row (e.g. `td:first-child`) |
| `record_header_ftir` | FTIR number shown on the opened record page (for verification) |
| `reply_individually_button` | "Reply Individually" button/link |
| `reply_textarea` | The reply text input (textarea or contenteditable div) |
| `save_button` | Save button on the reply form |
| `save_confirmation` | Element that appears on successful save (e.g. a success toast) |
| `save_error_banner` | Error banner element (or `""` if none) |
| `logout_button` | Logout button (or `""` if the portal has no logout) |

### 5. Prepare your Excel file

The bot expects these columns in the **first row (header)**:

| Column | Description |
|---|---|
| `FTIR Number` | The case number to search |
| `Reply` | The full reply text (multi-line is fine) |
| `Status` | Leave blank or write `Pending` — the bot fills this in |

The bot will update `Status` to `Completed`, `Dry-run OK`, or `Failed: <reason>` after each row, saving immediately.

---

## Before your first real run

Follow these steps **exactly** to safely validate the bot before touching live data:

### Step 1: Fill config and set `dry_run: true`

Make sure `config.json` has:
- All credentials and selectors filled in
- `"dry_run": true` (this is the default in `config.example.json`)
- `"headless": false` (so you can watch what it does)

### Step 2: Prepare a test sheet with 2–3 rows

Create a small Excel file with 2–3 real FTIR numbers and their replies. Set all `Status` cells to blank.

### Step 3: Run the dry-run

```bash
python bot.py
```

The bot will:
1. Show a pre-flight summary (total rows, duplicates, blank cells)
2. Ask you to confirm with `y/n`
3. Log in, search each FTIR, paste the reply, **verify the paste**
4. **NOT click Save** — it will log `DRY RUN: would have saved row X`
5. Mark each row as `Dry-run OK` (distinct from `Completed`)

### Step 4: Review the dry-run log

Open `bot.log` and check:
- Did it find the right records?
- Did paste verification pass on every row?
- Were there any warnings or errors?

### Step 5: Switch to live mode

Once satisfied:

1. Clear the `Status` column in your test sheet (or use a fresh sheet)
2. Set `"dry_run": false` in `config.json`
3. Run `python bot.py` again
4. Confirm `y` at the pre-flight prompt

### Step 6: Scale up

After confirming the live run works on 2–3 rows, point `excel_path` at your full data sheet and run again.

---

## Running the bot

```bash
python bot.py
```

Or in VS Code: open `bot.py`, press `Ctrl+Shift+P` → "Run Python File in Terminal".

## What happens during a run

1. **Pre-flight validation**: Scans the Excel file and reports total/pending/completed/blank/duplicate rows. Requires `y/n` confirmation before opening the browser.
2. Chrome opens (visible unless `headless: true`).
3. Bot logs into the portal.
4. For each pending row (top-to-bottom):
   - Searches the FTIR number in Quick Search.
   - If 2+ results appear, picks the exact match (or flags as `Failed: ambiguous search result`).
   - Opens the record, **verifies the header FTIR matches exactly** (prevents wrong-record paste).
   - Clicks "Reply Individually", Tabs to the text field.
   - Pastes the reply using JavaScript (preserves multi-line formatting).
   - **Reads back the pasted text and verifies it** — retries up to N times on mismatch.
   - **[Dry run]**: Logs "would have saved" and marks `Dry-run OK`.
   - **[Live run]**: Re-verifies header + reply field one more time, clicks Save, waits for confirmation, marks `Completed`.
   - Updates the Excel `Status` column and **saves the file immediately** (crash-safe).
   - Waits `delay_between_rows_seconds` before the next row.
5. Logs out and closes the browser.
6. Prints a summary of completed/failed/skipped rows.

## Re-running after failures

Just run `python bot.py` again. The bot:
- **Skips** rows marked `Completed` or `Dry-run OK`
- **Retries** rows marked `Failed: ...`, `Pending`, or blank

To retry specific failed rows, you can also clear their `Status` cell back to blank.

> **Safety guarantee**: A `Completed` row is never re-processed, so re-runs never risk double-submitting a reply.

## Interrupting the bot

- Press `Ctrl+C` at any time. The bot will:
  - Stop processing immediately
  - Save the Excel file with all progress so far
  - Log out and close the browser
  - Print "stopped at row X, safe to resume"
- The interrupted row's `Status` will **not** be `Completed` (it stays blank/Pending), so a re-run picks it up.

## Output files

| File | Contents |
|---|---|
| Your Excel file | `Status` column updated per row |
| `bot.log` | Timestamped log with redacted FTIR previews and row numbers for audit |

## Data privacy

- **Console output**: Shows row numbers and status only — no FTIR numbers or reply text.
- **Log file**: Shows redacted FTIR previews (e.g., `FTI...45`) and reply hashes — never full content unless `verbose_logging` is `true`.
- **Credentials**: Stored in `config.json` only (in `.gitignore`).
- **Session**: Bot logs out and closes the browser in a `finally` block, even on crash or `Ctrl+C`.
- **Synced folders**: Bot warns on startup if the working directory is inside OneDrive, Dropbox, etc.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `chromedriver` version mismatch | `webdriver-manager` handles this automatically. If it still fails, update Chrome. |
| "Element not interactable" | The CSS selector might be wrong, or the page hasn't loaded. Increase `wait_timeout_seconds`. |
| Paste verification keeps failing | The site may use a framework that ignores `.value` changes. The bot falls back to clipboard paste on attempt 3+. If still failing, inspect the field type — it might be a `contenteditable` div, not a textarea. Set `verbose_logging: true` to see exact expected vs. actual content in the log. |
| Excel file locked / can't save | Close the Excel file in any other program before running the bot. |
| `PermissionError` on Excel save | Same as above — the file is open elsewhere. |
| Pre-flight shows duplicates | Check your Excel for duplicate FTIR numbers — could be a data-entry mistake. |
| "Record header mismatch" | The FTIR number on the opened page doesn't match the Excel row. Check the `record_header_ftir` selector. |

## Security notes

- Credentials are stored in `config.json` (local file, not committed to version control).
- `config.json` and `bot.log` are in `.gitignore`.
- The bot never sends data anywhere except the configured portal URL.
- The bot logs out and destroys the browser session after every run.
