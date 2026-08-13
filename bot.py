#!/usr/bin/env python3
"""
FTIR Reply Automation Bot — Hardened for Real/Private Data
==========================================================
Reads FTIR numbers and pre-written replies from an Excel file, logs into a web
portal, searches each FTIR, opens the correct record, pastes the reply, saves,
and marks the row as Completed/Failed in the Excel file.

Run:  python bot.py
Config:  config.json (copy from config.example.json and fill in values)

See README.md for full setup and usage instructions.
"""

import hashlib
import json
import logging
import os
import re
import signal
import sys
import time
from datetime import datetime

import openpyxl
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------------------------------------------------------
# Constants — Excel column names (change here if the real sheet differs)
# ---------------------------------------------------------------------------
COL_FTIR = "FTIR Number"
COL_REPLY = "Reply"
COL_STATUS = "Status"

STATUS_COMPLETED = "Completed"
STATUS_PENDING = "Pending"
STATUS_DRY_RUN_OK = "Dry-run OK"

# ---------------------------------------------------------------------------
# --- redaction --- Helpers for safe logging of sensitive data
# ---------------------------------------------------------------------------

def redact_ftir(ftir: str) -> str:
    """
    Return a redacted preview of an FTIR number for log output.
    Shows first 3 and last 2 characters with middle masked.
    Example: 'FTIR-12345-XY' → 'FTI...XY'
    """
    ftir = str(ftir).strip()
    if len(ftir) <= 5:
        return ftir[:2] + "..." + ftir[-1:]
    return ftir[:3] + "..." + ftir[-2:]


def redact_reply(reply: str) -> str:
    """
    Return a redacted preview of reply text for log output.
    Shows first 20 chars + length + SHA-256 hash (for matching without exposing content).
    """
    reply = str(reply).strip()
    preview = reply[:20].replace("\n", "\\n")
    h = hashlib.sha256(reply.encode("utf-8")).hexdigest()[:12]
    return f"'{preview}...' ({len(reply)} chars, hash={h})"


# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

def setup_logging(log_file: str) -> logging.Logger:
    """
    Configure dual logging: console (INFO) + file (DEBUG).
    # --- redaction --- Console output uses INFO level and never includes
    # full FTIR numbers or reply text — only row numbers and status.
    # File handler gets DEBUG level with redacted previews.
    """
    logger = logging.getLogger("ftir_reply_bot")
    logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler — verbose (but still redacted by default;
    # callers use redact_ftir/redact_reply unless verbose_logging is on)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Console handler — info+ only; callers must never log full content at INFO
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


# ---------------------------------------------------------------------------
# Config loader
# ---------------------------------------------------------------------------

def load_config(path: str = "config.json") -> dict:
    """Load config.json. Aborts with a clear message if missing."""
    if not os.path.isfile(path):
        print(f"ERROR: Config file '{path}' not found.")
        print("Copy config.example.json → config.json and fill in your values.")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# --- data safety --- Synced-folder detection
# ---------------------------------------------------------------------------

SYNCED_FOLDER_MARKERS = [
    "onedrive", "dropbox", "google drive", "icloud", "box sync",
    "sharepoint", "teams",
]


def warn_if_synced_folder(working_dir: str, logger: logging.Logger):
    """
    Print a warning if the working directory appears to be inside a
    cloud-synced folder (OneDrive, Dropbox, etc.), which could leak
    config.json / log files to cloud backups.
    """
    path_lower = working_dir.lower().replace("\\", "/")
    for marker in SYNCED_FOLDER_MARKERS:
        if marker in path_lower:
            logger.warning("!" * 70)
            logger.warning(
                f"  WARNING: Working directory appears to be inside a synced "
                f"folder (detected '{marker}' in path)."
            )
            logger.warning(
                "  config.json (credentials) and bot.log (case data) may be "
                "synced to cloud storage!"
            )
            logger.warning(
                "  Consider moving this project to a local, non-synced directory."
            )
            logger.warning("!" * 70)
            return


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def find_column_indices(sheet) -> dict:
    """
    Scan the header row (row 1) to find column indices for our expected columns.
    Returns a dict like {"FTIR Number": 2, "Reply": 3, "Status": 4}.
    Aborts if any required column is missing.
    """
    headers = {}
    for col_idx in range(1, sheet.max_column + 1):
        val = sheet.cell(row=1, column=col_idx).value
        if val:
            headers[str(val).strip()] = col_idx

    missing = [c for c in (COL_FTIR, COL_REPLY, COL_STATUS) if c not in headers]
    if missing:
        raise ValueError(
            f"Excel is missing required column(s): {missing}. "
            f"Found headers: {list(headers.keys())}"
        )
    return headers


def build_row_queue(sheet, col_map: dict, logger: logging.Logger) -> list:
    """
    Build a list of row numbers to process, strictly top-to-bottom.
    Includes rows where Status is blank, 'Pending', or starts with 'Failed:'
    (so re-runs automatically retry failures).
    Skips rows that are 'Completed' or 'Dry-run OK'.
    This is computed once but we re-check Status fresh before processing each row
    (see main loop) to guard against stale-index bugs.
    """
    queue = []
    ftir_col = col_map[COL_FTIR]
    status_col = col_map[COL_STATUS]

    for row_idx in range(2, sheet.max_row + 1):
        ftir_val = sheet.cell(row=row_idx, column=ftir_col).value
        status_val = sheet.cell(row=row_idx, column=status_col).value

        # Skip rows with no FTIR number (empty rows at the bottom of the sheet)
        if not ftir_val or str(ftir_val).strip() == "":
            continue

        status_str = str(status_val).strip() if status_val else ""
        status_lower = status_str.lower()

        # --- resume --- Skip only Completed and Dry-run OK rows; retry everything else
        if status_lower in (STATUS_COMPLETED.lower(), STATUS_DRY_RUN_OK.lower()):
            logger.debug(f"  Skipped row {row_idx}: Status='{status_str}'")
        else:
            queue.append(row_idx)
            logger.debug(f"  Queued row {row_idx}: Status='{status_str}'")

    return queue


def update_row_status(wb, sheet, col_map, row_idx, status_text, excel_path, logger):
    """
    Write a status value to the given row and immediately save the workbook.
    Saving after every row protects against losing progress on crash.
    """
    status_col = col_map[COL_STATUS]
    sheet.cell(row=row_idx, column=status_col, value=status_text)
    try:
        wb.save(excel_path)
        logger.debug(f"  Row {row_idx} → Status='{status_text}', Excel saved.")
    except PermissionError:
        logger.error(
            f"  Could not save Excel (file may be open in another program). "
            f"Row {row_idx} status '{status_text}' is in memory but NOT persisted to disk."
        )


# ---------------------------------------------------------------------------
# --- preflight --- Pre-flight validation (runs before browser opens)
# ---------------------------------------------------------------------------

def preflight_validate(sheet, col_map: dict, config: dict, logger: logging.Logger) -> dict:
    """
    Scan the entire Excel sheet and report:
    - Total rows with FTIR numbers
    - Rows that are already Completed
    - Rows that are Dry-run OK
    - Rows that are pending (blank / Pending / Failed)
    - Rows with blank FTIR Number
    - Rows with blank Reply text
    - Duplicate FTIR numbers
    - Unconfigured CSS selectors (<TODO:...>) in config.json
    Returns a summary dict. Caller decides whether to proceed.
    """
    ftir_col = col_map[COL_FTIR]
    reply_col = col_map[COL_REPLY]
    status_col = col_map[COL_STATUS]

    total = 0
    completed = 0
    dry_run_ok = 0
    pending = 0
    blank_ftir_rows = []
    blank_reply_rows = []
    ftir_seen = {}  # ftir_value → list of row numbers
    issues = []

    # Check config selectors for placeholders
    selectors = config.get("selectors", {})
    todo_selectors = [k for k, v in selectors.items() if not v or str(v).strip().startswith("<TODO")]
    if todo_selectors:
        issues.append(
            f"  CONFIG UNCONFIGURED: The following selectors in config.json still have '<TODO>' placeholders: "
            f"{', '.join(todo_selectors)}"
        )

    for row_idx in range(2, sheet.max_row + 1):
        ftir_val = sheet.cell(row=row_idx, column=ftir_col).value
        reply_val = sheet.cell(row=row_idx, column=reply_col).value
        status_val = sheet.cell(row=row_idx, column=status_col).value

        ftir_str = str(ftir_val).strip() if ftir_val else ""
        reply_str = str(reply_val).strip() if reply_val else ""
        status_str = str(status_val).strip() if status_val else ""

        # Skip completely empty rows (no FTIR, no reply, no status)
        if not ftir_str and not reply_str and not status_str:
            continue

        total += 1

        if status_str.lower() == STATUS_COMPLETED.lower():
            completed += 1
            continue
        if status_str.lower() == STATUS_DRY_RUN_OK.lower():
            dry_run_ok += 1
            continue

        pending += 1

        # Check for blank FTIR
        if not ftir_str:
            blank_ftir_rows.append(row_idx)
            issues.append(f"  Row {row_idx}: FTIR Number is blank")

        # Check for blank Reply
        if ftir_str and not reply_str:
            blank_reply_rows.append(row_idx)
            issues.append(f"  Row {row_idx}: Reply is blank (FTIR={redact_ftir(ftir_str)})")

        # Track duplicates
        if ftir_str:
            ftir_seen.setdefault(ftir_str, []).append(row_idx)

    # Find duplicates
    duplicates = {k: v for k, v in ftir_seen.items() if len(v) > 1}
    for ftir_val, rows in duplicates.items():
        issues.append(
            f"  DUPLICATE: FTIR {redact_ftir(ftir_val)} appears in rows {rows} "
            f"— possible data-entry mistake"
        )

    return {
        "total": total,
        "completed": completed,
        "dry_run_ok": dry_run_ok,
        "pending": pending,
        "blank_ftir_rows": blank_ftir_rows,
        "blank_reply_rows": blank_reply_rows,
        "duplicates": duplicates,
        "issues": issues,
    }


def print_preflight_and_confirm(summary: dict, dry_run: bool, logger: logging.Logger) -> bool:
    """
    Print preflight summary and ask for y/n confirmation.
    Returns True if user confirms, False otherwise.
    """
    logger.info("")
    logger.info("=" * 70)
    logger.info("PRE-FLIGHT VALIDATION")
    logger.info("=" * 70)
    logger.info(f"  Total data rows     : {summary['total']}")
    logger.info(f"  Already Completed   : {summary['completed']} (will be skipped)")
    logger.info(f"  Dry-run OK          : {summary['dry_run_ok']} (will be skipped)")
    logger.info(f"  Pending / to process: {summary['pending']}")
    logger.info(f"  Mode                : {'DRY RUN (Save will NOT be clicked)' if dry_run else 'LIVE RUN'}")

    if summary["issues"]:
        logger.info("")
        logger.warning("  ⚠ Issues found:")
        for issue in summary["issues"]:
            logger.warning(f"    {issue}")

    logger.info("=" * 70)

    if summary["pending"] == 0:
        logger.info("No pending rows to process. Exiting.")
        return False

    # --- preflight --- Require manual y/n confirmation before proceeding
    print()  # blank line for readability
    prompt = "Proceed? (y/n): "
    if dry_run:
        prompt = "Proceed with DRY RUN? (y/n): "
    try:
        answer = input(prompt).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        logger.info("User cancelled at pre-flight confirmation.")
        return False

    if answer != "y":
        logger.info("User declined to proceed. Exiting.")
        return False

    logger.info("User confirmed. Starting bot...")
    return True


# ---------------------------------------------------------------------------
# Whitespace-normalized comparison for paste verification
# ---------------------------------------------------------------------------

def normalize_whitespace(text: str) -> str:
    """
    Collapse all whitespace runs to single spaces and strip.
    Used to compare intended vs. actual paste content, because browsers
    sometimes alter internal spacing in textareas.
    """
    return re.sub(r"\s+", " ", text).strip()


# ---------------------------------------------------------------------------
# Browser helpers
# ---------------------------------------------------------------------------

def create_driver(config: dict, logger: logging.Logger) -> webdriver.Chrome:
    """Create and return a Selenium Chrome WebDriver."""
    options = Options()
    if config.get("headless", False):
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1366,900")
    # Disable automation flags that some portals detect
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception:
        logger.warning("webdriver-manager failed to resolve chromedriver; falling back to PATH.")
        driver = webdriver.Chrome(options=options)

    driver.implicitly_wait(0)  # We use explicit waits everywhere
    logger.info("Chrome browser launched.")
    return driver


def wait_and_find(driver, selector: str, timeout: int, clickable: bool = False):
    """
    Explicit wait for an element to be present AND visible (optionally clickable).
    Returns the WebElement. Raises TimeoutException or ValueError on failure.
    """
    if not selector or str(selector).strip().startswith("<TODO"):
        raise ValueError(
            f"Invalid CSS selector '{selector}'. Please update your config.json with actual "
            f"element selectors from your portal."
        )

    condition = (
        EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        if clickable
        else EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
    )
    return WebDriverWait(driver, timeout).until(condition)


def safe_clear_field(driver, element, logger):
    """
    Robustly clear an input/textarea field.
    First tries .clear(), then select-all + delete as a fallback,
    then verifies the field is actually empty.
    """
    element.clear()
    time.sleep(0.2)

    # Fallback: select-all → delete
    element.send_keys(Keys.CONTROL + "a")
    time.sleep(0.1)
    element.send_keys(Keys.DELETE)
    time.sleep(0.1)

    # Verify it's empty
    remaining = element.get_attribute("value") or element.text
    if remaining and remaining.strip():
        logger.debug("  Field not fully cleared, using JS clear.")
        driver.execute_script("arguments[0].value = '';", element)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));", element
        )


def login(driver, config: dict, sel: dict, timeout: int, logger: logging.Logger):
    """Navigate to the portal and log in."""
    url = config["portal_url"]
    logger.info("Navigating to portal...")
    driver.get(url)

    # Wait for username field
    user_field = wait_and_find(driver, sel["username_field"], timeout, clickable=True)
    safe_clear_field(driver, user_field, logger)
    user_field.send_keys(config["username"])
    logger.debug("  Entered username.")

    pwd_field = wait_and_find(driver, sel["password_field"], timeout, clickable=True)
    safe_clear_field(driver, pwd_field, logger)
    pwd_field.send_keys(config["password"])
    logger.debug("  Entered password.")

    login_btn = wait_and_find(driver, sel["login_button"], timeout, clickable=True)
    login_btn.click()
    logger.info("  Clicked login button. Waiting for post-login element...")

    wait_and_find(driver, sel["post_login_element"], timeout)
    logger.info("  Login successful.")


# --- data safety --- Explicit logout before closing browser
def logout_and_quit(driver, config: dict, sel: dict, logger: logging.Logger):
    """
    Attempt to log out of the portal, then quit the browser.
    Called in a finally block to ensure the session is never left active.
    If logout fails (no selector, page error), we still quit the browser
    which destroys the session cookies.
    """
    if driver is None:
        return

    try:
        logout_sel = sel.get("logout_button", "")
        if logout_sel and not logout_sel.startswith("<TODO"):
            try:
                logout_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, logout_sel))
                )
                logout_btn.click()
                logger.info("  Logged out of portal.")
                time.sleep(1)
            except Exception:
                logger.debug("  Logout button not found or click failed; session will end with browser close.")
        else:
            logger.debug("  No logout_button selector configured; session ends with browser close.")
    finally:
        try:
            driver.quit()
            logger.info("  Browser closed and session destroyed.")
        except Exception:
            logger.debug("  Browser was already closed.")


def search_ftir(driver, ftir_number: str, sel: dict, timeout: int, logger: logging.Logger):
    """
    Enter the FTIR number into Quick Search and submit.
    Returns once the results container is visible.
    """
    # ---- Locate Quick Search box (must be present AND interactable) ----
    search_box = wait_and_find(driver, sel["quick_search_box"], timeout, clickable=True)

    # ---- Clear the field robustly ----
    safe_clear_field(driver, search_box, logger)

    # ---- Type the FTIR number ----
    search_box.send_keys(ftir_number)
    # --- redaction --- Only log redacted FTIR to file, not console
    logger.debug(f"  Typed FTIR {redact_ftir(ftir_number)} into Quick Search.")

    # ---- Submit: click button if selector given, else press Enter ----
    submit_sel = sel.get("search_submit_button", "")
    if submit_sel and not submit_sel.startswith("<TODO"):
        submit_btn = wait_and_find(driver, submit_sel, timeout, clickable=True)
        submit_btn.click()
        logger.debug("  Clicked search submit button.")
    else:
        search_box.send_keys(Keys.RETURN)
        logger.debug("  Pressed Enter to submit search.")

    # ---- Wait for results container to appear ----
    wait_and_find(driver, sel["results_container"], timeout)
    logger.debug("  Search results loaded.")


def select_correct_result(
    driver, ftir_number: str, sel: dict, timeout: int, logger: logging.Logger
):
    """
    After search results load, find and click the row whose FTIR number
    exactly matches `ftir_number`.

    If there's exactly one result, we still verify the text before clicking.
    If there are two+ results and we can't find an exact match, we raise
    an exception so the row is flagged as 'Failed: ambiguous search result'.
    """
    # Find all result rows
    rows = driver.find_elements(By.CSS_SELECTOR, sel["result_rows"])
    logger.debug(f"  Found {len(rows)} search result row(s).")

    if len(rows) == 0:
        raise NoSuchElementException(
            f"No search results found for FTIR '{redact_ftir(ftir_number)}'."
        )

    # Try to find the exact-match row
    matched_row = None
    ftir_normalized = ftir_number.strip().lower()

    for row in rows:
        try:
            ftir_cell = row.find_element(By.CSS_SELECTOR, sel["result_ftir_text"])
            cell_text = (ftir_cell.text or "").strip().lower()
            if cell_text == ftir_normalized:
                matched_row = row
                break
        except (NoSuchElementException, StaleElementReferenceException):
            continue

    if matched_row is None:
        if len(rows) == 1:
            # Single result but text didn't match exactly — risky but log a warning
            logger.warning(
                "  Single result found but FTIR text didn't match exactly. "
                "Clicking it anyway; verify selectors are correct."
            )
            matched_row = rows[0]
        else:
            raise ValueError(
                f"Ambiguous search result: {len(rows)} rows found, "
                f"none matched FTIR exactly."
            )

    # Click the matched row to open the record
    matched_row.click()
    logger.debug("  Clicked matching result row.")


# --- record match safety --- Exact-match header verification
def verify_record_header(driver, ftir_number: str, sel: dict, timeout: int, logger: logging.Logger):
    """
    After navigating into a record, verify the page header / title area
    shows the expected FTIR number. This is the critical data-integrity
    check that prevents pasting a reply onto the wrong record.

    Uses EXACT string match (not substring/contains) to prevent false
    positives on similar numbers (e.g. 'FTIR-123' matching 'FTIR-1234').
    """
    header_sel = sel.get("record_header_ftir", "")
    if not header_sel or header_sel.startswith("<TODO"):
        logger.warning(
            "  record_header_ftir selector not configured — "
            "skipping record verification (DATA INTEGRITY RISK)."
        )
        return

    header_el = wait_and_find(driver, header_sel, timeout)
    header_text = (header_el.text or "").strip()

    # --- record match safety --- EXACT match, not substring 'in' check.
    # Compare normalized (case-insensitive, stripped) values.
    if header_text.lower() != ftir_number.strip().lower():
        raise ValueError(
            f"Record header mismatch! Expected FTIR exact match but page shows "
            f"different value. Aborting this row (no save attempted)."
        )
    logger.debug("  Record header verified (exact match) ✓")


def open_reply_field(driver, sel: dict, timeout: int, logger: logging.Logger):
    """
    Click 'Reply Individually' and then Tab into the actual text field.
    Returns a FRESHLY located reference to the reply textarea.
    """
    reply_btn_sel = sel.get("reply_individually_button", "")
    if reply_btn_sel and not reply_btn_sel.startswith("<TODO"):
        reply_btn = wait_and_find(driver, reply_btn_sel, timeout, clickable=True)
        reply_btn.click()
        logger.debug("  Clicked 'Reply Individually' button.")
        time.sleep(0.5)  # Brief pause for the reply form to expand/render

    # Press Tab to focus the actual input field (mirrors the manual process)
    ActionChains(driver).send_keys(Keys.TAB).perform()
    time.sleep(0.3)

    # Now locate the reply textarea fresh
    textarea = wait_and_find(driver, sel["reply_textarea"], timeout, clickable=True)
    logger.debug("  Reply textarea located.")
    return textarea


def paste_reply(
    driver,
    textarea,
    reply_text: str,
    sel: dict,
    max_retries: int,
    verbose_logging: bool,
    logger: logging.Logger,
) -> bool:
    """
    Paste the reply text into the textarea with verification.

    Strategy (chosen to preserve multi-line / blank-line formatting):
    ─────────────────────────────────────────────────────────────────
    1. PRIMARY: Use JavaScript to set .value directly and dispatch
       'input' + 'change' events so the site's JS listeners fire.
       This avoids all send_keys() issues with newlines.

    2. FALLBACK: If JS-set value doesn't stick (some frameworks like
       React/Angular ignore .value changes), use clipboard paste
       (pyperclip + Ctrl+V) which preserves formatting exactly.

    3. VERIFY: After each attempt, read back the field value and
       compare (whitespace-normalized) against the intended text.
       Retry up to `max_retries` times before giving up.
    ─────────────────────────────────────────────────────────────────

    Returns True if paste was verified successfully, False otherwise.
    """
    intended_normalized = normalize_whitespace(reply_text)

    for attempt in range(1, max_retries + 1):
        logger.debug(f"  Paste attempt {attempt}/{max_retries}...")

        # ── Re-locate the textarea fresh each retry to avoid stale refs ──
        textarea = driver.find_element(By.CSS_SELECTOR, sel["reply_textarea"])

        # ── Clear any existing content ──
        safe_clear_field(driver, textarea, logger)

        if attempt <= 2:
            # ── PRIMARY: JavaScript .value set ──
            # Escape backslashes and backticks for the JS template literal
            escaped = reply_text.replace("\\", "\\\\").replace("`", "\\`")
            driver.execute_script(
                f"""
                var el = arguments[0];
                el.value = `{escaped}`;
                el.dispatchEvent(new Event('input', {{bubbles: true}}));
                el.dispatchEvent(new Event('change', {{bubbles: true}}));
                """,
                textarea,
            )
            logger.debug("  Used JavaScript to set textarea value.")
        else:
            # ── FALLBACK: Clipboard paste (preserves exact formatting) ──
            try:
                pyperclip.copy(reply_text)
                textarea.click()
                time.sleep(0.1)
                ActionChains(driver).key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
                logger.debug("  Used clipboard paste (Ctrl+V) as fallback.")
                time.sleep(0.3)
            except pyperclip.PyperclipException as e:
                logger.warning(f"  Clipboard paste failed: {e}. Falling back to send_keys.")
                textarea.send_keys(reply_text)

        # ── VERIFY: Read back the field value ──
        time.sleep(0.3)
        actual_value = textarea.get_attribute("value") or textarea.text or ""
        actual_normalized = normalize_whitespace(actual_value)

        if actual_normalized == intended_normalized:
            logger.debug("  Paste verified ✓")
            return True
        else:
            logger.warning(
                f"  Paste verification FAILED (attempt {attempt}). "
                f"Expected ({len(intended_normalized)} chars) vs "
                f"Actual ({len(actual_normalized)} chars)."
            )
            # --- redaction --- Only log full content if verbose_logging is enabled
            if verbose_logging:
                if len(actual_normalized) < 200 and len(intended_normalized) < 200:
                    logger.debug(f"    [VERBOSE] Expected: '{intended_normalized}'")
                    logger.debug(f"    [VERBOSE] Actual:   '{actual_normalized}'")

    return False


# --- record match safety --- Pre-save re-verification
def pre_save_recheck(
    driver, ftir_number: str, reply_text: str, sel: dict, timeout: int, logger: logging.Logger
):
    """
    Final safety check immediately before clicking Save:
    1. Re-read the record header FTIR and confirm it still matches.
    2. Re-read the reply field and confirm it still contains the intended text.

    This protects against the page changing state between paste and save
    (e.g., AJAX reload, accidental navigation, session timeout).
    Raises ValueError if anything doesn't match — caller will abort the row.
    """
    # --- record match safety --- Re-verify header FTIR (exact match)
    header_sel = sel.get("record_header_ftir", "")
    if header_sel and not header_sel.startswith("<TODO"):
        try:
            header_el = wait_and_find(driver, header_sel, timeout)
            header_text = (header_el.text or "").strip()
            if header_text.lower() != ftir_number.strip().lower():
                raise ValueError(
                    "Pre-save header re-check FAILED! Record may have changed. "
                    "Aborting save."
                )
            logger.debug("  Pre-save header re-check ✓")
        except TimeoutException:
            raise ValueError(
                "Pre-save header re-check FAILED! Header element not found. "
                "Aborting save."
            )

    # --- record match safety --- Re-verify reply field content
    try:
        textarea = driver.find_element(By.CSS_SELECTOR, sel["reply_textarea"])
        actual_value = textarea.get_attribute("value") or textarea.text or ""
        intended_normalized = normalize_whitespace(reply_text)
        actual_normalized = normalize_whitespace(actual_value)

        if actual_normalized != intended_normalized:
            raise ValueError(
                "Pre-save reply re-check FAILED! Reply field content has changed "
                "since paste was verified. Aborting save."
            )
        logger.debug("  Pre-save reply re-check ✓")
    except NoSuchElementException:
        raise ValueError(
            "Pre-save reply re-check FAILED! Reply field not found. Aborting save."
        )


def click_save_and_confirm(driver, sel: dict, timeout: int, logger: logging.Logger):
    """
    Click the Save button and wait for either a success confirmation element
    or verify no error banner appeared.
    """
    save_btn = wait_and_find(driver, sel["save_button"], timeout, clickable=True)
    save_btn.click()
    logger.debug("  Clicked Save button.")

    # ── Check for success confirmation ──
    confirm_sel = sel.get("save_confirmation", "")
    error_sel = sel.get("save_error_banner", "")

    if confirm_sel and not confirm_sel.startswith("<TODO"):
        try:
            wait_and_find(driver, confirm_sel, timeout)
            logger.debug("  Save confirmation element found ✓")
            return
        except TimeoutException:
            raise TimeoutException(
                f"Save confirmation element did not appear within {timeout}s."
            )

    # If no confirmation selector, at least check for error banner
    if error_sel and not error_sel.startswith("<TODO"):
        time.sleep(1)  # Brief wait for any error banner to render
        error_elements = driver.find_elements(By.CSS_SELECTOR, error_sel)
        if error_elements and any(e.is_displayed() for e in error_elements):
            error_text = error_elements[0].text[:200]
            raise RuntimeError(f"Save error banner appeared: '{error_text}'")

    # Neither selector configured — log a warning
    logger.warning(
        "  No save_confirmation or save_error_banner selector configured. "
        "Assuming save succeeded (configure selectors for reliability)."
    )
    time.sleep(1)


# ---------------------------------------------------------------------------
# Main bot logic
# ---------------------------------------------------------------------------

def run_bot():
    """Main entry point: load config, open browser, process each row."""
    # ── Load config ──
    config = load_config()
    sel = config.get("selectors", {})
    timeout = config.get("wait_timeout_seconds", 15)
    max_paste_retries = config.get("paste_retry_count", 3)
    excel_path = config.get("excel_path", "FTIR_Replies.xlsx")
    log_file = config.get("log_file", "bot.log")

    # --- dry run --- Read dry_run flag from config
    dry_run = config.get("dry_run", False)

    # --- rate/pace --- Configurable delay between rows
    delay_between_rows = config.get("delay_between_rows_seconds", 2)

    # --- redaction --- Verbose logging flag (full content in log file)
    verbose_logging = config.get("verbose_logging", False)

    logger = setup_logging(log_file)
    logger.info("=" * 70)
    logger.info("FTIR Reply Automation Bot — Starting")
    logger.info(f"  Excel : {excel_path}")
    logger.info(f"  Log   : {log_file}")
    logger.info(f"  Mode  : {'DRY RUN' if dry_run else 'LIVE'}")
    logger.info(f"  Headless : {config.get('headless', False)}")
    logger.info(f"  Delay between rows : {delay_between_rows}s")
    logger.info(f"  Verbose logging    : {verbose_logging}")
    logger.info("=" * 70)

    # --- data safety --- Warn if working directory is inside a synced folder
    warn_if_synced_folder(os.getcwd(), logger)

    # ── Open Excel ──
    if not os.path.isfile(excel_path):
        logger.error(f"Excel file not found: {excel_path}")
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active
    col_map = find_column_indices(sheet)
    logger.debug(f"  Column mapping: {col_map}")

    # --- preflight --- Run pre-flight validation before opening browser
    preflight_summary = preflight_validate(sheet, col_map, config, logger)
    if not print_preflight_and_confirm(preflight_summary, dry_run, logger):
        wb.close()
        return

    # ── Build processing queue ──
    row_queue = build_row_queue(sheet, col_map, logger)
    total_rows = len(row_queue)
    logger.info(f"  Rows queued for processing: {total_rows}")

    if total_rows == 0:
        logger.info("Nothing to process.")
        wb.close()
        return

    # ── Counters ──
    completed = 0
    failed = 0
    skipped = 0
    dry_run_ok_count = 0
    failure_reasons = []
    last_processed_row = None

    # ── Launch browser & login ──
    driver = None
    try:
        driver = create_driver(config, logger)
        login(driver, config, sel, timeout, logger)

        # ── Process each row ──
        for queue_idx, row_idx in enumerate(row_queue, start=1):
            ftir_col = col_map[COL_FTIR]
            reply_col = col_map[COL_REPLY]
            status_col = col_map[COL_STATUS]

            # ── Re-check Status fresh (guard against stale-index bugs) ──
            # --- resume --- Never re-process Completed or Dry-run OK rows
            current_status = sheet.cell(row=row_idx, column=status_col).value
            current_status_str = str(current_status).strip() if current_status else ""

            if current_status_str.lower() in (
                STATUS_COMPLETED.lower(),
                STATUS_DRY_RUN_OK.lower(),
            ):
                logger.info(
                    f"[{queue_idx}/{total_rows}] Row {row_idx}: "
                    f"already '{current_status_str}', skipping."
                )
                skipped += 1
                continue

            ftir_number = str(sheet.cell(row=row_idx, column=ftir_col).value or "").strip()
            reply_text = str(sheet.cell(row=row_idx, column=reply_col).value or "").strip()

            if not ftir_number:
                logger.warning(f"[{queue_idx}/{total_rows}] Row {row_idx}: empty FTIR, skipping.")
                skipped += 1
                continue

            if not reply_text:
                reason = "Reply column is empty"
                logger.warning(f"[{queue_idx}/{total_rows}] Row {row_idx}: {reason}")
                update_row_status(wb, sheet, col_map, row_idx, f"Failed: {reason}", excel_path, logger)
                failed += 1
                failure_reasons.append(f"Row {row_idx}: {reason}")
                continue

            # --- redaction --- Console gets row number only; file gets redacted FTIR
            logger.info(f"[{queue_idx}/{total_rows}] Processing row {row_idx}...")
            logger.debug(f"  FTIR={redact_ftir(ftir_number)}, Reply={redact_reply(reply_text)}")
            last_processed_row = row_idx

            try:
                # ── 4a-d: Search for the FTIR number ──
                search_ftir(driver, ftir_number, sel, timeout, logger)

                # ── 4e: Select the correct result (handle ambiguous results) ──
                select_correct_result(driver, ftir_number, sel, timeout, logger)

                # ── Wait for record page to load ──
                time.sleep(0.5)

                # ── Verify the record header shows the right FTIR (exact match) ──
                verify_record_header(driver, ftir_number, sel, timeout, logger)

                # ── 4f-g: Open Reply Individually & Tab to field ──
                textarea = open_reply_field(driver, sel, timeout, logger)

                # ── 4h-i: Paste reply with verification ──
                paste_ok = paste_reply(
                    driver, textarea, reply_text, sel, max_paste_retries,
                    verbose_logging, logger,
                )
                if not paste_ok:
                    reason = f"Paste verification failed after {max_paste_retries} attempts"
                    logger.error(f"  {reason}")
                    update_row_status(
                        wb, sheet, col_map, row_idx, f"Failed: {reason}", excel_path, logger
                    )
                    failed += 1
                    failure_reasons.append(f"Row {row_idx}: {reason}")
                    continue

                # --- dry run --- In dry-run mode, skip Save entirely
                if dry_run:
                    logger.info(
                        f"  DRY RUN: would have saved row {row_idx}. "
                        f"Reply pasted and verified successfully, but Save was NOT clicked."
                    )
                    update_row_status(
                        wb, sheet, col_map, row_idx, STATUS_DRY_RUN_OK, excel_path, logger
                    )
                    dry_run_ok_count += 1
                    logger.info(f"  ✓ Row {row_idx} → Dry-run OK")
                else:
                    # --- record match safety --- Final pre-save re-verification
                    pre_save_recheck(driver, ftir_number, reply_text, sel, timeout, logger)

                    # ── 4j-k: Click Save and confirm ──
                    click_save_and_confirm(driver, sel, timeout, logger)

                    # ── 4l: Mark as Completed ──
                    # --- crash safety --- Status is set to Completed ONLY after
                    # Save has been clicked AND confirmed. If the bot crashes
                    # between paste and save, the row stays Pending/blank and
                    # will be retried on the next run.
                    update_row_status(
                        wb, sheet, col_map, row_idx, STATUS_COMPLETED, excel_path, logger
                    )
                    completed += 1
                    logger.info(f"  ✓ Row {row_idx} completed successfully.")

            except KeyboardInterrupt:
                # --- crash safety --- Re-raise to be caught by the outer handler
                raise

            except Exception as e:
                reason = str(e)[:200]
                # Categorize common failures for clearer status messages
                if "ambiguous" in reason.lower():
                    status_msg = "Failed: ambiguous search result"
                elif "mismatch" in reason.lower() or "pre-save" in reason.lower():
                    # --- record match safety --- Hard stop on mismatch, no fallback
                    status_msg = "Failed: record header mismatch"
                elif "no search results" in reason.lower():
                    status_msg = "Failed: no search results"
                else:
                    status_msg = f"Failed: {reason}"

                logger.error(f"  ✗ Row {row_idx} failed: {reason}")
                logger.debug(f"  Full exception:", exc_info=True)
                update_row_status(wb, sheet, col_map, row_idx, status_msg, excel_path, logger)
                failed += 1
                failure_reasons.append(f"Row {row_idx}: {reason}")

                # Try to navigate back to a clean state for the next row
                try:
                    driver.get(config["portal_url"])
                    wait_and_find(driver, sel["post_login_element"], timeout)
                except Exception:
                    logger.warning("  Could not navigate back to portal home. Re-logging in...")
                    try:
                        login(driver, config, sel, timeout, logger)
                    except Exception as login_err:
                        logger.critical(f"  Re-login failed. Stopping bot.")
                        break

            # --- rate/pace --- Delay between rows to avoid hammering the portal
            if queue_idx < total_rows:
                logger.debug(f"  Waiting {delay_between_rows}s before next row...")
                time.sleep(delay_between_rows)

    except KeyboardInterrupt:
        # --- crash safety --- Handle Ctrl+C gracefully
        logger.info("")
        logger.warning("=" * 70)
        logger.warning("  INTERRUPTED by user (Ctrl+C)")
        if last_processed_row:
            logger.warning(f"  Last row being processed: {last_processed_row}")
            logger.warning(
                f"  That row's status was NOT set to Completed (safe to resume)."
            )
        logger.warning("  Saving Excel and closing browser...")
        logger.warning("=" * 70)

    except Exception as e:
        logger.critical(f"Fatal error: {e}", exc_info=True)

    finally:
        # --- data safety --- Always log out, close browser, and save Excel
        logout_and_quit(driver, config, sel, logger)

        # Save Excel one final time to capture any in-memory status updates
        try:
            wb.save(excel_path)
            logger.debug("  Final Excel save completed.")
        except Exception:
            logger.error("  Could not perform final Excel save.")
        wb.close()

    # ── Summary ──
    logger.info("")
    logger.info("=" * 70)
    logger.info("FTIR Reply Automation Bot — Run Summary")
    logger.info("=" * 70)
    logger.info(f"  Total rows in queue : {total_rows}")
    if dry_run:
        logger.info(f"  Dry-run OK          : {dry_run_ok_count}")
    else:
        logger.info(f"  Completed           : {completed}")
    logger.info(f"  Failed              : {failed}")
    logger.info(f"  Skipped             : {skipped}")
    if failure_reasons:
        logger.info("  Failure details:")
        for reason in failure_reasons:
            logger.info(f"    - {reason}")
    if last_processed_row:
        logger.info(f"  Last row touched    : {last_processed_row}")
    logger.info("=" * 70)

    if dry_run and dry_run_ok_count > 0:
        logger.info(
            "  Dry run complete. Review bot.log, then set dry_run=false in "
            "config.json for the real run."
        )
    logger.info("Done.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    run_bot()
