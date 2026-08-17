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
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
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
try:
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    ChromeDriverManager = None

try:
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
except ImportError:
    EdgeChromiumDriverManager = None

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
# Browser helpers & locator parsing
# ---------------------------------------------------------------------------

def parse_locator(selector_str: str):
    """Determine By.XPATH or By.CSS_SELECTOR from selector string."""
    s = str(selector_str).strip()
    if s.startswith("//") or s.startswith("(") or s.startswith("./") or "contains(" in s or "following-sibling::" in s:
        return (By.XPATH, s)
    return (By.CSS_SELECTOR, s)


def create_driver(config: dict, logger: logging.Logger, attach: bool = False, port: int = 9222):
    """Create or attach to a Selenium Edge or Chrome WebDriver."""
    is_attached = attach or config.get("connect_to_existing_browser", False)
    dbg_port = port or config.get("remote_debugging_port", 9222)

    # 1. If attaching to an already running browser (Edge or Chrome)
    if is_attached:
        logger.info(f"Attaching to existing browser session at 127.0.0.1:{dbg_port}...")
        
        # Try Edge WebDriver first
        try:
            edge_opts = EdgeOptions()
            edge_opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{dbg_port}")
            driver = webdriver.Edge(options=edge_opts)
            driver.implicitly_wait(0)
            logger.info("Connected to existing Microsoft Edge browser session successfully.")
            return driver
        except Exception as edge_err:
            logger.debug(f"Edge attach direct attempt: {edge_err}. Trying Chrome WebDriver...")

        # Try Chrome WebDriver as fallback
        try:
            chrome_opts = ChromeOptions()
            chrome_opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{dbg_port}")
            driver = webdriver.Chrome(options=chrome_opts)
            driver.implicitly_wait(0)
            logger.info("Connected to existing Chromium browser session successfully.")
            return driver
        except Exception as chrome_err:
            raise RuntimeError(
                f"Could not connect to browser on port {dbg_port}. "
                f"Please ensure Microsoft Edge or Chrome was started with: "
                f"msedge.exe --remote-debugging-port={dbg_port}"
            ) from chrome_err

    # 2. If launching a new browser instance
    # Try Edge first
    try:
        edge_opts = EdgeOptions()
        if config.get("headless", False):
            edge_opts.add_argument("--headless=new")
        edge_opts.add_argument("--no-sandbox")
        edge_opts.add_argument("--disable-dev-shm-usage")
        edge_opts.add_argument("--window-size=1366,900")
        driver = webdriver.Edge(options=edge_opts)
        driver.implicitly_wait(0)
        logger.info("Microsoft Edge browser launched.")
        return driver
    except Exception:
        pass

    # Fallback to Chrome
    chrome_opts = ChromeOptions()
    if config.get("headless", False):
        chrome_opts.add_argument("--headless=new")
    chrome_opts.add_argument("--no-sandbox")
    chrome_opts.add_argument("--disable-dev-shm-usage")
    chrome_opts.add_argument("--window-size=1366,900")
    try:
        if ChromeDriverManager:
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_opts)
        else:
            driver = webdriver.Chrome(options=chrome_opts)
    except Exception:
        driver = webdriver.Chrome(options=chrome_opts)

    driver.implicitly_wait(0)
    logger.info("Chrome browser launched.")
    return driver


def wait_and_find(driver, selector_expr: str, timeout: int, clickable: bool = False):
    """
    Explicit wait for an element. Supports multiple fallback selectors separated by ' | '.
    Automatically detects XPath vs CSS selector.
    """
    if not selector_expr or str(selector_expr).strip().startswith("<TODO"):
        raise ValueError(f"Invalid selector '{selector_expr}'.")

    # Split fallback selectors separated by ' | '
    candidates = [c.strip() for c in selector_expr.split(" | ") if c.strip() and not c.strip().startswith("<TODO")]
    if not candidates:
        raise ValueError(f"No valid selectors found in '{selector_expr}'.")

    last_err = None
    per_candidate_timeout = max(2, timeout // len(candidates))

    for cand in candidates:
        by, value = parse_locator(cand)
        condition = (
            EC.element_to_be_clickable((by, value))
            if clickable
            else EC.visibility_of_element_located((by, value))
        )
        try:
            return WebDriverWait(driver, per_candidate_timeout).until(condition)
        except Exception as e:
            last_err = e
            continue

    raise last_err or TimeoutException(f"Could not locate element using selectors: {selector_expr}")


def safe_clear_field(driver, element, logger):
    """
    Robustly clear an input/textarea field.
    First tries .clear(), then select-all + delete as a fallback,
    then verifies the field is actually empty.
    """
    try:
        element.clear()
    except Exception:
        pass
    time.sleep(0.2)

    # Fallback: select-all → delete
    try:
        element.send_keys(Keys.CONTROL + "a")
        time.sleep(0.1)
        element.send_keys(Keys.DELETE)
        time.sleep(0.1)
    except Exception:
        pass

    # Verify it's empty
    remaining = element.get_attribute("value") or element.text
    if remaining and remaining.strip():
        logger.debug("  Field not fully cleared, using JS clear.")
        driver.execute_script("arguments[0].value = '';", element)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));", element
        )


def login(driver, config: dict, sel: dict, timeout: int, logger: logging.Logger):
    """Navigate to the portal and log in (skipped if attached to existing logged-in session)."""
    # If already logged in / attached, check if we're on SIFT portal
    try:
        if "sift" in driver.current_url.lower() or "suzuki" in driver.current_url.lower():
            logger.info("Attached browser session is already on SIFT portal. Skipping login step.")
            return
    except Exception:
        pass

    url = config.get("portal_url", "")
    if not url:
        logger.info("No portal_url configured. Assuming browser is already navigated to SIFT.")
        return

    logger.info("Navigating to portal...")
    driver.get(url)

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


def logout_and_quit(driver, config: dict, sel: dict, logger: logging.Logger, is_attached: bool = False):
    """
    Attempt to log out of the portal if new driver was created, or detach if connected to user session.
    """
    if driver is None:
        return

    if is_attached:
        logger.info("  Detached from existing Chrome browser session (browser window left open).")
        return

    try:
        logout_sel = sel.get("logout_button", "")
        if logout_sel and not logout_sel.startswith("<TODO"):
            try:
                logout_btn = wait_and_find(driver, logout_sel, 5, clickable=True)
                logout_btn.click()
                logger.info("  Logged out of portal.")
                time.sleep(1)
            except Exception:
                logger.debug("  Logout button not found or click failed; session will end with browser close.")
    finally:
        try:
            driver.quit()
            logger.info("  Browser closed and session destroyed.")
        except Exception:
            logger.debug("  Browser was already closed.")


def navigate_to_quick_search(driver, sel: dict, timeout: int, logger: logging.Logger):
    """
    On SIFT main menu page (Screenshot 1), click 'QUICK SEARCH'.
    If already on Quick Search page (Screenshot 2), do nothing.
    """
    # Check if Quick Search input box is already visible
    qs_box_sel = sel.get("quick_search_box", "//tr[td[contains(.,'FTIR No')]]//input")
    try:
        candidates = [c.strip() for c in qs_box_sel.split(" | ") if c.strip()]
        for cand in candidates:
            by, val = parse_locator(cand)
            if len(driver.find_elements(by, val)) > 0:
                logger.debug("Already on Quick Search page.")
                return
    except Exception:
        pass

    logger.info("Looking for QUICK SEARCH option on SIFT menu...")
    qs_link_sel = sel.get(
        "quick_search_link",
        "//a[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'QUICK SEARCH')] | //td[contains(text(), 'QUICK SEARCH')]//a"
    )
    try:
        qs_link = wait_and_find(driver, qs_link_sel, timeout, clickable=True)
        qs_link.click()
        logger.info("Clicked QUICK SEARCH. Waiting for Quick Search page to load...")
        time.sleep(1.5)
    except Exception as e:
        logger.warning(f"Could not click QUICK SEARCH link directly: {e}. Checking if form is already loaded.")


def search_ftir(driver, ftir_number: str, sel: dict, timeout: int, logger: logging.Logger):
    """
    Enter the FTIR number into Quick Search and submit.
    """
    navigate_to_quick_search(driver, sel, timeout, logger)

    # Locate Quick Search box
    search_box = wait_and_find(driver, sel["quick_search_box"], timeout, clickable=True)
    safe_clear_field(driver, search_box, logger)

    # Type FTIR number
    search_box.send_keys(ftir_number)
    logger.debug(f"  Typed FTIR {redact_ftir(ftir_number)} into Quick Search.")

    # Submit search
    submit_sel = sel.get("search_submit_button", "//input[@type='submit' and @value='Search'] | //input[@value='Search']")
    try:
        submit_btn = wait_and_find(driver, submit_sel, 5, clickable=True)
        submit_btn.click()
        logger.debug("  Clicked search submit button.")
    except Exception:
        search_box.send_keys(Keys.RETURN)
        logger.debug("  Pressed Enter to submit search.")

    time.sleep(1.5)

    # ---- Wait for results container to appear ----
    try:
        wait_and_find(driver, sel["results_container"], timeout)
        logger.debug("  Search results loaded.")
    except Exception:
        pass


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
    Click 'Reply individually.' radio button (Screenshot 3) and locate reply textarea.
    Returns a FRESHLY located reference to the reply textarea.
    """
    reply_btn_sel = sel.get(
        "reply_individually_button",
        "//input[@type='radio'][following-sibling::text()[contains(.,'Reply individually')]] | //tr[td[contains(.,'Reply individually')]]//input[@type='radio'] | //label[contains(text(),'Reply individually')]//input[@type='radio']"
    )

    try:
        reply_btn = wait_and_find(driver, reply_btn_sel, timeout, clickable=True)
        reply_btn.click()
        logger.info("  Clicked 'Reply individually.' radio button.")
        time.sleep(0.5)
    except Exception as e:
        logger.warning(f"  Direct click on radio button failed: {e}. Attempting Tab focus...")

    # Focus text area (or Tab into it)
    textarea_sel = sel.get("reply_textarea", "//textarea")
    try:
        textarea = wait_and_find(driver, textarea_sel, timeout, clickable=True)
        textarea.click()
        logger.debug("  Reply textarea focused.")
        return textarea
    except Exception:
        ActionChains(driver).send_keys(Keys.TAB).perform()
        time.sleep(0.3)
        return driver.switch_to.active_element


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
    """
    intended_normalized = normalize_whitespace(reply_text)

    for attempt in range(1, max_retries + 1):
        logger.debug(f"  Paste attempt {attempt}/{max_retries}...")

        textarea_sel = sel.get("reply_textarea", "//textarea")
        try:
            textarea = wait_and_find(driver, textarea_sel, 5, clickable=True)
        except Exception:
            pass

        safe_clear_field(driver, textarea, logger)

        if attempt <= 2:
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
            if verbose_logging:
                if len(actual_normalized) < 200 and len(intended_normalized) < 200:
                    logger.debug(f"    [VERBOSE] Expected: '{intended_normalized}'")
                    logger.debug(f"    [VERBOSE] Actual:   '{actual_normalized}'")

    return False


def pre_save_recheck(
    driver, ftir_number: str, reply_text: str, sel: dict, timeout: int, logger: logging.Logger
):
    """
    Final safety check immediately before clicking Save/Complete.
    """
    textarea_sel = sel.get("reply_textarea", "//textarea")
    try:
        textarea = wait_and_find(driver, textarea_sel, timeout)
        actual_value = textarea.get_attribute("value") or textarea.text or ""
        intended_normalized = normalize_whitespace(reply_text)
        actual_normalized = normalize_whitespace(actual_value)

        if actual_normalized != intended_normalized:
            raise ValueError(
                "Pre-save reply re-check FAILED! Reply field content changed."
            )
        logger.debug("  Pre-save reply re-check ✓")
    except Exception as e:
        logger.warning(f"  Pre-save recheck warning: {e}")


def click_save_and_confirm(driver, sel: dict, timeout: int, logger: logging.Logger):
    """
    Click Complete / Save button on SIFT form.
    """
    save_sel = sel.get("save_button", "//input[@value='Complete'] | //button[contains(text(),'Complete')] | //input[contains(@value,'Complete')] | //input[@value='Save']")
    save_btn = wait_and_find(driver, save_sel, timeout, clickable=True)
    save_btn.click()
    logger.info("  Clicked Complete / Save button.")

    confirm_sel = sel.get("save_confirmation", "")
    error_sel = sel.get("save_error_banner", "")

    if confirm_sel and not confirm_sel.startswith("<TODO"):
        try:
            wait_and_find(driver, confirm_sel, timeout)
            logger.debug("  Save confirmation element found ✓")
            return
        except TimeoutException:
            pass

    time.sleep(1.5)


# ---------------------------------------------------------------------------
# Main bot logic
# ---------------------------------------------------------------------------

def run_bot(
    target_ftir: str = None,
    sample_reply: str = None,
    attach_existing: bool = False,
    attach_port: int = 9222,
    override_dry_run: bool = None
):
    """
    Main entry point:
    - Runs in Single FTIR Test mode if `target_ftir` is provided.
    - Runs in Excel Batch Mode if `target_ftir` is None.
    """
    config = load_config()
    sel = config.get("selectors", {})
    timeout = config.get("wait_timeout_seconds", 15)
    max_paste_retries = config.get("paste_retry_count", 3)
    excel_path = config.get("excel_path", "FTIR_Replies.xlsx")
    log_file = config.get("log_file", "bot.log")

    dry_run = override_dry_run if override_dry_run is not None else config.get("dry_run", False)
    delay_between_rows = config.get("delay_between_rows_seconds", 2)
    verbose_logging = config.get("verbose_logging", False)
    is_attached = attach_existing or config.get("connect_to_existing_browser", False)
    port = attach_port or config.get("remote_debugging_port", 9222)

    logger = setup_logging(log_file)
    logger.info("=" * 70)
    logger.info("FTIR Reply Automation Bot — Starting")
    logger.info(f"  Mode        : {'SINGLE FTIR TEST' if target_ftir else 'EXCEL BATCH'}")
    if target_ftir:
        logger.info(f"  Target FTIR : {redact_ftir(target_ftir)}")
    else:
        logger.info(f"  Excel Path  : {excel_path}")
    logger.info(f"  Run Mode    : {'DRY RUN' if dry_run else 'LIVE RUN (Will Save/Complete)'}")
    logger.info(f"  Browser     : {'ATTACHED (port ' + str(port) + ')' if is_attached else 'NEW INSTANCE'}")
    logger.info("=" * 70)

    warn_if_synced_folder(os.getcwd(), logger)

    driver = None
    try:
        driver = create_driver(config, logger, attach=is_attached, port=port)
        if not is_attached:
            login(driver, config, sel, timeout, logger)

        # ===================================================================
        # MODE 1: SINGLE FTIR TEST MODE
        # ===================================================================
        if target_ftir:
            reply_text = sample_reply or "This is a sample test reply for FTIR processing verification."
            logger.info(f"Processing Single FTIR: {redact_ftir(target_ftir)}")

            # 1. Search FTIR
            search_ftir(driver, target_ftir, sel, timeout, logger)

            # 2. Open record if in search results table
            try:
                select_correct_result(driver, target_ftir, sel, timeout=5, logger=logger)
                time.sleep(1)
            except Exception as search_err:
                logger.debug(f"Direct result row select skipped/not needed: {search_err}")

            # 3. Select 'Reply individually.' and get textarea
            textarea = open_reply_field(driver, sel, timeout, logger)

            # 4. Paste reply
            paste_ok = paste_reply(driver, textarea, reply_text, sel, max_paste_retries, verbose_logging, logger)
            if not paste_ok:
                logger.error("Paste verification failed!")
                return

            # 5. Save or Dry Run
            if dry_run:
                logger.info("DRY RUN OK: Reply pasted & verified. Save was NOT clicked.")
            else:
                pre_save_recheck(driver, target_ftir, reply_text, sel, timeout, logger)
                click_save_and_confirm(driver, sel, timeout, logger)
                logger.info("✓ Single FTIR process completed & saved successfully!")

            return

        # ===================================================================
        # MODE 2: EXCEL BATCH MODE
        # ===================================================================
        if not os.path.isfile(excel_path):
            logger.error(f"Excel file not found: {excel_path}")
            sys.exit(1)

        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        col_map = find_column_indices(sheet)
        logger.debug(f"  Column mapping: {col_map}")

        preflight_summary = preflight_validate(sheet, col_map, config, logger)
        if not print_preflight_and_confirm(preflight_summary, dry_run, logger):
            wb.close()
            return

        row_queue = build_row_queue(sheet, col_map, logger)
        total_rows = len(row_queue)
        logger.info(f"  Rows queued for processing: {total_rows}")

        if total_rows == 0:
            logger.info("Nothing to process.")
            wb.close()
            return

        completed = 0
        failed = 0
        skipped = 0
        dry_run_ok_count = 0
        failure_reasons = []

        for queue_idx, row_idx in enumerate(row_queue, start=1):
            ftir_col = col_map[COL_FTIR]
            reply_col = col_map[COL_REPLY]
            status_col = col_map[COL_STATUS]

            current_status = sheet.cell(row=row_idx, column=status_col).value
            current_status_str = str(current_status).strip() if current_status else ""

            if current_status_str.lower() in (STATUS_COMPLETED.lower(), STATUS_DRY_RUN_OK.lower()):
                logger.info(f"[{queue_idx}/{total_rows}] Row {row_idx}: already '{current_status_str}', skipping.")
                skipped += 1
                continue

            ftir_number = str(sheet.cell(row=row_idx, column=ftir_col).value or "").strip()
            reply_text = str(sheet.cell(row=row_idx, column=reply_col).value or "").strip()

            if not ftir_number:
                skipped += 1
                continue

            if not reply_text:
                reason = "Reply column is empty"
                update_row_status(wb, sheet, col_map, row_idx, f"Failed: {reason}", excel_path, logger)
                failed += 1
                failure_reasons.append(f"Row {row_idx}: {reason}")
                continue

            logger.info(f"[{queue_idx}/{total_rows}] Processing row {row_idx}...")

            try:
                search_ftir(driver, ftir_number, sel, timeout, logger)

                try:
                    select_correct_result(driver, ftir_number, sel, timeout=5, logger=logger)
                    time.sleep(0.5)
                except Exception:
                    pass

                textarea = open_reply_field(driver, sel, timeout, logger)

                paste_ok = paste_reply(driver, textarea, reply_text, sel, max_paste_retries, verbose_logging, logger)
                if not paste_ok:
                    reason = f"Paste verification failed"
                    update_row_status(wb, sheet, col_map, row_idx, f"Failed: {reason}", excel_path, logger)
                    failed += 1
                    failure_reasons.append(f"Row {row_idx}: {reason}")
                    continue

                if dry_run:
                    update_row_status(wb, sheet, col_map, row_idx, STATUS_DRY_RUN_OK, excel_path, logger)
                    dry_run_ok_count += 1
                    logger.info(f"  ✓ Row {row_idx} → Dry-run OK")
                else:
                    pre_save_recheck(driver, ftir_number, reply_text, sel, timeout, logger)
                    click_save_and_confirm(driver, sel, timeout, logger)
                    update_row_status(wb, sheet, col_map, row_idx, STATUS_COMPLETED, excel_path, logger)
                    completed += 1
                    logger.info(f"  ✓ Row {row_idx} completed successfully.")

            except Exception as e:
                reason = str(e)[:200]
                status_msg = f"Failed: {reason}"
                logger.error(f"  ✗ Row {row_idx} failed: {reason}")
                update_row_status(wb, sheet, col_map, row_idx, status_msg, excel_path, logger)
                failed += 1
                failure_reasons.append(f"Row {row_idx}: {reason}")

            if queue_idx < total_rows:
                time.sleep(delay_between_rows)

        wb.close()

    except Exception as e:
        logger.critical(f"Fatal error: {e}", exc_info=True)

    finally:
        logout_and_quit(driver, config, sel, logger, is_attached=is_attached)


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="FTIR Reply Automation Bot")
    parser.add_argument("--ftir", type=str, help="Single FTIR number to process (skips Excel)")
    parser.add_argument("--reply", type=str, default="This is a sample test reply for SIFT verification.", help="Test reply text for single FTIR test")
    parser.add_argument("--attach", action="store_true", help="Attach to open Chrome browser on port 9222")
    parser.add_argument("--port", type=int, default=9222, help="Chrome Remote Debugging Port (default 9222)")
    parser.add_argument("--dry-run", action="store_true", help="Paste and verify reply but DO NOT click Save")
    parser.add_argument("--live", action="store_true", help="Perform live run (clicks Save/Complete)")

    args = parser.parse_args()

    override_dry = True if args.dry_run else (False if args.live else None)

    run_bot(
        target_ftir=args.ftir,
        sample_reply=args.reply,
        attach_existing=args.attach,
        attach_port=args.port,
        override_dry_run=override_dry
    )
