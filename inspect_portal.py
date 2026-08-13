#!/usr/bin/env python3
"""
Suzuki SIFT Portal Selector Detector
====================================
Launches Chrome to https://sift.bizapps.suzuki/sift/ and automatically detects
element IDs and CSS selectors for login inputs, search boxes, and textareas,
then updates config.json automatically.

Run:  py inspect_portal.py
"""

import json
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

CONFIG_PATH = "config.json"

def detect_selectors():
    if not os.path.exists(CONFIG_PATH):
        print(f"ERROR: {CONFIG_PATH} not found. Please create it first.")
        sys.exit(1)

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        config = json.load(f)

    url = config.get("portal_url", "https://sift.bizapps.suzuki/sift/")
    print(f"Launching Chrome browser to: {url}")
    print("Please perform any login manually if prompted...")

    options = Options()
    options.add_argument("--window-size=1366,900")
    
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except Exception:
        driver = webdriver.Chrome(options=options)

    driver.get(url)
    time.sleep(3)

    print("\n--- Scanning page for input elements ---")
    inputs = driver.find_elements(By.CSS_SELECTOR, "input, textarea, button, a")
    
    detected = {}
    for i, el in enumerate(inputs):
        tag = el.tag_name
        el_id = el.get_attribute("id")
        name = el.get_attribute("name")
        el_type = el.get_attribute("type")
        placeholder = el.get_attribute("placeholder")
        text = el.text[:30].strip()

        desc = f"<{tag}"
        if el_id: desc += f" id='{el_id}'"
        if name: desc += f" name='{name}'"
        if el_type: desc += f" type='{el_type}'"
        if placeholder: desc += f" placeholder='{placeholder}'"
        desc += f"> {text}"
        
        print(f"[{i+1}] {desc}")

        # Smart detection rules for common fields
        if el_type in ["text", "email"] or tag == "input":
            if any(k in (el_id or "") + (name or "") + (placeholder or "") for k in ["user", "login", "email", "id"]):
                if "username_field" not in detected:
                    detected["username_field"] = f"#{el_id}" if el_id else f"input[name='{name}']"
            if any(k in (el_id or "") + (name or "") + (placeholder or "") for k in ["pass", "pwd", "secret"]):
                if "password_field" not in detected:
                    detected["password_field"] = f"#{el_id}" if el_id else f"input[type='password']"
            if any(k in (el_id or "") + (name or "") + (placeholder or "") for k in ["search", "quick", "ftir", "find"]):
                if "quick_search_box" not in detected:
                    detected["quick_search_box"] = f"#{el_id}" if el_id else f"input[name='{name}']"

        if tag == "textarea":
            if "reply_textarea" not in detected:
                detected["reply_textarea"] = f"#{el_id}" if el_id else "textarea"

        if el_type == "submit" or tag == "button":
            if any(k in (text + (el_id or "") + (name or "")).lower() for k in ["login", "sign in", "submit"]):
                if "login_button" not in detected:
                    detected["login_button"] = f"#{el_id}" if el_id else f"button[type='submit']"

    print("\n--- Auto-detected Selectors ---")
    for k, v in detected.items():
        print(f"  {k}: {v}")

    print("\nKeep browser open? Press Enter in terminal when done...")
    input()
    driver.quit()

if __name__ == "__main__":
    detect_selectors()
