#!/usr/bin/env python3
"""
main_msforms.py

Submit rows from an Excel sheet to a Microsoft Form using Playwright.

- Downloads Excel if excel_path is an HTTP(S) URL (supports optional Bearer token via EXCEL_BEARER_TOKEN env or config).
- Supports single_row (1-based) to target exactly one row.
- When running in live mode the script launches Playwright, fills the form and submits.

WARNING: If the form or the Excel file require authentication, provide the appropriate secrets
(EXCEL_BEARER_TOKEN for SharePoint downloads). This script does not perform Microsoft sign-in;
it fills the public form page. If the form requires sign-in, automation must be extended.
"""
from pathlib import Path
import tempfile
import time
import logging
import csv
import os
import requests
import pandas as pd
import yaml
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

# Load config
cfg_path = Path("config.yml")
if not cfg_path.exists():
    logging.error("Missing config.yml in repo root.")
    raise SystemExit(1)
config = yaml.safe_load(cfg_path.read_text())

EXCEL_PATH = config.get("excel_path", "data/input.xlsx")
SHEET_NAME = config.get("sheet_name", 0)
START_ROW = int(config.get("start_row", 0))
SINGLE_ROW = config.get("single_row", None)  # 1-based row number if provided
DELAY = float(config.get("delay_seconds", 2))
HEADLESS = bool(config.get("headless", True))
SIMULATE = bool(config.get("simulate", False))

FORM_URL = os.environ.get("FORM_URL") or config.get("form_url")
if not FORM_URL:
    logging.error("FORM_URL not set in env or config.yml.")
    raise SystemExit(1)

# Mapping: "Excel column header" -> "Visible question label text in form"
MAPPING = config.get("mapping", {})

# Optional auth for downloading Excel from SharePoint / protected host
EXCEL_BEARER_TOKEN = os.environ.get("EXCEL_BEARER_TOKEN") or config.get("excel_bearer_token")


def download_file(url: str, dest: Path, token: Optional[str] = None) -> Path:
    logging.info(f"Downloading Excel from {url} to {dest}")
    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    with requests.get(url, headers=headers, stream=True, allow_redirects=True, timeout=60) as r:
        r.raise_for_status()
        with dest.open("wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    logging.info("Download complete.")
    return dest


def prepare_excel_path(excel_path: str) -> Path:
    if excel_path.startswith("http://") or excel_path.startswith("https://"):
        tmp = Path(tempfile.gettempdir()) / f"downloaded_excel_{int(time.time())}.xlsx"
        try:
            download_file(excel_path, tmp, token=EXCEL_BEARER_TOKEN)
        except Exception as exc:
            logging.error(f"Failed to download Excel from URL: {exc}")
            raise
        return tmp
    else:
        p = Path(excel_path)
        if not p.exists():
            logging.error(f"Local Excel path not found: {excel_path}")
            raise SystemExit(1)
        return p


def get_value(row, col):
    val = row.get(col, "")
    if pd.isna(val):
        return ""
    return str(val)


def simulate_row_actions(row_idx, row):
    print(f"--- SIMULATION: Row index (0-based) = {row_idx} ---")
    for excel_col, q_label in MAPPING.items():
        value = get_value(row, excel_col)
        if value == "":
            print(f"  [skip] Column '{excel_col}' -> question '{q_label}' (empty)")
            continue
        if len(value) < 50 and ("\n" not in value) and (value.lower() in ("yes", "no") or len(value.split()) <= 3):
            print(f"  [would choose option] question='{q_label}' -> option='{value}'")
        else:
            print(f"  [would fill text] question='{q_label}' -> '{value}'")
    print("  [would click] Submit button")
    print("--- END SIMULATION ---\n")


def find_and_fill_input(page, label: str, value: str) -> bool:
    # Try get_by_label
    try:
        fld = page.get_by_label(label, exact=False)
        if fld.count() > 0:
            fld.first.fill(value)
            return True
    except Exception:
        pass

    # Try aria-label/placeholder selectors
    try:
        sel = f"input[aria-label*='{label}'], textarea[aria-label*='{label}'], input[placeholder*='{label}'], textarea[placeholder*='{label}']"
        fld = page.locator(sel)
        if fld.count() > 0:
            fld.first.fill(value)
            return True
    except Exception:
        pass

    # Find a question block containing the label and fill an input inside it
    try:
        esc = label.replace('"', '\\"')
        xpath = "//div[contains(normalize-space(.), \"{}\")]" .format(esc)
        q = page.locator("xpath=" + xpath).first
        if q and q.count() > 0:
            fld = q.locator("input, textarea, [role='textbox']")
            if fld.count() > 0:
                fld.first.fill(value)
                return True
    except Exception:
        pass

    return False


def find_and_choose_option(page, label: str, value: str) -> bool:
    try:
        esc = label.replace('"', '\\"')
        xpath = "//div[contains(normalize-space(.), \"{}\")]" .format(esc)
        q = page.locator("xpath=" + xpath).first
        if not q or q.count() == 0:
            return False

        # Exact text match first
        opt = q.locator(f"text=\"{value}\"")
        if opt.count() > 0:
            opt.first.click()
            return True

        # Partial match inside question block
        opt = q.locator(f"xpath=.//*[contains(normalize-space(.), \"{value}\")]")
        if opt.count() > 0:
            opt.first.click()
            return True

        # Try radio/checkbox roles
        try:
            rb = q.get_by_role("radio", name=value)
            if rb.count() > 0:
                rb.first.click()
                return True
        except Exception:
            pass

        try:
            cb = q.get_by_role("checkbox", name=value)
            if cb.count() > 0:
                cb.first.click()
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def submit_one_row_live(page, row_idx, row):
    result = {"index": int(row_idx)}
    try:
        for excel_col, q_label in MAPPING.items():
            value = get_value(row, excel_col)
            if value == "":
                continue

            if find_and_fill_input(page, q_label, value):
                continue

            if find_and_choose_option(page, q_label, value):
                continue

            logging.warning(f"Row {row_idx}: could not find field for label '{q_label}' (col '{excel_col}')")
            result["status"] = f"missing_field:{q_label}"
            return result

        # Click submit (multiple fallbacks)
        submit_selectors = [
            "button:has-text('Submit')",
            "button:has-text('Send')",
            "button:has-text('Done')",
            "input[type='submit']",
            "button[aria-label='Submit']",
            "button[role='button']"
        ]
        clicked = False
        for sel in submit_selectors:
            try:
                btn = page.locator(sel).first
                if btn and btn.count() > 0:
                    btn.click()
                    clicked = True
                    break
            except PWTimeoutError:
                continue
            except Exception:
                continue

        if not clicked:
            try:
                page.keyboard.press("Enter")
            except Exception:
                logging.warning("Could not find a submit button and Enter failed.")

        time.sleep(3)
        result["status"] = "submitted"
    except Exception as exc:
        logging.exception(f"Exception submitting row {row_idx}: {exc}")
        result["status"] = "error"
        result["error"] = str(exc)
    return result


def main():
    local_excel = prepare_excel_path(EXCEL_PATH)

    try:
        df = pd.read_excel(local_excel, sheet_name=SHEET_NAME, engine="openpyxl")
    except Exception as exc:
        logging.error(f"Failed to read sheet '{SHEET_NAME}' from Excel: {exc}")
        raise SystemExit(1)

    df = df.fillna("")

    if SINGLE_ROW is not None:
        try:
            single_idx = int(SINGLE_ROW) - 1
            if single_idx < 0 or single_idx >= len(df):
                logging.error(f"single_row {SINGLE_ROW} out of range (sheet has {len(df)} rows)")
                raise SystemExit(1)
            df = df.iloc[[single_idx]].reset_index(drop=True)
            start_offset = single_idx
        except Exception as exc:
            logging.error(f"Invalid single_row value: {exc}")
            raise SystemExit(1)
    else:
        if START_ROW > 0:
            df = df.iloc[START_ROW:].reset_index(drop=True)
            start_offset = START_ROW
        else:
            start_offset = 0

    results = []

    if SIMULATE:
        logging.info("Running in SIMULATE mode — no browser will be launched.")
        for idx, row in df.iterrows():
            simulate_row_actions(idx + start_offset, row)
            results.append({"index": int(idx + start_offset), "status": "simulated"})
        out_path = Path("processed.csv")
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["index", "status"])
            writer.writeheader()
            writer.writerows(results)
        logging.info(f"Simulation complete. Results written to {out_path}")
        return

    # Live mode: launch browser and submit
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        for idx, row in df.iterrows():
            logging.info(f"Processing row {idx + start_offset}")
            try:
                page.goto(FORM_URL, wait_until="networkidle", timeout=30000)
            except Exception:
                logging.warning("Navigation timed out; retrying without timeout")
                page.goto(FORM_URL)
            time.sleep(1.5)
            res = submit_one_row_live(page, idx + start_offset, row)
            results.append(res)
            time.sleep(DELAY)
        browser.close()

    out_path = Path("processed.csv")
    keys = sorted({k for r in results for k in r.keys()}) if results else ["index", "status"]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(results)
    logging.info(f"Done. {len(results)} rows processed. Results written to {out_path}")


if __name__ == "__main__":
    main()