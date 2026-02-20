#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.
"""

import csv
import logging
import os
import re
import time
from urllib.parse import urljoin

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_URL = "https://members.glass.org/cvweb/cgi-bin/"
SEARCH_URL = BASE_URL + "utilities.dll/OpenPage?wrp=ngaSearch.htm"
MEMBER_URL_TEMPLATE = BASE_URL + "organizationdll.dll/Info?orgcd={orgcd}&wrp=organizationinfo.htm"

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
CSV_FILE = os.path.join(OUTPUT_DIR, "nga_members.csv")
XLSX_FILE = os.path.join(OUTPUT_DIR, "nga_members.xlsx")

FIELDS = ["Name", "Classification", "Address", "City", "State", "Phone", "Web Address"]

REQUEST_DELAY = 1.5          # seconds between detail-page requests
PAGE_LOAD_TIMEOUT = 30_000   # ms
RETRY_COUNT = 3
RETRY_BACKOFF = 2            # seconds, doubles each retry

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def retry(func, retries=RETRY_COUNT, backoff=RETRY_BACKOFF):
    """Call *func* with exponential-backoff retries."""
    for attempt in range(1, retries + 1):
        try:
            return func()
        except Exception as exc:
            if attempt == retries:
                raise
            wait = backoff * (2 ** (attempt - 1))
            log.warning("Attempt %d failed (%s). Retrying in %ss …", attempt, exc, wait)
            time.sleep(wait)


def safe_text(locator):
    """Return trimmed inner text of a locator, or empty string on failure."""
    try:
        txt = locator.inner_text(timeout=3000)
        return txt.strip() if txt else ""
    except Exception:
        return ""


def extract_field_by_label(page, label):
    """Try common cvweb patterns to pull a labelled field value."""
    # Pattern 1: <td>Label:</td><td>Value</td>  (table-based layouts)
    try:
        cell = page.locator(f"td:has-text('{label}')").first
        if cell.count():
            sibling = cell.locator("xpath=following-sibling::td[1]")
            val = safe_text(sibling)
            if val:
                return val
    except Exception:
        pass

    # Pattern 2: <span/label class="...">Label</span> … <span class="...">Value</span>
    try:
        lbl = page.locator(f"text='{label}'").first
        if lbl.count():
            parent = lbl.locator("xpath=..")
            # grab text after the label
            full = safe_text(parent)
            if ":" in full:
                return full.split(":", 1)[1].strip()
    except Exception:
        pass

    return ""


# ---------------------------------------------------------------------------
# Scraping logic
# ---------------------------------------------------------------------------
def collect_member_links(page):
    """Navigate search results and return a list of (orgcd, name) tuples."""
    members = []

    log.info("Navigating to search page: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(2000)

    # Log the page title and a snippet of the page for debugging
    title = page.title()
    log.info("Page title: %s", title)

    # ----- Try to submit the search form with empty criteria -----
    # Look for common submit button patterns
    submit_selectors = [
        "input[type='submit']",
        "button[type='submit']",
        "input[value='Search']",
        "input[value='search']",
        "input[value='Find']",
        "button:has-text('Search')",
        "button:has-text('Find')",
        "a:has-text('Search')",
        "#btnSearch",
        ".btn-search",
        "input[name='submit']",
    ]

    submitted = False
    for sel in submit_selectors:
        try:
            btn = page.locator(sel).first
            if btn.count():
                log.info("Found submit element: %s", sel)
                btn.click(timeout=5000)
                page.wait_for_timeout(3000)
                submitted = True
                break
        except Exception:
            continue

    if not submitted:
        # Fallback: try pressing Enter in the first text input
        try:
            first_input = page.locator("input[type='text']").first
            if first_input.count():
                first_input.press("Enter")
                page.wait_for_timeout(3000)
                submitted = True
                log.info("Pressed Enter on first text input as fallback.")
        except Exception:
            pass

    if not submitted:
        log.warning("Could not find a submit button. Will try to parse current page for links anyway.")

    # ----- Collect member links from result pages -----
    page_num = 0
    while True:
        page_num += 1
        log.info("Parsing results page %d …", page_num)

        # Look for organization detail links
        # Pattern: organizationdll.dll/Info?orgcd=XXXXXX
        links = page.locator("a[href*='organizationdll.dll/Info']").all()
        if not links:
            # Also try case-insensitive via evaluate
            links = page.locator("a[href*='organizationdll']").all()

        new_count = 0
        for link in links:
            try:
                href = link.get_attribute("href") or ""
                name = safe_text(link)
                match = re.search(r"orgcd=(\d+)", href, re.IGNORECASE)
                if match:
                    orgcd = match.group(1)
                    if orgcd not in {m[0] for m in members}:
                        members.append((orgcd, name))
                        new_count += 1
            except Exception:
                continue

        log.info("  Found %d new member links on page %d (total: %d)", new_count, page_num, len(members))

        if new_count == 0 and page_num > 1:
            break

        # ----- Try pagination -----
        navigated = False

        # Pattern A: "Next" link / button
        next_selectors = [
            "a:has-text('Next')",
            "a:has-text('next')",
            "a:has-text('>>')",
            "a:has-text('>')",
            "a.next",
            "li.next > a",
            "a[title='Next']",
            "input[value='Next']",
        ]
        for sel in next_selectors:
            try:
                nxt = page.locator(sel).first
                if nxt.count() and nxt.is_visible():
                    nxt.click(timeout=5000)
                    page.wait_for_timeout(2000)
                    navigated = True
                    log.info("  Navigated via '%s'", sel)
                    break
            except Exception:
                continue

        # Pattern B: cvweb Range-based pagination in URL
        if not navigated:
            current_url = page.url
            range_match = re.search(r"Range=(\d+)/(\d+)", current_url, re.IGNORECASE)
            if range_match:
                start = int(range_match.group(1))
                size = int(range_match.group(2))
                next_start = start + size
                next_url = re.sub(r"Range=\d+/\d+", f"Range={next_start}/{size}", current_url)
                log.info("  Trying Range pagination: %s", next_url)
                try:
                    page.goto(next_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)
                    page.wait_for_timeout(2000)
                    navigated = True
                except Exception:
                    pass

        if not navigated:
            log.info("No more pages to navigate.")
            break

    log.info("Total unique member links collected: %d", len(members))
    return members


def scrape_member_detail(page, orgcd, list_name=""):
    """Visit a member detail page and extract fields. Returns a dict."""
    url = MEMBER_URL_TEMPLATE.format(orgcd=orgcd)
    record = {f: "" for f in FIELDS}

    def load():
        page.goto(url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)

    try:
        retry(load)
    except Exception as exc:
        log.error("Failed to load member page orgcd=%s: %s", orgcd, exc)
        record["Name"] = list_name
        return record

    page.wait_for_timeout(1000)

    # --- Name ---
    # Try <h1>, <h2>, or specific selectors
    for sel in ["h1", "h2", ".org-name", "#orgName", "td.orgname"]:
        txt = safe_text(page.locator(sel).first)
        if txt and len(txt) < 200:
            record["Name"] = txt
            break
    if not record["Name"]:
        record["Name"] = list_name

    # --- Structured extraction: scan the page body text for labelled fields ---
    body_text = safe_text(page.locator("body"))

    # Helper to pull value after a label in the body text
    def pull(label, body=body_text):
        pattern = rf"{re.escape(label)}\s*[:\-]?\s*(.+)"
        m = re.search(pattern, body, re.IGNORECASE)
        if m:
            return m.group(1).split("\n")[0].strip()
        return ""

    # Classification (business type)
    if not record["Classification"]:
        record["Classification"] = extract_field_by_label(page, "Classification")
    if not record["Classification"]:
        record["Classification"] = extract_field_by_label(page, "Business Type")
    if not record["Classification"]:
        record["Classification"] = extract_field_by_label(page, "Type")
    if not record["Classification"]:
        record["Classification"] = pull("Classification")
    if not record["Classification"]:
        record["Classification"] = pull("Business Type")

    # Address
    if not record["Address"]:
        record["Address"] = extract_field_by_label(page, "Address")
    if not record["Address"]:
        record["Address"] = pull("Address")

    # City
    if not record["City"]:
        record["City"] = extract_field_by_label(page, "City")
    if not record["City"]:
        record["City"] = pull("City")

    # State
    if not record["State"]:
        record["State"] = extract_field_by_label(page, "State")
    if not record["State"]:
        record["State"] = pull("State")

    # Phone
    if not record["Phone"]:
        record["Phone"] = extract_field_by_label(page, "Phone")
    if not record["Phone"]:
        record["Phone"] = pull("Phone")
    # Fallback: find phone-number pattern in body
    if not record["Phone"]:
        phone_match = re.search(r"\(?\d{3}\)?[\s\-.]?\d{3}[\s\-.]?\d{4}", body_text)
        if phone_match:
            record["Phone"] = phone_match.group(0).strip()

    # Web Address
    if not record["Web Address"]:
        record["Web Address"] = extract_field_by_label(page, "Web Address")
    if not record["Web Address"]:
        record["Web Address"] = extract_field_by_label(page, "Website")
    if not record["Web Address"]:
        record["Web Address"] = extract_field_by_label(page, "URL")
    if not record["Web Address"]:
        record["Web Address"] = pull("Web Address")
    if not record["Web Address"]:
        record["Web Address"] = pull("Website")
    # Fallback: find a link that looks like a member website
    if not record["Web Address"]:
        ext_links = page.locator("a[href^='http']").all()
        for lnk in ext_links:
            try:
                href = lnk.get_attribute("href") or ""
                if "glass.org" not in href and "mailto:" not in href and href.startswith("http"):
                    record["Web Address"] = href
                    break
            except Exception:
                continue

    # --- Try to parse address block if city/state still empty ---
    # cvweb often puts address in a single block; try to split
    if record["Address"] and not record["City"]:
        # Look for "City, ST ZIP" pattern after address
        addr_pattern = re.search(
            r"([A-Za-z\s]+),\s*([A-Z]{2})\s+(\d{5})",
            body_text[body_text.find(record["Address"]):]
        )
        if addr_pattern:
            record["City"] = addr_pattern.group(1).strip()
            record["State"] = addr_pattern.group(2).strip()

    log.info("  Scraped: %s | %s | %s, %s",
             record["Name"], record["Classification"], record["City"], record["State"])
    return record


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def export_csv(records):
    """Write records to CSV."""
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(records)
    log.info("CSV saved to %s (%d rows)", CSV_FILE, len(records))


def export_xlsx(records):
    """Write records to XLSX."""
    wb = Workbook()
    ws = wb.active
    ws.title = "NGA Members"
    ws.append(FIELDS)
    for rec in records:
        ws.append([rec.get(f, "") for f in FIELDS])
    # Auto-width columns
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
    wb.save(XLSX_FILE)
    log.info("XLSX saved to %s (%d rows)", XLSX_FILE, len(records))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )
        page = context.new_page()

        # Phase 1 – collect all member links from search results
        member_links = collect_member_links(page)

        if not member_links:
            log.error("No member links found. Saving debug screenshot.")
            page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_screenshot.png"), full_page=True)
            # Save page HTML for debugging
            with open(os.path.join(OUTPUT_DIR, "debug_page.html"), "w", encoding="utf-8") as f:
                f.write(page.content())
            browser.close()
            return

        # Phase 2 – visit each member detail page
        records = []
        total = len(member_links)
        failed = 0
        for idx, (orgcd, name) in enumerate(member_links, 1):
            log.info("[%d/%d] Scraping orgcd=%s (%s) …", idx, total, orgcd, name)
            try:
                rec = scrape_member_detail(page, orgcd, list_name=name)
                records.append(rec)
            except Exception as exc:
                log.error("  FAILED orgcd=%s: %s", orgcd, exc)
                failed += 1
                records.append({f: "" for f in FIELDS})
                records[-1]["Name"] = name

            if idx < total:
                time.sleep(REQUEST_DELAY)

        browser.close()

    # Phase 3 – export
    export_csv(records)
    export_xlsx(records)

    log.info("=" * 60)
    log.info("DONE — Total: %d | Scraped: %d | Failed: %d", total, total - failed, failed)
    log.info("=" * 60)


if __name__ == "__main__":
    main()
