#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Strategy:
  Phase 1 — Parse the results table across all paginated pages to collect
            Name, Classification, Address, City, State, Phone, and detail links.
  Phase 2 — Visit each member detail page to extract Web Address.
  Phase 3 — Export to CSV and XLSX.
"""

import csv
import logging
import os
import re
import time

from playwright.sync_api import sync_playwright
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_URL = "https://members.glass.org/cvweb/cgi-bin/"
SEARCH_URL = BASE_URL + "utilities.dll/OpenPage?wrp=ngaSearch.htm"

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
CSV_FILE = os.path.join(OUTPUT_DIR, "nga_members.csv")
XLSX_FILE = os.path.join(OUTPUT_DIR, "nga_members.xlsx")

FIELDS = ["Name", "Classification", "Address", "City", "State", "Phone", "Web Address"]

DETAIL_DELAY = 0.5           # seconds between detail-page requests
PAGE_LOAD_TIMEOUT = 45_000   # ms
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


def split_city_state(city_state_text):
    """Split 'City, ST' or 'City, ST 12345' into (city, state)."""
    text = city_state_text.strip()
    if not text:
        return "", ""
    # Pattern: "City, ST" or "City, ST 12345"
    m = re.match(r"^(.+?),\s*([A-Z]{2})\b", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # Fallback: just return the whole thing as city
    return text, ""


# ---------------------------------------------------------------------------
# Phase 1: Parse results table across all pages
# ---------------------------------------------------------------------------
def parse_results_table(page):
    """
    Parse the results table on the current page.
    Returns a list of dicts with keys from FIELDS (Web Address empty)
    and an extra '_detail_url' key for the detail page link.
    """
    records = []

    # Find all data rows in the results table.
    # The table has header row(s) and data rows. We look for rows inside
    # <table> elements that contain <td> cells (not <th>).
    rows = page.locator("table tr").all()

    for row in rows:
        cells = row.locator("td").all()
        # The results table from the screenshot has 6 columns:
        # Name | Classification | Address | City/State | Phone | Map
        # Skip rows that don't have enough cells (headers, spacers, etc.)
        if len(cells) < 5:
            continue

        try:
            # Column 0: Name (with a link to detail page)
            name_cell = cells[0]
            name = name_cell.inner_text(timeout=3000).strip()

            # Skip header rows or empty rows
            if not name or name.lower() == "name" or name.lower() == "organization":
                continue

            # Try to get the detail page link from the name cell
            detail_url = ""
            link = name_cell.locator("a").first
            if link.count():
                href = link.get_attribute("href") or ""
                if href:
                    # Make absolute URL if relative
                    if href.startswith("http"):
                        detail_url = href
                    elif href.startswith("/"):
                        detail_url = "https://members.glass.org" + href
                    else:
                        detail_url = BASE_URL + href

            # Column 1: Classification (business type)
            classification = cells[1].inner_text(timeout=3000).strip() if len(cells) > 1 else ""

            # Column 2: Address
            address = cells[2].inner_text(timeout=3000).strip() if len(cells) > 2 else ""

            # Column 3: City/State (combined, needs splitting)
            city_state_raw = cells[3].inner_text(timeout=3000).strip() if len(cells) > 3 else ""
            city, state = split_city_state(city_state_raw)

            # Column 4: Phone
            phone = cells[4].inner_text(timeout=3000).strip() if len(cells) > 4 else ""

            record = {
                "Name": name,
                "Classification": classification,
                "Address": address,
                "City": city,
                "State": state,
                "Phone": phone,
                "Web Address": "",
                "_detail_url": detail_url,
            }
            records.append(record)

        except Exception as exc:
            log.warning("Failed to parse a row: %s", exc)
            continue

    return records


def collect_all_members(page):
    """
    Navigate to search, submit form, and paginate through all result pages.
    Returns a list of member dicts.
    """
    all_members = []

    log.info("Navigating to search page: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    title = page.title()
    log.info("Page title: %s", title)

    # ----- Submit the search form -----
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
        "input[type='image']",
    ]

    submitted = False
    for sel in submit_selectors:
        try:
            btn = page.locator(sel).first
            if btn.count() and btn.is_visible():
                log.info("Clicking submit element: %s", sel)
                btn.click(timeout=5000)
                page.wait_for_timeout(4000)
                submitted = True
                break
        except Exception:
            continue

    if not submitted:
        # Try pressing Enter on first text input
        try:
            first_input = page.locator("input[type='text']").first
            if first_input.count():
                first_input.press("Enter")
                page.wait_for_timeout(4000)
                submitted = True
                log.info("Pressed Enter on first text input.")
        except Exception:
            pass

    if not submitted:
        log.warning("Could not find submit button. Parsing current page anyway.")

    # Take a screenshot of the results page for debugging
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"), full_page=True)

    # ----- Paginate through all result pages -----
    page_num = 0
    max_pages = 200  # safety limit

    while page_num < max_pages:
        page_num += 1
        log.info("Parsing results page %d …", page_num)

        new_records = parse_results_table(page)
        if new_records:
            all_members.extend(new_records)
            log.info("  Found %d members on page %d (total: %d)",
                     len(new_records), page_num, len(all_members))
        else:
            log.info("  No data rows found on page %d.", page_num)
            # If this is the first page and we found nothing, save debug info
            if page_num == 1:
                page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_screenshot.png"), full_page=True)
                with open(os.path.join(OUTPUT_DIR, "debug_page.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())
                log.error("No data found on first page. Debug files saved.")
            break

        # ----- Navigate to next page -----
        navigated = False

        # Try clicking the → (right arrow) link for next page
        # The screenshot shows pagination like: 1 2 3 ... 83 84 85 →
        next_selectors = [
            "a:has-text('→')",
            "a:has-text('›')",
            "a:has-text('»')",
            "a:has-text('Next')",
            "a:has-text('next')",
            "a:has-text('>>')",
        ]
        for sel in next_selectors:
            try:
                nxt = page.locator(sel).first
                if nxt.count() and nxt.is_visible():
                    nxt.click(timeout=10000)
                    page.wait_for_timeout(3000)
                    navigated = True
                    log.info("  Navigated to next page via '%s'", sel)
                    break
            except Exception:
                continue

        # Fallback: try URL Range parameter manipulation
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
                except Exception as exc:
                    log.warning("  Range pagination failed: %s", exc)

        if not navigated:
            log.info("No more pages to navigate. Stopping pagination.")
            break

    log.info("Total members collected from tables: %d", len(all_members))
    return all_members


# ---------------------------------------------------------------------------
# Phase 2: Visit detail pages for Web Address
# ---------------------------------------------------------------------------
def scrape_web_address(page, detail_url):
    """Visit a member detail page and extract the Web Address."""
    if not detail_url:
        return ""

    try:
        retry(lambda: page.goto(detail_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    except Exception as exc:
        log.warning("  Could not load detail page %s: %s", detail_url, exc)
        return ""

    page.wait_for_timeout(500)

    # Strategy 1: look for a labelled "Web Address" or "Website" field
    body_text = ""
    try:
        body_text = page.locator("body").inner_text(timeout=5000)
    except Exception:
        pass

    # Try to find "Web Address:" followed by a URL in the text
    for label in ["Web Address", "Website", "Web Site", "URL", "Home Page"]:
        pattern = rf"{re.escape(label)}\s*[:\-]?\s*(https?://\S+)"
        m = re.search(pattern, body_text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Strategy 2: look for td label / td value pattern
    for label in ["Web Address", "Website", "Web Site", "URL"]:
        try:
            cells = page.locator(f"td:has-text('{label}')").all()
            for cell in cells:
                sibling = cell.locator("xpath=following-sibling::td[1]")
                if sibling.count():
                    # Check for a link inside the sibling
                    link = sibling.locator("a").first
                    if link.count():
                        href = link.get_attribute("href") or ""
                        if href and href.startswith("http"):
                            return href
                    # Or just get text
                    txt = sibling.inner_text(timeout=2000).strip()
                    if txt and ("." in txt) and (" " not in txt):
                        if not txt.startswith("http"):
                            txt = "http://" + txt
                        return txt
        except Exception:
            continue

    # Strategy 3: find any external link that doesn't point to glass.org
    try:
        all_links = page.locator("a[href^='http']").all()
        for lnk in all_links:
            href = lnk.get_attribute("href") or ""
            if href and "glass.org" not in href and "mailto:" not in href:
                # Check it looks like a real website, not a social/map link
                if any(skip in href.lower() for skip in ["google.com/maps", "facebook.com", "twitter.com", "linkedin.com"]):
                    continue
                return href
    except Exception:
        pass

    return ""


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def export_csv(records):
    """Write records to CSV."""
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
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

        # Phase 1 – collect members from results table
        members = collect_all_members(page)

        if not members:
            log.error("No members found. Check debug files in output/.")
            browser.close()
            return

        # Phase 2 – visit detail pages for Web Address
        total = len(members)
        log.info("=" * 60)
        log.info("Phase 2: Fetching Web Address from %d detail pages …", total)
        log.info("=" * 60)
        web_found = 0
        web_failed = 0
        for idx, member in enumerate(members, 1):
            detail_url = member.get("_detail_url", "")
            if not detail_url:
                web_failed += 1
                continue

            log.info("[%d/%d] %s", idx, total, member["Name"])
            try:
                web = scrape_web_address(page, detail_url)
                member["Web Address"] = web
                if web:
                    web_found += 1
                    log.info("  Web: %s", web)
                else:
                    log.info("  No web address found.")
            except Exception as exc:
                log.warning("  Error fetching web address: %s", exc)
                web_failed += 1

            if idx < total:
                time.sleep(DETAIL_DELAY)

        browser.close()

    # Phase 3 – export
    export_csv(members)
    export_xlsx(members)

    log.info("=" * 60)
    log.info("DONE — Members: %d | Web addresses found: %d | Failed: %d",
             total, web_found, web_failed)
    log.info("=" * 60)


if __name__ == "__main__":
    main()
