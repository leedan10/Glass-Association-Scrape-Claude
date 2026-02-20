#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Strategy:
  Phase 1 — Use JavaScript DOM evaluation to extract table data from all
            paginated result pages (Name, Classification, Address, City,
            State, Phone, detail URL).
  Phase 2 — Visit each member detail page to extract Web Address.
  Phase 3 — Export to CSV and XLSX.
"""

import csv
import json
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
    for attempt in range(1, retries + 1):
        try:
            return func()
        except Exception as exc:
            if attempt == retries:
                raise
            wait = backoff * (2 ** (attempt - 1))
            log.warning("Attempt %d failed (%s). Retrying in %ss …", attempt, exc, wait)
            time.sleep(wait)


# JavaScript that runs inside the browser to extract table rows.
# This handles nested tables, mixed th/td, and any cvweb layout quirks.
JS_EXTRACT_TABLE = """
() => {
    // Find all tables on the page
    const tables = document.querySelectorAll('table');
    let dataRows = [];

    for (const table of tables) {
        const rows = table.querySelectorAll('tr');
        for (const row of rows) {
            // Get all cells (td or th)
            const cells = Array.from(row.querySelectorAll('td, th'));
            if (cells.length < 5) continue;

            // Read text from each cell
            const texts = cells.map(c => (c.innerText || '').trim());

            // Skip header rows: look for rows where first cell is "Name" or similar
            const first = texts[0].toLowerCase();
            if (first === 'name' || first === 'organization' || first === '' ||
                first === 'company' || first.includes('▼') || first.includes('▲'))
                continue;

            // Skip rows that look like pagination or footer
            if (texts[0].match(/^\\d+$/) || texts[0].includes('Page') ||
                texts[0].includes('Match') || texts[0].includes('MATCH'))
                continue;

            // Try to get a link from the first cell (member name)
            let detailUrl = '';
            const link = cells[0].querySelector('a[href]');
            if (link) {
                detailUrl = link.href || '';
            }

            // Build the record: Name, Classification, Address, CityState, Phone
            dataRows.push({
                name: texts[0],
                classification: texts[1] || '',
                address: texts[2] || '',
                cityState: texts[3] || '',
                phone: texts[4] || '',
                detailUrl: detailUrl
            });
        }
    }

    // Deduplicate by name+address (in case layout tables cause duplicates)
    const seen = new Set();
    const unique = [];
    for (const row of dataRows) {
        const key = row.name + '|' + row.address;
        if (!seen.has(key) && row.name.length > 0) {
            seen.add(key);
            unique.push(row);
        }
    }

    return unique;
}
"""

# JavaScript to find and click the next-page arrow
JS_CLICK_NEXT = """
() => {
    // Look for pagination links: →, ›, Next, >>
    const allLinks = document.querySelectorAll('a');
    for (const a of allLinks) {
        const text = (a.innerText || '').trim();
        if (text === '→' || text === '›' || text === '»' ||
            text === '>>' || text.toLowerCase() === 'next') {
            a.click();
            return true;
        }
    }
    // Also check for input buttons
    const buttons = document.querySelectorAll('input[type="button"], input[type="submit"]');
    for (const btn of buttons) {
        const val = (btn.value || '').trim().toLowerCase();
        if (val === 'next' || val === '>' || val === '>>') {
            btn.click();
            return true;
        }
    }
    return false;
}
"""

# JavaScript to extract the web address from a detail page
JS_EXTRACT_WEB_ADDRESS = """
() => {
    const body = document.body.innerText || '';

    // Strategy 1: Look for "Web Address" or "Website" label followed by a URL
    const labels = ['Web Address', 'Website', 'Web Site', 'URL', 'Home Page', 'Web'];
    for (const label of labels) {
        // Find in table cells: <td>Label</td><td>value</td>
        const allCells = document.querySelectorAll('td, th');
        for (let i = 0; i < allCells.length; i++) {
            const cellText = (allCells[i].innerText || '').trim();
            if (cellText.toLowerCase().replace(/[:\\s]/g, '') === label.toLowerCase().replace(/[:\\s]/g, '')) {
                // Check the next sibling cell
                const next = allCells[i + 1] || allCells[i].nextElementSibling;
                if (next) {
                    // First check for a link
                    const link = next.querySelector('a[href]');
                    if (link && link.href && !link.href.includes('glass.org') && !link.href.includes('mailto:')) {
                        return link.href;
                    }
                    // Then check text content
                    const txt = (next.innerText || '').trim();
                    if (txt && txt.includes('.') && !txt.includes(' ') && txt.length > 3) {
                        return txt.startsWith('http') ? txt : 'http://' + txt;
                    }
                }
            }
        }
    }

    // Strategy 2: Find "Web Address" in body text followed by URL-like text
    for (const label of labels) {
        const regex = new RegExp(label + '[:\\\\s]*([\\\\S]+\\\\.[\\\\S]+)', 'i');
        const match = body.match(regex);
        if (match && match[1] && match[1].includes('.') && !match[1].includes('glass.org')) {
            const url = match[1].replace(/[,;)]+$/, '');
            return url.startsWith('http') ? url : 'http://' + url;
        }
    }

    // Strategy 3: Find external links not pointing to glass.org
    const skipDomains = ['glass.org', 'google.com', 'facebook.com', 'twitter.com',
                         'linkedin.com', 'youtube.com', 'instagram.com', 'bing.com',
                         'javascript:', 'mailto:'];
    const links = document.querySelectorAll('a[href]');
    for (const link of links) {
        const href = link.href || '';
        if (!href.startsWith('http')) continue;
        let dominated = false;
        for (const skip of skipDomains) {
            if (href.includes(skip)) { dominated = true; break; }
        }
        if (!dominated) return href;
    }

    return '';
}
"""


def split_city_state(city_state_text):
    """Split 'City, ST' or 'City, ST 12345' into (city, state)."""
    text = city_state_text.strip()
    if not text:
        return "", ""
    m = re.match(r"^(.+?),\s*([A-Z]{2})\b", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return text, ""


# ---------------------------------------------------------------------------
# Phase 1: Collect all members from paginated results
# ---------------------------------------------------------------------------
def collect_all_members(page):
    all_members = []

    log.info("Navigating to search page: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    title = page.title()
    log.info("Page title: %s", title)

    # Submit the search form
    submit_selectors = [
        "input[type='submit']",
        "button[type='submit']",
        "input[value='Search']",
        "input[value='search']",
        "input[value='Find']",
        "button:has-text('Search')",
        "button:has-text('Find')",
        "a:has-text('Search')",
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

    # Save first-page screenshot for debugging
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"), full_page=True)

    # Also dump raw HTML for debugging
    with open(os.path.join(OUTPUT_DIR, "results_page1.html"), "w", encoding="utf-8") as f:
        f.write(page.content())
    log.info("Saved results_page1.html for debugging.")

    # Paginate through all result pages
    page_num = 0
    max_pages = 200

    while page_num < max_pages:
        page_num += 1
        log.info("Parsing results page %d …", page_num)

        # Use JavaScript to extract table data
        raw_rows = page.evaluate(JS_EXTRACT_TABLE)
        log.info("  JS extraction returned %d rows", len(raw_rows))

        if raw_rows:
            for row in raw_rows:
                city, state = split_city_state(row.get("cityState", ""))
                record = {
                    "Name": row.get("name", ""),
                    "Classification": row.get("classification", ""),
                    "Address": row.get("address", ""),
                    "City": city,
                    "State": state,
                    "Phone": row.get("phone", ""),
                    "Web Address": "",
                    "_detail_url": row.get("detailUrl", ""),
                }
                all_members.append(record)

            log.info("  Collected %d members on page %d (total: %d)",
                     len(raw_rows), page_num, len(all_members))
        else:
            log.info("  No data rows on page %d.", page_num)
            if page_num == 1:
                page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_screenshot.png"), full_page=True)
                with open(os.path.join(OUTPUT_DIR, "debug_page.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())
                log.error("No data on first page. Debug files saved.")
            break

        # Navigate to next page using JavaScript click
        clicked = page.evaluate(JS_CLICK_NEXT)
        if clicked:
            page.wait_for_timeout(3000)
            log.info("  Navigated to next page.")
        else:
            # Fallback: try URL Range parameter
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
                except Exception as exc:
                    log.warning("  Range pagination failed: %s. Stopping.", exc)
                    break
            else:
                log.info("No more pages. Stopping pagination.")
                break

    log.info("Total members collected: %d", len(all_members))
    return all_members


# ---------------------------------------------------------------------------
# Phase 2: Visit detail pages for Web Address
# ---------------------------------------------------------------------------
def scrape_web_address(page, detail_url):
    if not detail_url:
        return ""

    try:
        retry(lambda: page.goto(detail_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    except Exception as exc:
        log.warning("  Could not load detail page: %s", exc)
        return ""

    page.wait_for_timeout(500)

    # Use JavaScript to extract the web address
    web = page.evaluate(JS_EXTRACT_WEB_ADDRESS)
    return (web or "").strip()


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def export_csv(records):
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)
    log.info("CSV saved to %s (%d rows)", CSV_FILE, len(records))


def export_xlsx(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "NGA Members"
    ws.append(FIELDS)
    for rec in records:
        ws.append([rec.get(f, "") for f in FIELDS])
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
        # Launch with stealth flags to avoid bot detection
        browser = pw.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/131.0.0.0 Safari/537.36"
            ),
            viewport={"width": 1280, "height": 900},
        )
        # Remove the webdriver flag that marks headless browsers
        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => false})")
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

            if idx % 50 == 0 or idx == 1:
                log.info("[%d/%d] %s", idx, total, member["Name"])

            try:
                web = scrape_web_address(page, detail_url)
                member["Web Address"] = web
                if web:
                    web_found += 1
            except Exception as exc:
                log.warning("  Error fetching web address for %s: %s", member["Name"], exc)
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
