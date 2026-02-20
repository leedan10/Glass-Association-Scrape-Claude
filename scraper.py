#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Strategy:
  Phase 1 — Submit the search form, then use URL-based Range pagination
            to iterate through all result pages. Extract table data via
            JavaScript DOM evaluation.
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
    m = re.match(r"^(.+?),\s*([A-Z]{2})\b", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return text, ""


# ---------------------------------------------------------------------------
# JavaScript snippets executed inside the browser
# ---------------------------------------------------------------------------

# Extract table data rows from the page.  Handles nested tables, mixed
# th/td cells, and cvweb layout quirks.
JS_EXTRACT_TABLE = """
() => {
    const tables = document.querySelectorAll('table');
    let dataRows = [];

    for (const table of tables) {
        const rows = table.querySelectorAll('tr');
        for (const row of rows) {
            const cells = Array.from(row.querySelectorAll('td, th'));
            if (cells.length < 5) continue;

            const texts = cells.map(c => (c.innerText || '').trim());

            // Skip header/nav/empty rows
            const first = texts[0].toLowerCase();
            if (!first || first === 'name' || first === 'organization' ||
                first === 'company' || first.includes('\\u25bc') ||
                first.includes('\\u25b2'))
                continue;
            if (/^\\d+$/.test(texts[0]) || texts[0].includes('Page') ||
                texts[0].includes('Match') || texts[0].includes('MATCH'))
                continue;

            let detailUrl = '';
            const link = cells[0].querySelector('a[href]');
            if (link) detailUrl = link.href || '';

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

    // Deduplicate
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

# Get the href of the next-page link (don't click — just return the URL).
JS_GET_NEXT_PAGE_URL = """
() => {
    const arrows = ['\\u2192', '\\u203a', '\\u00bb', '>>', '\\u25b6', 'Next', 'next'];
    const allLinks = document.querySelectorAll('a[href]');
    for (const a of allLinks) {
        const text = (a.innerText || '').trim();
        for (const arrow of arrows) {
            if (text === arrow || text.toLowerCase() === arrow.toLowerCase()) {
                return a.href;
            }
        }
    }
    return '';
}
"""

# Extract the total match count from page text, e.g. "2106 MATCH(ES)"
JS_GET_MATCH_COUNT = """
() => {
    const body = document.body.innerText || '';
    const m = body.match(/(\\d[\\d,]*)\\s*MATCH/i);
    return m ? parseInt(m[1].replace(/,/g, ''), 10) : 0;
}
"""

# Extract the web address from a member detail page.
JS_EXTRACT_WEB_ADDRESS = """
() => {
    const body = document.body.innerText || '';
    const labels = ['Web Address', 'Website', 'Web Site', 'URL', 'Home Page', 'Web'];

    // Strategy 1: table-cell label/value pairs
    for (const label of labels) {
        const allCells = document.querySelectorAll('td, th');
        for (let i = 0; i < allCells.length; i++) {
            const cellText = (allCells[i].innerText || '').trim()
                              .replace(/[:]/g, '').trim().toLowerCase();
            if (cellText === label.toLowerCase()) {
                const next = allCells[i + 1] || allCells[i].nextElementSibling;
                if (next) {
                    const link = next.querySelector('a[href]');
                    if (link && link.href && !link.href.includes('glass.org') &&
                        !link.href.includes('mailto:')) {
                        return link.href;
                    }
                    const txt = (next.innerText || '').trim();
                    if (txt && txt.includes('.') && !txt.includes(' ') && txt.length > 3) {
                        return txt.startsWith('http') ? txt : 'http://' + txt;
                    }
                }
            }
        }
    }

    // Strategy 2: regex in body text
    for (const label of labels) {
        const re = new RegExp(label + '[:\\\\-]?\\\\s*(https?://[^\\\\s]+)', 'i');
        const m = body.match(re);
        if (m) return m[1].replace(/[,;)]+$/, '');
    }

    // Strategy 3: any external link not to glass.org / social / maps
    const skip = ['glass.org', 'google.com', 'facebook.com', 'twitter.com',
                  'linkedin.com', 'youtube.com', 'instagram.com', 'bing.com',
                  'javascript:', 'mailto:'];
    for (const a of document.querySelectorAll('a[href]')) {
        const href = a.href || '';
        if (!href.startsWith('http')) continue;
        if (skip.some(s => href.includes(s))) continue;
        return href;
    }
    return '';
}
"""


# ---------------------------------------------------------------------------
# Phase 1: Collect all members from paginated results
# ---------------------------------------------------------------------------
def collect_all_members(page):
    all_members = []

    log.info("Navigating to search page: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)
    log.info("Page title: %s", page.title())

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
                log.info("Clicking submit: %s", sel)
                btn.click(timeout=5000)
                page.wait_for_timeout(5000)
                submitted = True
                break
        except Exception:
            continue
    if not submitted:
        try:
            first_input = page.locator("input[type='text']").first
            if first_input.count():
                first_input.press("Enter")
                page.wait_for_timeout(5000)
                submitted = True
                log.info("Pressed Enter on first text input.")
        except Exception:
            pass
    if not submitted:
        log.warning("Could not find submit button. Parsing current page anyway.")

    # Save page 1 debug files
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"), full_page=True)
    with open(os.path.join(OUTPUT_DIR, "results_page1.html"), "w", encoding="utf-8") as f:
        f.write(page.content())

    # Get total match count and current URL
    total_matches = page.evaluate(JS_GET_MATCH_COUNT)
    log.info("Total matches reported by page: %d", total_matches)

    current_url = page.url
    log.info("Results URL: %s", current_url)

    # --- Determine pagination strategy ---
    # Check if URL already has a Range parameter
    range_match = re.search(r"Range=(\d+)/(\d+)", current_url, re.IGNORECASE)
    if range_match:
        page_size = int(range_match.group(2))
    else:
        page_size = 25  # default cvweb page size

    # Also try to get the next-page URL from the → link
    next_url_from_link = page.evaluate(JS_GET_NEXT_PAGE_URL)
    log.info("Next-page URL from arrow link: %s", next_url_from_link or "(none)")

    # If we got a next-page URL, extract the Range pattern from it
    if next_url_from_link:
        nm = re.search(r"Range=(\d+)/(\d+)", next_url_from_link, re.IGNORECASE)
        if nm:
            page_size = int(nm.group(2))
            log.info("Detected page size from next link: %d", page_size)

    # Calculate total pages
    if total_matches > 0:
        total_pages = (total_matches + page_size - 1) // page_size
    else:
        total_pages = 200  # safety fallback
    log.info("Expected pages: %d (page_size=%d)", total_pages, page_size)

    # --- Scrape page 1 ---
    raw_rows = page.evaluate(JS_EXTRACT_TABLE)
    log.info("Page 1: JS extraction returned %d rows", len(raw_rows))

    if not raw_rows:
        page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_screenshot.png"), full_page=True)
        with open(os.path.join(OUTPUT_DIR, "debug_page.html"), "w", encoding="utf-8") as f:
            f.write(page.content())
        log.error("No data on page 1. Debug files saved.")
        return all_members

    for row in raw_rows:
        city, state = split_city_state(row.get("cityState", ""))
        all_members.append({
            "Name": row.get("name", ""),
            "Classification": row.get("classification", ""),
            "Address": row.get("address", ""),
            "City": city,
            "State": state,
            "Phone": row.get("phone", ""),
            "Web Address": "",
            "_detail_url": row.get("detailUrl", ""),
        })
    log.info("Page 1: collected %d members", len(all_members))

    # --- Determine base URL for pagination ---
    # Use the next-page link as a template if available; otherwise modify
    # the current URL.
    if next_url_from_link and "Range=" in next_url_from_link:
        base_pagination_url = next_url_from_link
    elif "Range=" in current_url:
        base_pagination_url = current_url
    else:
        # Construct pagination URL from next-page link (full URL)
        base_pagination_url = next_url_from_link if next_url_from_link else ""

    if not base_pagination_url:
        log.warning("Cannot determine pagination URL. Only page 1 data collected.")
        return all_members

    log.info("Base pagination URL: %s", base_pagination_url)

    # --- Scrape pages 2 through N ---
    for page_num in range(2, total_pages + 1):
        range_start = (page_num - 1) * page_size + 1
        # Replace the Range= parameter in the URL
        page_url = re.sub(
            r"Range=\d+/\d+",
            f"Range={range_start}/{page_size}",
            base_pagination_url,
        )

        log.info("Page %d/%d — navigating to Range=%d/%d …",
                 page_num, total_pages, range_start, page_size)

        try:
            retry(lambda url=page_url: page.goto(url, wait_until="domcontentloaded",
                                                  timeout=PAGE_LOAD_TIMEOUT))
        except Exception as exc:
            log.error("  Failed to load page %d: %s. Stopping.", page_num, exc)
            break

        page.wait_for_timeout(1500)

        raw_rows = page.evaluate(JS_EXTRACT_TABLE)
        if not raw_rows:
            log.info("  No rows on page %d. Stopping.", page_num)
            break

        for row in raw_rows:
            city, state = split_city_state(row.get("cityState", ""))
            all_members.append({
                "Name": row.get("name", ""),
                "Classification": row.get("classification", ""),
                "Address": row.get("address", ""),
                "City": city,
                "State": state,
                "Phone": row.get("phone", ""),
                "Web Address": "",
                "_detail_url": row.get("detailUrl", ""),
            })

        log.info("  Page %d: %d rows (total: %d)", page_num, len(raw_rows), len(all_members))

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
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => false})"
        )
        page = context.new_page()

        # Phase 1
        members = collect_all_members(page)
        if not members:
            log.error("No members found. Check debug files in output/.")
            browser.close()
            return

        # Phase 2
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
            if idx % 100 == 0 or idx == 1:
                log.info("[%d/%d] %s", idx, total, member["Name"])
            try:
                web = scrape_web_address(page, detail_url)
                member["Web Address"] = web
                if web:
                    web_found += 1
            except Exception as exc:
                log.warning("  Error: %s", exc)
                web_failed += 1
            if idx < total:
                time.sleep(DETAIL_DELAY)

        browser.close()

    # Phase 3
    export_csv(members)
    export_xlsx(members)

    log.info("=" * 60)
    log.info("DONE — Members: %d | Web addresses found: %d | Failed: %d",
             total, web_found, web_failed)
    log.info("=" * 60)


if __name__ == "__main__":
    main()
