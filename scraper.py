#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Strategy:
  Phase 1 — Submit search, then grab ALL pagination link hrefs from the page.
            Try requesting all results at once (Range=1/5000). If that doesn't
            yield more results, iterate every pagination URL found on page 1.
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

DETAIL_DELAY = 0.5
PAGE_LOAD_TIMEOUT = 60_000
RETRY_COUNT = 3
RETRY_BACKOFF = 2

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


def split_city_state(text):
    text = text.strip()
    if not text:
        return "", ""
    m = re.match(r"^(.+?),\s*([A-Z]{2})\b", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return text, ""


# ---------------------------------------------------------------------------
# JavaScript snippets
# ---------------------------------------------------------------------------

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
            const first = texts[0].toLowerCase();

            // Skip headers, empty, pagination, and match-count rows
            if (!first) continue;
            if (first === 'name' || first === 'organization' || first === 'company') continue;
            if (first.includes('\u25bc') || first.includes('\u25b2')) continue;
            if (/^\d+$/.test(texts[0])) continue;
            if (texts[0].includes('Page') || texts[0].includes('MATCH')) continue;

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

    // Deduplicate by name+address
    const seen = new Set();
    return dataRows.filter(r => {
        const key = r.name + '|' + r.address;
        if (seen.has(key) || !r.name) return false;
        seen.add(key);
        return true;
    });
}
"""

# Return ALL link hrefs on the page (for debugging pagination).
JS_GET_ALL_LINKS = """
() => {
    return Array.from(document.querySelectorAll('a[href]')).map(a => ({
        text: (a.innerText || '').trim().substring(0, 50),
        href: a.href
    }));
}
"""

# Get every pagination-area link: look for links whose text is a number,
# or arrows / "Next".  Return their hrefs.
JS_GET_PAGINATION_LINKS = """
() => {
    const results = [];
    const allLinks = document.querySelectorAll('a[href]');
    for (const a of allLinks) {
        const text = (a.innerText || '').trim();
        const href = a.href || '';
        if (!href) continue;
        // Numeric page links: "2", "3", "85"
        if (/^\d+$/.test(text)) {
            results.push({ text, href });
            continue;
        }
        // Arrow / next links (many possible Unicode arrows)
        const code = text.charCodeAt(0);
        if (text.length <= 2 && (
            text === '>' || text === '>>' ||
            code === 0x2192 || code === 0x203A || code === 0x00BB ||
            code === 0x25B6 || code === 0x27A1 || code === 0x2023 ||
            code === 0x276F || code === 0x279C
        )) {
            results.push({ text: '[arrow]', href });
            continue;
        }
        if (text.toLowerCase() === 'next') {
            results.push({ text: 'next', href });
        }
    }
    return results;
}
"""

JS_GET_MATCH_COUNT = """
() => {
    const body = document.body.innerText || '';
    const m = body.match(/(\\d[\\d,]*)\\s*MATCH/i);
    return m ? parseInt(m[1].replace(/,/g, ''), 10) : 0;
}
"""

JS_EXTRACT_WEB_ADDRESS = """
() => {
    // Strategy 1: Find table cells labelled "Web Address", "Website", etc.
    const labels = ['web address', 'website', 'web site', 'url', 'home page', 'web'];
    const allCells = Array.from(document.querySelectorAll('td, th'));
    for (const label of labels) {
        for (let i = 0; i < allCells.length; i++) {
            const cellText = (allCells[i].innerText || '').trim()
                              .replace(/:/g, '').trim().toLowerCase();
            if (cellText !== label) continue;
            // Look at next cell(s)
            for (let j = i + 1; j < Math.min(i + 3, allCells.length); j++) {
                const next = allCells[j];
                // Check for an anchor tag
                const link = next.querySelector('a[href]');
                if (link) {
                    const href = link.href || '';
                    if (href.startsWith('http') && !href.includes('glass.org') && !href.includes('mailto:'))
                        return href;
                }
                // Check text
                const txt = (next.innerText || '').trim();
                if (txt && txt.includes('.') && !txt.includes(' ') && txt.length > 3 && !txt.includes('glass.org')) {
                    return txt.startsWith('http') ? txt : 'http://' + txt;
                }
            }
        }
    }

    // Strategy 2: Look for a link with visible text that looks like a URL (www.something.com)
    for (const a of document.querySelectorAll('a[href]')) {
        const href = a.href || '';
        const text = (a.innerText || '').trim().toLowerCase();
        if (text.startsWith('www.') || text.startsWith('http')) {
            if (!href.includes('glass.org') && !href.includes('mailto:'))
                return href.startsWith('http') ? href : 'http://' + text;
        }
    }

    // Strategy 3: Any external link not to glass.org / social / maps / nav
    const skip = ['glass.org', 'google.com', 'facebook.com', 'twitter.com',
                  'linkedin.com', 'youtube.com', 'instagram.com', 'bing.com',
                  'javascript:', 'mailto:', 'maps.google', 'cvweb'];
    for (const a of document.querySelectorAll('a[href]')) {
        const href = a.href || '';
        if (!href.startsWith('http')) continue;
        if (skip.some(s => href.toLowerCase().includes(s))) continue;
        return href;
    }
    return '';
}
"""


# ---------------------------------------------------------------------------
# Phase 1: Collect all members
# ---------------------------------------------------------------------------
def collect_all_members(page):
    all_members = []

    log.info("Navigating to search page: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)
    log.info("Page title: %s", page.title())

    # Submit the search form
    for sel in [
        "input[type='submit']", "button[type='submit']",
        "input[value='Search']", "button:has-text('Search')",
        "input[value='Find']", "button:has-text('Find')",
        "a:has-text('Search')", "input[type='image']",
    ]:
        try:
            btn = page.locator(sel).first
            if btn.count() and btn.is_visible():
                log.info("Clicking submit: %s", sel)
                btn.click(timeout=5000)
                page.wait_for_timeout(5000)
                break
        except Exception:
            continue

    # Save debug files
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"), full_page=True)
    with open(os.path.join(OUTPUT_DIR, "results_page1.html"), "w", encoding="utf-8") as f:
        f.write(page.content())

    current_url = page.url
    total_matches = page.evaluate(JS_GET_MATCH_COUNT)
    log.info("Results URL: %s", current_url)
    log.info("Total matches: %d", total_matches)

    # ----- Discover pagination links -----
    pagination_links = page.evaluate(JS_GET_PAGINATION_LINKS)
    log.info("Pagination links found: %d", len(pagination_links))
    for pl in pagination_links[:10]:
        log.info("  page-link: text=%s href=%s", pl.get("text", ""), pl.get("href", "")[:120])

    # Identify page size and base URL from pagination links
    page2_url = ""
    page_size = 25  # default
    for pl in pagination_links:
        if pl.get("text") == "2":
            page2_url = pl["href"]
            break
    if not page2_url:
        # Take the first arrow/next link
        for pl in pagination_links:
            if pl.get("text") in ("[arrow]", "next"):
                page2_url = pl["href"]
                break

    if page2_url:
        log.info("Page 2 URL: %s", page2_url)
        rm = re.search(r"Range=(\d+)/(\d+)", page2_url, re.IGNORECASE)
        if rm:
            page_size = int(rm.group(2))
            log.info("Page size from page-2 link: %d", page_size)

    # ----- Strategy A: Try to get ALL results in one request -----
    # Modify the URL to request Range=1/5000 (all at once)
    big_range_url = ""
    if page2_url and "Range=" in page2_url:
        big_range_url = re.sub(r"Range=\d+/\d+", "Range=1/5000", page2_url)
    elif "Range=" in current_url:
        big_range_url = re.sub(r"Range=\d+/\d+", "Range=1/5000", current_url)

    if big_range_url:
        log.info("Trying big-range request: %s", big_range_url[:150])
        try:
            page.goto(big_range_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)
            page.wait_for_timeout(5000)
            raw = page.evaluate(JS_EXTRACT_TABLE)
            log.info("Big-range returned %d rows", len(raw))
            if len(raw) > 30:  # more than one page's worth = success
                for row in raw:
                    city, state = split_city_state(row.get("cityState", ""))
                    all_members.append({
                        "Name": row.get("name", ""),
                        "Classification": row.get("classification", ""),
                        "Address": row.get("address", ""),
                        "City": city, "State": state,
                        "Phone": row.get("phone", ""),
                        "Web Address": "",
                        "_detail_url": row.get("detailUrl", ""),
                    })
                log.info("Big-range SUCCESS: got %d members in one request.", len(all_members))
                return all_members
            else:
                log.info("Big-range only returned %d rows — server may cap page size. Falling back to pagination.", len(raw))
        except Exception as exc:
            log.warning("Big-range failed: %s. Falling back to pagination.", exc)

    # ----- Strategy B: Paginate using discovered URLs -----
    # Go back to page 1 results
    retry(lambda: page.goto(current_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    # Scrape page 1
    raw_rows = page.evaluate(JS_EXTRACT_TABLE)
    log.info("Page 1: %d rows", len(raw_rows))
    if not raw_rows:
        page.screenshot(path=os.path.join(OUTPUT_DIR, "debug_screenshot.png"), full_page=True)
        log.error("No data on page 1.")
        return all_members

    for row in raw_rows:
        city, state = split_city_state(row.get("cityState", ""))
        all_members.append({
            "Name": row.get("name", ""),
            "Classification": row.get("classification", ""),
            "Address": row.get("address", ""),
            "City": city, "State": state,
            "Phone": row.get("phone", ""),
            "Web Address": "",
            "_detail_url": row.get("detailUrl", ""),
        })

    # Determine base URL for pagination
    base_url = page2_url if (page2_url and "Range=" in page2_url) else ""
    if not base_url and "Range=" in current_url:
        base_url = current_url

    if not base_url:
        # Last resort: try using each discovered page link directly
        log.warning("No Range-based URL found. Trying each pagination link directly.")
        visited = {current_url}
        for pl in pagination_links:
            href = pl.get("href", "")
            if href and href not in visited:
                visited.add(href)
                log.info("Visiting page link: text=%s", pl.get("text", ""))
                try:
                    page.goto(href, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)
                    page.wait_for_timeout(2000)
                    raw = page.evaluate(JS_EXTRACT_TABLE)
                    for row in raw:
                        city, state = split_city_state(row.get("cityState", ""))
                        all_members.append({
                            "Name": row.get("name", ""),
                            "Classification": row.get("classification", ""),
                            "Address": row.get("address", ""),
                            "City": city, "State": state,
                            "Phone": row.get("phone", ""),
                            "Web Address": "",
                            "_detail_url": row.get("detailUrl", ""),
                        })
                    log.info("  Got %d rows (total: %d)", len(raw), len(all_members))
                    # Now get NEW pagination links from this page (there might be more)
                    new_links = page.evaluate(JS_GET_PAGINATION_LINKS)
                    for nl in new_links:
                        nh = nl.get("href", "")
                        if nh and nh not in visited:
                            pagination_links.append(nl)
                except Exception as exc:
                    log.warning("  Failed: %s", exc)
        log.info("Total collected via direct links: %d", len(all_members))
        return all_members

    # Range-based pagination for pages 2..N
    if total_matches > 0:
        total_pages = (total_matches + page_size - 1) // page_size
    else:
        total_pages = 100

    log.info("Paginating: %d pages (size=%d) using base: %s", total_pages, page_size, base_url[:120])

    for page_num in range(2, total_pages + 1):
        start = (page_num - 1) * page_size + 1
        page_url = re.sub(r"Range=\d+/\d+", f"Range={start}/{page_size}", base_url)

        if page_num % 10 == 0 or page_num == 2:
            log.info("Page %d/%d — Range=%d/%d", page_num, total_pages, start, page_size)

        try:
            retry(lambda u=page_url: page.goto(u, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
        except Exception as exc:
            log.error("Failed page %d: %s. Stopping.", page_num, exc)
            break

        page.wait_for_timeout(1000)
        raw = page.evaluate(JS_EXTRACT_TABLE)
        if not raw:
            log.info("No rows on page %d. Stopping.", page_num)
            break

        for row in raw:
            city, state = split_city_state(row.get("cityState", ""))
            all_members.append({
                "Name": row.get("name", ""),
                "Classification": row.get("classification", ""),
                "Address": row.get("address", ""),
                "City": city, "State": state,
                "Phone": row.get("phone", ""),
                "Web Address": "",
                "_detail_url": row.get("detailUrl", ""),
            })

        if page_num % 10 == 0 or page_num == 2:
            log.info("  Page %d: %d rows (total: %d)", page_num, len(raw), len(all_members))

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
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
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

        # Phase 1 — collect all members
        members = collect_all_members(page)
        if not members:
            log.error("No members found. Check debug files in output/.")
            browser.close()
            return

        # Phase 2 — get web addresses
        total = len(members)
        log.info("=" * 60)
        log.info("Phase 2: Fetching Web Address from %d detail pages …", total)
        log.info("=" * 60)

        # Debug: visit first member and dump page info
        first_url = members[0].get("_detail_url", "")
        log.info("First member detail URL: %s", first_url)
        if first_url:
            try:
                page.goto(first_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)
                page.wait_for_timeout(2000)
                page.screenshot(path=os.path.join(OUTPUT_DIR, "detail_page1.png"), full_page=True)
                with open(os.path.join(OUTPUT_DIR, "detail_page1.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())
                log.info("Saved detail_page1 debug files.")
                # Try extraction on first page
                web = page.evaluate(JS_EXTRACT_WEB_ADDRESS)
                log.info("First member web address: %s", web or "(empty)")
                members[0]["Web Address"] = web or ""
            except Exception as exc:
                log.warning("Could not debug first detail page: %s", exc)

        web_found = 0
        web_failed = 0
        # Start from index 1 since we already did 0 above
        for idx, member in enumerate(members, 1):
            if idx == 1:
                if members[0]["Web Address"]:
                    web_found += 1
                continue  # already done above

            detail_url = member.get("_detail_url", "")
            if not detail_url:
                web_failed += 1
                continue

            if idx % 100 == 0:
                log.info("[%d/%d] %s — found so far: %d", idx, total, member["Name"], web_found)

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

    # Phase 3 — export
    export_csv(members)
    export_xlsx(members)

    log.info("=" * 60)
    log.info("DONE — Members: %d | Web addresses found: %d | Failed: %d",
             total, web_found, web_failed)
    log.info("=" * 60)


if __name__ == "__main__":
    main()
