#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Key techniques:
  - Network request interception to discover the real results URL
  - Playwright native clicks (not JS evaluate) for pagination
  - Multiple fallback strategies for every step
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
SEARCH_URL = "https://members.glass.org/cvweb/cgi-bin/utilities.dll/OpenPage?wrp=ngaSearch.htm"

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
CSV_FILE = os.path.join(OUTPUT_DIR, "nga_members.csv")
XLSX_FILE = os.path.join(OUTPUT_DIR, "nga_members.xlsx")

FIELDS = ["Name", "Classification", "Address", "City", "State", "Phone", "Web Address"]

DETAIL_DELAY = 0.4
PAGE_LOAD_TIMEOUT = 60_000
RETRY_COUNT = 3
RETRY_BACKOFF = 2

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)


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
    return (m.group(1).strip(), m.group(2).strip()) if m else (text, "")


# ---------------------------------------------------------------------------
# JS: extract table rows
# ---------------------------------------------------------------------------
JS_EXTRACT_TABLE = """
() => {
    const tables = document.querySelectorAll('table');
    let dataRows = [];
    for (const table of tables) {
        for (const row of table.querySelectorAll('tr')) {
            const cells = Array.from(row.querySelectorAll('td, th'));
            if (cells.length < 5) continue;
            const texts = cells.map(c => (c.innerText || '').trim());
            const first = texts[0].toLowerCase();
            if (!first || first === 'name' || first === 'organization' ||
                first === 'company') continue;
            if (first.includes('\u25bc') || first.includes('\u25b2')) continue;
            if (/^\d+$/.test(texts[0]) || texts[0].includes('MATCH')) continue;

            // Get detail link — capture href AND onclick
            let detailUrl = '';
            let detailOnclick = '';
            const link = cells[0].querySelector('a');
            if (link) {
                detailUrl = link.getAttribute('href') || '';
                detailOnclick = link.getAttribute('onclick') || '';
                // If href is a real URL (absolute), use link.href for full URL
                if (link.href && link.href.startsWith('http'))
                    detailUrl = link.href;
            }
            dataRows.push({
                name: texts[0],
                classification: texts[1] || '',
                address: texts[2] || '',
                cityState: texts[3] || '',
                phone: texts[4] || '',
                detailUrl: detailUrl,
                detailOnclick: detailOnclick
            });
        }
    }
    const seen = new Set();
    return dataRows.filter(r => {
        const key = r.name + '|' + r.address;
        if (seen.has(key) || !r.name) return false;
        seen.add(key);
        return true;
    });
}
"""

# Dump the first 5 pagination-area elements with ALL attributes
JS_DUMP_PAGINATION = """
() => {
    const info = [];
    // Find all <a> tags and look for ones with numeric text or arrow chars
    for (const a of document.querySelectorAll('a')) {
        const text = (a.innerText || '').trim();
        if (!text) continue;
        const isNumeric = /^\d+$/.test(text);
        const isArrow = text.length <= 2 && /[→›»>▶➡]/.test(text);
        const isNext = text.toLowerCase() === 'next';
        if (isNumeric || isArrow || isNext) {
            info.push({
                text: text,
                href: a.getAttribute('href') || '',
                onclick: a.getAttribute('onclick') || '',
                fullHref: a.href || '',
                className: a.className || '',
                parentTag: a.parentElement ? a.parentElement.tagName : ''
            });
        }
        if (info.length >= 15) break;
    }
    return info;
}
"""

JS_EXTRACT_WEB_ADDRESS = """
() => {
    // Strategy 1: td/th cells with "Web Address" or similar label
    const labels = ['web address', 'website', 'web site', 'url', 'home page', 'web'];
    const allCells = Array.from(document.querySelectorAll('td, th'));
    for (const label of labels) {
        for (let i = 0; i < allCells.length; i++) {
            const raw = (allCells[i].innerText || '').trim();
            const cellText = raw.replace(/:/g, '').trim().toLowerCase();
            if (cellText !== label) continue;
            for (let j = i + 1; j < Math.min(i + 4, allCells.length); j++) {
                const next = allCells[j];
                const link = next.querySelector('a[href]');
                if (link) {
                    const href = link.href || '';
                    if (href.startsWith('http') && !href.includes('glass.org') && !href.includes('mailto:'))
                        return href;
                }
                const txt = (next.innerText || '').trim();
                if (txt && txt.includes('.') && !txt.includes(' ') && txt.length > 3 && !txt.includes('glass.org')) {
                    return txt.startsWith('http') ? txt : 'http://' + txt;
                }
            }
        }
    }
    // Strategy 2: <a> whose text looks like a URL
    for (const a of document.querySelectorAll('a[href]')) {
        const text = (a.innerText || '').trim().toLowerCase();
        const href = a.href || '';
        if ((text.startsWith('www.') || text.match(/^https?:\\/\\//)) &&
            !href.includes('glass.org') && !href.includes('mailto:'))
            return href.startsWith('http') ? href : 'http://' + text;
    }
    // Strategy 3: external link not to glass.org or common sites
    const skip = ['glass.org','google.com','facebook.com','twitter.com','linkedin.com',
                  'youtube.com','instagram.com','bing.com','javascript:','mailto:','cvweb'];
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
    captured_requests = []

    # Set up request interception BEFORE navigating
    def on_request(request):
        url = request.url.lower()
        if any(k in url for k in ['customlist', 'range=', 'orgsearch', 'organizationdll']):
            captured_requests.append({
                'url': request.url,
                'method': request.method,
                'post_data': request.post_data,
            })

    page.on('request', on_request)

    # Navigate to search page
    log.info("Navigating to: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)
    log.info("Search page title: %s", page.title())
    log.info("Search page URL: %s", page.url)

    # Submit the search form
    submitted = False
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
                btn.click(timeout=10000)
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(2000)
                submitted = True
                break
        except Exception:
            continue
    if not submitted:
        try:
            inp = page.locator("input[type='text']").first
            if inp.count():
                inp.press("Enter")
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(2000)
                submitted = True
        except Exception:
            pass

    # Log everything we know about the results page
    results_url = page.url
    log.info("Results page URL: %s", results_url)
    log.info("Results page title: %s", page.title())

    # Log captured network requests
    log.info("Captured %d relevant network requests:", len(captured_requests))
    for cr in captured_requests:
        log.info("  %s %s", cr['method'], cr['url'][:200])
        if cr['post_data']:
            log.info("    POST data: %s", str(cr['post_data'])[:300])

    # Find a results URL with Range parameter from captured requests
    results_base_with_range = ""
    for cr in captured_requests:
        if 'range=' in cr['url'].lower() and 'customlist' in cr['url'].lower():
            results_base_with_range = cr['url']
            break
    if not results_base_with_range:
        for cr in captured_requests:
            if 'range=' in cr['url'].lower():
                results_base_with_range = cr['url']
                break

    log.info("Results URL with Range: %s", results_base_with_range or "(none found)")

    # Save debug files
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"), full_page=True)
    with open(os.path.join(OUTPUT_DIR, "results_page1.html"), "w", encoding="utf-8") as f:
        f.write(page.content())

    # Dump pagination link info
    pag_info = page.evaluate(JS_DUMP_PAGINATION)
    log.info("Pagination elements found: %d", len(pag_info))
    for pi in pag_info:
        log.info("  text=%s href=%s onclick=%s fullHref=%s",
                 pi.get('text',''), pi.get('href','')[:100],
                 pi.get('onclick','')[:100], pi.get('fullHref','')[:100])

    # Extract page 1 data
    raw_rows = page.evaluate(JS_EXTRACT_TABLE)
    log.info("Page 1: extracted %d rows", len(raw_rows))

    if not raw_rows:
        log.error("No data on page 1!")
        return all_members

    # Log first row detail URL info for debugging
    if raw_rows:
        log.info("First row detail: url=%s onclick=%s",
                 raw_rows[0].get('detailUrl', '')[:150],
                 raw_rows[0].get('detailOnclick', '')[:150])

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

    first_name_page1 = raw_rows[0].get("name", "") if raw_rows else ""

    # ===================================================================
    # PAGINATION — try multiple strategies
    # ===================================================================

    # ----- Strategy A: Big Range URL from captured request -----
    if results_base_with_range:
        big_url = re.sub(r'[Rr]ange=\d+/\d+', 'Range=1/5000', results_base_with_range)
        log.info("Strategy A: trying big Range: %s", big_url[:200])
        try:
            page.goto(big_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)
            page.wait_for_timeout(5000)
            raw = page.evaluate(JS_EXTRACT_TABLE)
            log.info("  Big Range returned %d rows", len(raw))
            if len(raw) > len(all_members):
                all_members.clear()
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
                log.info("Strategy A SUCCESS: %d members", len(all_members))
                return all_members
        except Exception as exc:
            log.warning("Strategy A failed: %s", exc)

    # ----- Strategy B: Range pagination from captured or pagination URLs -----
    range_base = results_base_with_range
    if not range_base:
        # Check pagination link hrefs for Range pattern
        for pi in pag_info:
            for url_field in ['fullHref', 'href']:
                url = pi.get(url_field, '')
                if url and 'range=' in url.lower():
                    range_base = url
                    break
            if range_base:
                break

    if range_base and 'range=' in range_base.lower():
        rm = re.search(r'[Rr]ange=(\d+)/(\d+)', range_base)
        if rm:
            page_size = int(rm.group(2))
            # Try to get total from page text
            match_count = page.evaluate("""
                () => {
                    const m = (document.body.innerText || '').match(/(\\d[\\d,]*)\\s*MATCH/i);
                    return m ? parseInt(m[1].replace(/,/g,''), 10) : 0;
                }
            """)
            total_pages = (match_count + page_size - 1) // page_size if match_count else 100
            log.info("Strategy B: Range pagination — size=%d, total=%d, pages=%d",
                     page_size, match_count, total_pages)

            for pg in range(2, total_pages + 1):
                start = (pg - 1) * page_size + 1
                pg_url = re.sub(r'[Rr]ange=\d+/\d+', f'Range={start}/{page_size}', range_base)
                try:
                    retry(lambda u=pg_url: page.goto(u, wait_until="domcontentloaded",
                                                      timeout=PAGE_LOAD_TIMEOUT))
                    page.wait_for_timeout(1000)
                    raw = page.evaluate(JS_EXTRACT_TABLE)
                    if not raw:
                        log.info("  No rows on page %d. Stopping.", pg)
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
                    if pg % 10 == 0:
                        log.info("  Page %d/%d: total %d", pg, total_pages, len(all_members))
                except Exception as exc:
                    log.error("  Page %d failed: %s. Stopping.", pg, exc)
                    break

            if len(all_members) > 30:
                log.info("Strategy B SUCCESS: %d members", len(all_members))
                return all_members

    # ----- Strategy C: Playwright native click pagination -----
    log.info("Strategy C: Playwright click pagination")
    # Go back to page 1 results
    retry(lambda: page.goto(results_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    for page_num in range(2, 200):
        prev_count = len(all_members)
        first_name_before = all_members[-1]["Name"] if all_members else ""

        # Try to click the next page — try multiple selectors
        clicked = False
        for text_pattern in [str(page_num), "→", "›", "»", ">>", "Next", "next", ">"]:
            try:
                links = page.locator(f"a:text-is('{text_pattern}')").all()
                if not links:
                    continue
                for link in links:
                    if link.is_visible():
                        log.info("  Clicking pagination link: '%s'", text_pattern)
                        link.click(timeout=10000)
                        # Wait for page to change
                        page.wait_for_load_state("networkidle", timeout=15000)
                        page.wait_for_timeout(2000)
                        clicked = True
                        break
                if clicked:
                    break
            except Exception:
                continue

        if not clicked:
            log.info("  No pagination link found for page %d. Stopping.", page_num)
            break

        # Extract data from new page
        raw = page.evaluate(JS_EXTRACT_TABLE)
        if not raw:
            log.info("  No data on page %d. Stopping.", page_num)
            break

        # Verify page actually changed (not same data)
        new_first = raw[0].get("name", "") if raw else ""
        if new_first == first_name_page1 and page_num > 2:
            log.info("  Page didn't change (still showing page 1 data). Stopping.")
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

        if page_num % 10 == 0:
            log.info("  Page %d: %d new rows (total: %d)", page_num, len(raw), len(all_members))

        first_name_page1_ref = raw[0].get("name", "")  # track for next iteration

    log.info("Total members collected: %d", len(all_members))
    return all_members


# ---------------------------------------------------------------------------
# Phase 2: Web addresses from detail pages
# ---------------------------------------------------------------------------
def collect_web_addresses(page, members):
    total = len(members)
    log.info("Phase 2: Fetching web addresses from %d detail pages …", total)

    # Debug: check what detail URLs look like
    sample_urls = [m.get("_detail_url", "") for m in members[:5]]
    log.info("Sample detail URLs:")
    for i, u in enumerate(sample_urls):
        log.info("  [%d] %s", i, u[:200] if u else "(empty)")

    # If detail URLs are empty or javascript:, try clicking first member name
    first_url = members[0].get("_detail_url", "") if members else ""
    if not first_url or first_url.startswith("javascript:"):
        log.info("Detail URLs are empty or javascript. Trying to click first member name.")
        # Go back to results page 1 to click the member name
        try:
            retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(3000)
            # Submit search again
            for sel in ["input[type='submit']", "button[type='submit']",
                        "input[value='Search']", "button:has-text('Search')",
                        "input[type='image']"]:
                try:
                    btn = page.locator(sel).first
                    if btn.count() and btn.is_visible():
                        btn.click(timeout=10000)
                        page.wait_for_load_state("networkidle", timeout=15000)
                        page.wait_for_timeout(2000)
                        break
                except Exception:
                    continue

            # Click the first member name link
            first_member_name = members[0]["Name"]
            log.info("Clicking first member: %s", first_member_name)
            member_link = page.locator(f"a:text-is('{first_member_name}')").first
            if member_link.count():
                member_link.click(timeout=10000)
                page.wait_for_load_state("networkidle", timeout=15000)
                page.wait_for_timeout(2000)
                detail_url = page.url
                log.info("Detail page URL after click: %s", detail_url)
                page.screenshot(path=os.path.join(OUTPUT_DIR, "detail_page1.png"), full_page=True)
                with open(os.path.join(OUTPUT_DIR, "detail_page1.html"), "w", encoding="utf-8") as f:
                    f.write(page.content())

                # Try extracting web address
                web = page.evaluate(JS_EXTRACT_WEB_ADDRESS)
                log.info("First member web address: %s", web or "(empty)")
                members[0]["Web Address"] = web or ""

                # If we got a real URL with orgcd, build a template
                orgcd_match = re.search(r'orgcd=(\d+)', detail_url, re.IGNORECASE)
                if orgcd_match:
                    url_template = detail_url
                    first_orgcd = orgcd_match.group(1)
                    log.info("Discovered detail URL template with orgcd=%s", first_orgcd)

                    # Now try to find orgcd in the original table links
                    # Re-extract to get orgcds from onclick or href
                    # For now, we have the URL template
        except Exception as exc:
            log.warning("Could not click first member: %s", exc)

    # Now iterate through all members and fetch web addresses
    web_found = 0
    web_failed = 0
    for idx, member in enumerate(members):
        if idx == 0 and member.get("Web Address"):
            web_found += 1
            continue

        detail_url = member.get("_detail_url", "")
        if not detail_url or detail_url.startswith("javascript:") or not detail_url.startswith("http"):
            web_failed += 1
            continue

        try:
            retry(lambda u=detail_url: page.goto(u, wait_until="domcontentloaded",
                                                   timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(400)
            web = page.evaluate(JS_EXTRACT_WEB_ADDRESS)
            member["Web Address"] = (web or "").strip()
            if web:
                web_found += 1

            # Save debug for first few
            if idx < 3:
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"detail_page{idx+1}.png"),
                                full_page=True)
                log.info("[%d] %s — web: %s", idx+1, member["Name"], web or "(empty)")

        except Exception as exc:
            log.warning("[%d] Error for %s: %s", idx+1, member["Name"], exc)
            web_failed += 1

        if idx > 0 and idx % 100 == 0:
            log.info("  Progress: %d/%d — web found: %d", idx, total, web_found)

        time.sleep(DETAIL_DELAY)

    log.info("Web addresses — found: %d, failed/empty: %d", web_found, web_failed)
    return web_found


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
def export_csv(records):
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)
    log.info("CSV saved: %s (%d rows)", CSV_FILE, len(records))


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
    log.info("XLSX saved: %s (%d rows)", XLSX_FILE, len(records))


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

        # Phase 1: collect all members
        members = collect_all_members(page)
        if not members:
            log.error("No members found.")
            browser.close()
            return

        # Phase 2: get web addresses
        collect_web_addresses(page, members)

        browser.close()

    # Phase 3: export
    export_csv(members)
    export_xlsx(members)

    log.info("=" * 60)
    log.info("DONE — Total members: %d", len(members))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
