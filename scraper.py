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

            // Get detail link — capture href, onclick, AND orgcd
            let detailUrl = '';
            let detailOnclick = '';
            let orgcd = '';
            const link = cells[0].querySelector('a');
            if (link) {
                detailUrl = link.getAttribute('href') || '';
                detailOnclick = link.getAttribute('onclick') || '';
                // If href is a real URL (absolute), use link.href for full URL
                if (link.href && link.href.startsWith('http'))
                    detailUrl = link.href;
                // Extract orgcd from onclick or href
                const onclickStr = detailOnclick + ' ' + detailUrl + ' ' + (link.href || '');
                const orgcdMatch = onclickStr.match(/orgcd=(\\d+)/i);
                if (orgcdMatch) orgcd = orgcdMatch[1];
            }
            dataRows.push({
                name: texts[0],
                classification: texts[1] || '',
                address: texts[2] || '',
                cityState: texts[3] || '',
                phone: texts[4] || '',
                detailUrl: detailUrl,
                detailOnclick: detailOnclick,
                orgcd: orgcd
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
    // Helper: fix double-protocol bug (e.g. http://https://example.com -> https://example.com)
    function fixUrl(url) {
        if (!url) return url;
        url = url.trim();
        // Fix double protocol: http://https:// or https://http://
        url = url.replace(/^https?:\\/\\/https?:\\/\\//i, function(match) {
            // Keep the inner protocol
            const inner = match.match(/https?:\\/\\/$/i);
            return inner ? inner[0] : 'https://';
        });
        // Simpler approach: repeatedly strip leading http(s):// until only one remains
        while (/^https?:\\/\\/https?:\\/\\//i.test(url)) {
            url = url.replace(/^https?:\\/\\//i, '');
        }
        if (!url.startsWith('http')) url = 'https://' + url;
        return url;
    }

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
                    const href = link.getAttribute('href') || '';
                    if (href && !href.includes('glass.org') && !href.includes('mailto:') && !href.startsWith('javascript:'))
                        return fixUrl(href);
                }
                const txt = (next.innerText || '').trim();
                if (txt && txt.includes('.') && !txt.includes(' ') && txt.length > 3 && !txt.includes('glass.org')) {
                    return fixUrl(txt);
                }
            }
        }
    }
    // Strategy 2: <a> whose text looks like a URL
    for (const a of document.querySelectorAll('a[href]')) {
        const text = (a.innerText || '').trim().toLowerCase();
        const href = a.getAttribute('href') || '';
        if ((text.startsWith('www.') || text.match(/^https?:\\/\\//)) &&
            !href.includes('glass.org') && !href.includes('mailto:'))
            return fixUrl(href || text);
    }
    // Strategy 3: external link not to glass.org or common sites
    const skip = ['glass.org','google.com','facebook.com','twitter.com','linkedin.com',
                  'youtube.com','instagram.com','bing.com','javascript:','mailto:','cvweb'];
    for (const a of document.querySelectorAll('a[href]')) {
        const href = a.getAttribute('href') || '';
        if (!href || href.startsWith('javascript:') || href.startsWith('mailto:')) continue;
        const full = href.startsWith('http') ? href : 'http://' + href;
        if (skip.some(s => full.toLowerCase().includes(s))) continue;
        return fixUrl(href);
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
        log.info("First row detail: url=%s onclick=%s orgcd=%s",
                 raw_rows[0].get('detailUrl', '')[:150],
                 raw_rows[0].get('detailOnclick', '')[:150],
                 raw_rows[0].get('orgcd', ''))
        orgcd_count = sum(1 for r in raw_rows if r.get('orgcd'))
        log.info("Rows with orgcd on page 1: %d / %d", orgcd_count, len(raw_rows))

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
            "_orgcd": row.get("orgcd", ""),
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
                        "_orgcd": row.get("orgcd", ""),
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
                            "_orgcd": row.get("orgcd", ""),
                        })
                    if pg % 10 == 0:
                        log.info("  Page %d/%d: total %d", pg, total_pages, len(all_members))
                except Exception as exc:
                    log.error("  Page %d failed: %s. Stopping.", pg, exc)
                    break

            if len(all_members) > 30:
                log.info("Strategy B SUCCESS: %d members", len(all_members))
                return all_members

    # ----- Strategy C0: Discover pagination URL by clicking page 2 with interception -----
    # If pagination links are javascript:, clicking page "2" while intercepting requests
    # can reveal the actual URL pattern for subsequent pages
    if not range_base or 'range=' not in range_base.lower():
        log.info("Strategy C0: Click page 2 with request interception to discover URL pattern")
        try:
            retry(lambda: page.goto(results_url, wait_until="domcontentloaded",
                                    timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(3000)

            # Clear and re-capture requests
            captured_requests.clear()
            # Try clicking page "2"
            page2_link = page.locator("a:text-is('2')").first
            if page2_link.count() and page2_link.is_visible():
                log.info("  Clicking page 2 link to capture URL …")
                page2_link.click(timeout=10000)
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass
                page.wait_for_timeout(3000)

                # Check captured requests for a Range URL
                for cr in captured_requests:
                    if 'range=' in cr['url'].lower():
                        range_base = cr['url']
                        log.info("  C0 discovered Range URL: %s", range_base[:200])
                        break

                if range_base and 'range=' in range_base.lower():
                    # We discovered the pattern — now use Strategy B with it
                    rm = re.search(r'[Rr]ange=(\d+)/(\d+)', range_base)
                    if rm:
                        page_size = int(rm.group(2))
                        # Extract page 2 data first
                        raw = page.evaluate(JS_EXTRACT_TABLE)
                        if raw:
                            existing_names = set(m["Name"] + "|" + m["Address"]
                                                 for m in all_members)
                            added = 0
                            for row in raw:
                                city, state = split_city_state(row.get("cityState", ""))
                                key = row.get("name", "") + "|" + row.get("address", "")
                                if key in existing_names:
                                    continue
                                existing_names.add(key)
                                all_members.append({
                                    "Name": row.get("name", ""),
                                    "Classification": row.get("classification", ""),
                                    "Address": row.get("address", ""),
                                    "City": city, "State": state,
                                    "Phone": row.get("phone", ""),
                                    "Web Address": "",
                                    "_detail_url": row.get("detailUrl", ""),
                                    "_orgcd": row.get("orgcd", ""),
                                })
                                added += 1
                            log.info("  Page 2: added %d members (total: %d)",
                                     added, len(all_members))

                        # Now paginate through remaining pages
                        for pg in range(3, 200):
                            start = (pg - 1) * page_size + 1
                            pg_url = re.sub(r'[Rr]ange=\d+/\d+',
                                            f'Range={start}/{page_size}', range_base)
                            try:
                                retry(lambda u=pg_url: page.goto(
                                    u, wait_until="domcontentloaded",
                                    timeout=PAGE_LOAD_TIMEOUT))
                                page.wait_for_timeout(1000)
                                raw = page.evaluate(JS_EXTRACT_TABLE)
                                if not raw:
                                    log.info("  No rows on page %d. Stopping.", pg)
                                    break
                                added = 0
                                for row in raw:
                                    city, state = split_city_state(
                                        row.get("cityState", ""))
                                    key = (row.get("name", "") + "|" +
                                           row.get("address", ""))
                                    if key in existing_names:
                                        continue
                                    existing_names.add(key)
                                    all_members.append({
                                        "Name": row.get("name", ""),
                                        "Classification": row.get("classification", ""),
                                        "Address": row.get("address", ""),
                                        "City": city, "State": state,
                                        "Phone": row.get("phone", ""),
                                        "Web Address": "",
                                        "_detail_url": row.get("detailUrl", ""),
                                        "_orgcd": row.get("orgcd", ""),
                                    })
                                    added += 1
                                if pg % 10 == 0:
                                    log.info("  Page %d: %d new (total: %d)",
                                             pg, added, len(all_members))
                                if added == 0:
                                    log.info("  No new rows on page %d. Stopping.", pg)
                                    break
                            except Exception as exc:
                                log.error("  Page %d failed: %s. Stopping.", pg, exc)
                                break

                        if len(all_members) > 30:
                            log.info("Strategy C0 SUCCESS: %d members", len(all_members))
                            return all_members
        except Exception as exc:
            log.warning("Strategy C0 failed: %s", exc)

    # ----- Strategy C: Playwright native click pagination -----
    log.info("Strategy C: Playwright click pagination")
    # Go back to page 1 results
    retry(lambda: page.goto(results_url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    # Collect the names from page 1 to detect if we're stuck
    page1_names = set(m["Name"] for m in all_members)
    consecutive_stuck = 0

    for page_num in range(2, 200):
        prev_count = len(all_members)

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
                        try:
                            page.wait_for_load_state("networkidle", timeout=15000)
                        except Exception:
                            pass  # Timeout is OK, page might have loaded
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

        # Verify page actually changed (not same data as page 1)
        new_names = set(r.get("name", "") for r in raw)
        if new_names == page1_names or (new_names and new_names.issubset(page1_names)):
            consecutive_stuck += 1
            if consecutive_stuck >= 2:
                log.info("  Page didn't change (same data as page 1). Stopping.")
                break
        else:
            consecutive_stuck = 0

        new_count = 0
        existing_names = set(m["Name"] + "|" + m["Address"] for m in all_members)
        for row in raw:
            city, state = split_city_state(row.get("cityState", ""))
            key = row.get("name", "") + "|" + row.get("address", "")
            if key in existing_names:
                continue
            existing_names.add(key)
            all_members.append({
                "Name": row.get("name", ""),
                "Classification": row.get("classification", ""),
                "Address": row.get("address", ""),
                "City": city, "State": state,
                "Phone": row.get("phone", ""),
                "Web Address": "",
                "_detail_url": row.get("detailUrl", ""),
                "_orgcd": row.get("orgcd", ""),
            })
            new_count += 1

        log.info("  Page %d: %d new rows (total: %d)", page_num, new_count, len(all_members))

        if new_count == 0:
            consecutive_stuck += 1
            if consecutive_stuck >= 2:
                log.info("  No new data for 2 consecutive pages. Stopping.")
                break

    log.info("Total members collected: %d", len(all_members))
    return all_members


# ---------------------------------------------------------------------------
# Helper: fix double-protocol URLs in Python too
# ---------------------------------------------------------------------------
def fix_url(url):
    """Fix double-protocol bug: http://https://example.com -> https://example.com"""
    if not url:
        return url
    url = url.strip()
    # Repeatedly strip leading http(s):// until only one remains
    while re.match(r'^https?://https?://', url, re.IGNORECASE):
        url = re.sub(r'^https?://', '', url, count=1)
    if not url.startswith('http'):
        url = 'https://' + url
    return url


# ---------------------------------------------------------------------------
# Phase 2: Web addresses from detail pages
# ---------------------------------------------------------------------------
DETAIL_BASE_URL = "https://members.glass.org/cvweb/cgi-bin/organizationdll.dll/Info"


def build_detail_url(orgcd):
    """Build a direct detail page URL from an orgcd value."""
    return f"{DETAIL_BASE_URL}?orgcd={orgcd}&wrp=organizationinfo.htm"


def discover_detail_url_template(page, members):
    """Click the first member name to discover the detail URL pattern.
    Returns the base URL template (with orgcd placeholder) or None."""
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
        log.info("Clicking first member to discover URL template: %s", first_member_name)
        member_link = page.locator(f"a:text-is('{first_member_name}')").first
        if member_link.count():
            member_link.click(timeout=10000)
            page.wait_for_load_state("networkidle", timeout=15000)
            page.wait_for_timeout(2000)
            detail_url = page.url
            log.info("Detail page URL after click: %s", detail_url)

            orgcd_match = re.search(r'orgcd=(\d+)', detail_url, re.IGNORECASE)
            if orgcd_match:
                # Build template by replacing the orgcd value with a placeholder
                template = re.sub(r'orgcd=\d+', 'orgcd={orgcd}', detail_url, flags=re.IGNORECASE)
                log.info("Discovered URL template: %s", template)
                return template
    except Exception as exc:
        log.warning("Could not discover detail URL template: %s", exc)
    return None


def collect_web_addresses(page, members):
    total = len(members)
    log.info("Phase 2: Fetching web addresses from %d detail pages …", total)

    # Debug: check what detail URLs and orgcds look like
    log.info("Sample member data (first 5):")
    for i, m in enumerate(members[:5]):
        log.info("  [%d] name=%s orgcd=%s detail_url=%s",
                 i, m.get("Name", ""), m.get("_orgcd", ""),
                 (m.get("_detail_url", "") or "")[:150])

    # Determine how to build detail URLs
    # Priority: use orgcd to build URLs directly
    has_orgcd = sum(1 for m in members if m.get("_orgcd"))
    log.info("Members with orgcd: %d / %d", has_orgcd, total)

    url_template = None
    if has_orgcd == 0:
        # No orgcds extracted from table — try clicking first member to discover pattern
        log.info("No orgcd values found. Trying to discover URL template by clicking.")
        url_template = discover_detail_url_template(page, members)

    # Now iterate through all members and fetch web addresses
    web_found = 0
    web_failed = 0
    for idx, member in enumerate(members):
        # Build the detail URL for this member
        orgcd = member.get("_orgcd", "")
        detail_url = ""

        if orgcd:
            # Best case: build URL directly from orgcd
            detail_url = build_detail_url(orgcd)
        elif url_template and orgcd:
            detail_url = url_template.format(orgcd=orgcd)
        else:
            # Fall back to stored detail URL if it's a real HTTP URL
            stored = member.get("_detail_url", "")
            if stored and stored.startswith("http") and not stored.startswith("javascript:"):
                detail_url = stored

        if not detail_url:
            web_failed += 1
            if idx < 5:
                log.info("  [%d] %s — no detail URL available, skipping", idx, member["Name"])
            continue

        try:
            retry(lambda u=detail_url: page.goto(u, wait_until="domcontentloaded",
                                                   timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(400)
            web = page.evaluate(JS_EXTRACT_WEB_ADDRESS)
            raw_web = (web or "").strip()
            # Apply double-protocol fix in Python as a safety net
            member["Web Address"] = fix_url(raw_web) if raw_web else ""
            if raw_web:
                web_found += 1

            # Save debug for first few
            if idx < 3:
                page.screenshot(path=os.path.join(OUTPUT_DIR, f"detail_page{idx+1}.png"),
                                full_page=True)
                with open(os.path.join(OUTPUT_DIR, f"detail_page{idx+1}.html"), "w",
                          encoding="utf-8") as f:
                    f.write(page.content())
                log.info("[%d] %s — url=%s — web: %s", idx, member["Name"],
                         detail_url[:100], member["Web Address"] or "(empty)")

        except Exception as exc:
            log.warning("[%d] Error for %s: %s", idx, member["Name"], exc)
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
