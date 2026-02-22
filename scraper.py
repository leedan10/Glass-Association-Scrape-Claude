#!/usr/bin/env python3
"""
NGA (National Glass Association) Member Directory Scraper

Scrapes member data from https://members.glass.org using Playwright
and exports to CSV and XLSX formats.

Platform: ClearVantage AMS by Euclid Technology (ISAPI DLL-based)
Key techniques:
  - page.expect_response() to capture XHR after javascript: link clicks
  - Network request interception to discover the Range URL pattern
  - Direct detail page URLs built from orgcd parameter
  - Multiple fallback strategies for pagination
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
DETAIL_BASE = "https://members.glass.org/cvweb/cgi-bin/organizationdll.dll/Info"

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
CSV_FILE = os.path.join(OUTPUT_DIR, "nga_members.csv")
XLSX_FILE = os.path.join(OUTPUT_DIR, "nga_members.xlsx")

FIELDS = [
    "Name", "Classification", "Address", "City", "State",
    "Country", "Phone", "Web Address",
]

DETAIL_DELAY = 0.3
PAGE_LOAD_TIMEOUT = 60_000
RETRY_COUNT = 3
RETRY_BACKOFF = 2

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s")
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
            log.warning("Attempt %d failed (%s). Retrying in %ss …",
                        attempt, exc, wait)
            time.sleep(wait)


def split_city_state(text):
    text = text.strip()
    if not text:
        return "", ""
    m = re.match(r"^(.+?),\s*([A-Z]{2})\b", text)
    return (m.group(1).strip(), m.group(2).strip()) if m else (text, "")


def fix_url(url):
    """Fix double-protocol bug: http://https://x.com -> https://x.com"""
    if not url:
        return url
    url = url.strip()
    while re.match(r'^https?://https?://', url, re.IGNORECASE):
        url = re.sub(r'^https?://', '', url, count=1)
    if url and not url.startswith('http'):
        url = 'https://' + url
    return url


def build_detail_url(orgcd):
    return f"{DETAIL_BASE}?orgcd={orgcd}&wrp=organizationinfo.htm"


def _row_to_member(row):
    """Convert a raw JS-extracted row dict into our standard member dict."""
    city, state = split_city_state(row.get("cityState", ""))
    return {
        "Name": row.get("name", ""),
        "Classification": row.get("classification", ""),
        "Address": row.get("address", ""),
        "City": city,
        "State": state,
        "Country": "",
        "Phone": row.get("phone", ""),
        "Web Address": "",
        "_detail_url": row.get("detailUrl", ""),
        "_orgcd": row.get("orgcd", ""),
    }


# ---------------------------------------------------------------------------
# JS snippets
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
            if (first.includes('\\u25bc') || first.includes('\\u25b2')) continue;
            if (/^\\d+$/.test(texts[0]) || texts[0].includes('MATCH')) continue;

            let detailUrl = '', detailOnclick = '', orgcd = '';
            const link = cells[0].querySelector('a');
            if (link) {
                detailUrl = link.getAttribute('href') || '';
                detailOnclick = link.getAttribute('onclick') || '';
                if (link.href && link.href.startsWith('http'))
                    detailUrl = link.href;
                const combined = detailOnclick + ' ' + detailUrl + ' ' + (link.href || '');
                const m = combined.match(/orgcd=(\\d+)/i);
                if (m) orgcd = m[1];
            }
            dataRows.push({
                name: texts[0], classification: texts[1] || '',
                address: texts[2] || '', cityState: texts[3] || '',
                phone: texts[4] || '',
                detailUrl, detailOnclick, orgcd
            });
        }
    }
    const seen = new Set();
    return dataRows.filter(r => {
        const key = r.name + '|' + r.address;
        if (seen.has(key) || !r.name) return false;
        seen.add(key); return true;
    });
}
"""

JS_DUMP_PAGINATION = """
() => {
    const info = [];
    for (const a of document.querySelectorAll('a')) {
        const text = (a.innerText || '').trim();
        if (!text) continue;
        if (/^\\d+$/.test(text) || /[\\u2192\\u203A\\u00BB>\\u25B6\\u27A1]/.test(text) ||
            text.toLowerCase() === 'next') {
            info.push({
                text, href: a.getAttribute('href') || '',
                onclick: a.getAttribute('onclick') || '',
                fullHref: a.href || ''
            });
        }
        if (info.length >= 20) break;
    }
    return info;
}
"""

JS_EXTRACT_DETAIL = """
() => {
    function fixUrl(url) {
        if (!url) return url;
        url = url.trim();
        while (/^https?:\\/\\/https?:\\/\\//i.test(url)) {
            url = url.replace(/^https?:\\/\\//i, '');
        }
        if (url && !url.startsWith('http')) url = 'https://' + url;
        return url;
    }

    const result = { webAddress: '', country: '' };
    const allCells = Array.from(document.querySelectorAll('td, th'));

    for (let i = 0; i < allCells.length; i++) {
        const raw = (allCells[i].innerText || '').trim();
        const label = raw.replace(/:/g, '').trim().toLowerCase();

        // --- Web Address ---
        if (['web address', 'website', 'web site', 'url', 'home page', 'web']
                .includes(label) && !result.webAddress) {
            for (let j = i + 1; j < Math.min(i + 4, allCells.length); j++) {
                const next = allCells[j];
                const link = next.querySelector('a[href]');
                if (link) {
                    const href = link.getAttribute('href') || '';
                    if (href && !href.includes('glass.org') &&
                        !href.startsWith('javascript:') && !href.startsWith('mailto:')) {
                        result.webAddress = fixUrl(href);
                        break;
                    }
                }
                const txt = (next.innerText || '').trim();
                if (txt && txt.includes('.') && !txt.includes(' ') &&
                    txt.length > 3 && !txt.includes('glass.org')) {
                    result.webAddress = fixUrl(txt);
                    break;
                }
            }
        }

        // --- Country ---
        if (label === 'country' && !result.country) {
            for (let j = i + 1; j < Math.min(i + 3, allCells.length); j++) {
                const txt = (allCells[j].innerText || '').trim();
                if (txt && txt.length > 1 && txt.length < 80) {
                    result.country = txt;
                    break;
                }
            }
        }
    }

    // Fallback: <a> whose text looks like a URL
    if (!result.webAddress) {
        for (const a of document.querySelectorAll('a[href]')) {
            const text = (a.innerText || '').trim().toLowerCase();
            const href = a.getAttribute('href') || '';
            if ((text.startsWith('www.') || /^https?:\\/\\//.test(text)) &&
                !href.includes('glass.org') && !href.startsWith('mailto:') &&
                !href.startsWith('javascript:')) {
                result.webAddress = fixUrl(href || text);
                break;
            }
        }
    }

    // Fallback: any external link
    if (!result.webAddress) {
        const skip = ['glass.org','google.com','facebook.com','twitter.com',
                      'linkedin.com','youtube.com','instagram.com','bing.com',
                      'javascript:','mailto:','cvweb'];
        for (const a of document.querySelectorAll('a[href]')) {
            const href = a.getAttribute('href') || '';
            if (!href || href.startsWith('javascript:') || href.startsWith('mailto:'))
                continue;
            const full = href.startsWith('http') ? href : 'http://' + href;
            if (skip.some(s => full.toLowerCase().includes(s))) continue;
            result.webAddress = fixUrl(href);
            break;
        }
    }

    return result;
}
"""


# ---------------------------------------------------------------------------
# Phase 1 — Collect all members across all pages
# ---------------------------------------------------------------------------
def collect_all_members(page):
    all_members = []
    seen_keys = set()
    captured_urls = []

    def on_request(request):
        url = request.url.lower()
        if any(k in url for k in ['customlist', 'range=', 'orgsearch',
                                   'organizationdll']):
            captured_urls.append(request.url)

    page.on('request', on_request)

    # ---- Navigate & submit search ----
    log.info("Navigating to: %s", SEARCH_URL)
    retry(lambda: page.goto(SEARCH_URL, wait_until="domcontentloaded",
                            timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)
    log.info("Search page title: %s", page.title())

    submitted = False
    for sel in ["input[type='submit']", "button[type='submit']",
                "input[value='Search']", "button:has-text('Search')",
                "input[value='Find']", "button:has-text('Find')",
                "a:has-text('Search')", "input[type='image']"]:
        try:
            btn = page.locator(sel).first
            if btn.count() and btn.is_visible():
                log.info("Clicking submit: %s", sel)
                btn.click(timeout=10000)
                page.wait_for_load_state("networkidle", timeout=20000)
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
                page.wait_for_load_state("networkidle", timeout=20000)
                page.wait_for_timeout(2000)
        except Exception:
            pass

    results_url = page.url
    log.info("Results page URL: %s", results_url)

    # ---- Debug: save page 1 ----
    page.screenshot(path=os.path.join(OUTPUT_DIR, "results_page1.png"),
                    full_page=True)
    with open(os.path.join(OUTPUT_DIR, "results_page1.html"), "w",
              encoding="utf-8") as f:
        f.write(page.content())

    # ---- Log captured requests ----
    log.info("Captured %d relevant network requests", len(captured_urls))
    for u in captured_urls[:10]:
        log.info("  %s", u[:200])

    # ---- Dump pagination info ----
    pag_info = page.evaluate(JS_DUMP_PAGINATION)
    log.info("Pagination elements found: %d", len(pag_info))
    for pi in pag_info:
        log.info("  text=%s href=%s onclick=%s",
                 pi['text'], pi['href'][:100], pi['onclick'][:100])

    # ---- Get match count from page text ----
    match_count = page.evaluate("""
        () => {
            const m = (document.body.innerText || '').match(/(\\d[\\d,]*)\\s*MATCH/i);
            return m ? parseInt(m[1].replace(/,/g,''), 10) : 0;
        }
    """)
    log.info("Match count on page: %d", match_count)

    # ---- Extract page 1 ----
    raw_rows = page.evaluate(JS_EXTRACT_TABLE)
    log.info("Page 1: extracted %d rows", len(raw_rows))
    if not raw_rows:
        log.error("No data on page 1!")
        return all_members

    if raw_rows:
        log.info("First row: name=%s orgcd=%s detailUrl=%s onclick=%s",
                 raw_rows[0].get('name', ''),
                 raw_rows[0].get('orgcd', ''),
                 raw_rows[0].get('detailUrl', '')[:100],
                 raw_rows[0].get('detailOnclick', '')[:100])
        orgcd_count = sum(1 for r in raw_rows if r.get('orgcd'))
        log.info("Rows with orgcd: %d / %d", orgcd_count, len(raw_rows))

    def add_rows(rows):
        added = 0
        for row in rows:
            key = row.get("name", "") + "|" + row.get("address", "")
            if key in seen_keys:
                continue
            seen_keys.add(key)
            all_members.append(_row_to_member(row))
            added += 1
        return added

    add_rows(raw_rows)
    page_size = len(raw_rows)

    # ==================================================================
    # PAGINATION — 4 strategies, stop at first success
    # ==================================================================

    # Find a Range URL from captured requests
    range_url = ""
    for u in captured_urls:
        if 'range=' in u.lower() and 'customlist' in u.lower():
            range_url = u
            break
    if not range_url:
        for u in captured_urls:
            if 'range=' in u.lower():
                range_url = u
                break
    log.info("Range URL from capture: %s",
             range_url[:200] if range_url else "(none)")

    # ---- Strategy A: Big Range (all-at-once) ----
    if range_url:
        big = re.sub(r'[Rr]ange=\d+/\d+', 'Range=1/5000', range_url)
        log.info("Strategy A: big Range %s", big[:200])
        try:
            page.goto(big, wait_until="domcontentloaded",
                      timeout=PAGE_LOAD_TIMEOUT)
            page.wait_for_timeout(5000)
            raw = page.evaluate(JS_EXTRACT_TABLE)
            log.info("  Returned %d rows", len(raw))
            if len(raw) > len(all_members):
                all_members.clear()
                seen_keys.clear()
                add_rows(raw)
                log.info("Strategy A SUCCESS: %d members", len(all_members))
                return all_members
        except Exception as exc:
            log.warning("Strategy A failed: %s", exc)

    # ---- Strategy B: Range pagination (page-by-page) ----
    range_base = range_url
    if not range_base:
        for pi in pag_info:
            for fld in ['fullHref', 'href']:
                u = pi.get(fld, '')
                if u and 'range=' in u.lower():
                    range_base = u
                    break
            if range_base:
                break

    if range_base and re.search(r'[Rr]ange=\d+/\d+', range_base):
        rm = re.search(r'[Rr]ange=(\d+)/(\d+)', range_base)
        ps = int(rm.group(2))
        total_pages = (match_count + ps - 1) // ps if match_count else 200
        log.info("Strategy B: Range pagination — size=%d, pages=%d",
                 ps, total_pages)
        for pg in range(2, total_pages + 1):
            start = (pg - 1) * ps + 1
            pg_url = re.sub(r'[Rr]ange=\d+/\d+',
                            f'Range={start}/{ps}', range_base)
            try:
                retry(lambda u=pg_url: page.goto(
                    u, wait_until="domcontentloaded",
                    timeout=PAGE_LOAD_TIMEOUT))
                page.wait_for_timeout(800)
                raw = page.evaluate(JS_EXTRACT_TABLE)
                if not raw:
                    log.info("  Page %d empty. Stopping.", pg)
                    break
                n = add_rows(raw)
                if pg % 10 == 0:
                    log.info("  Page %d/%d: +%d (total %d)",
                             pg, total_pages, n, len(all_members))
            except Exception as exc:
                log.error("  Page %d failed: %s. Stopping.", pg, exc)
                break
        if len(all_members) > page_size:
            log.info("Strategy B SUCCESS: %d members", len(all_members))
            return all_members

    # ---- Strategy C0: Click page 2 to discover Range URL ----
    log.info("Strategy C0: click page 2 with request interception")
    try:
        retry(lambda: page.goto(results_url, wait_until="domcontentloaded",
                                timeout=PAGE_LOAD_TIMEOUT))
        page.wait_for_timeout(3000)

        captured_urls.clear()
        page2_link = page.locator("a:text-is('2')").first
        if page2_link.count() and page2_link.is_visible():
            log.info("  Clicking page 2 …")
            page2_link.click(timeout=10000)
            try:
                page.wait_for_load_state("networkidle", timeout=20000)
            except Exception:
                pass
            page.wait_for_timeout(3000)

            for u in captured_urls:
                if 'range=' in u.lower():
                    range_base = u
                    log.info("  C0 discovered: %s", range_base[:200])
                    break

            raw = page.evaluate(JS_EXTRACT_TABLE)
            if raw:
                n = add_rows(raw)
                log.info("  Page 2: +%d rows (total %d)",
                         n, len(all_members))

            if range_base and re.search(r'[Rr]ange=\d+/\d+', range_base):
                rm = re.search(r'[Rr]ange=(\d+)/(\d+)', range_base)
                ps = int(rm.group(2))
                total_pages = ((match_count + ps - 1) // ps
                               if match_count else 200)
                for pg in range(3, total_pages + 1):
                    start = (pg - 1) * ps + 1
                    pg_url = re.sub(r'[Rr]ange=\d+/\d+',
                                    f'Range={start}/{ps}', range_base)
                    try:
                        retry(lambda u=pg_url: page.goto(
                            u, wait_until="domcontentloaded",
                            timeout=PAGE_LOAD_TIMEOUT))
                        page.wait_for_timeout(800)
                        raw = page.evaluate(JS_EXTRACT_TABLE)
                        if not raw:
                            log.info("  Page %d empty. Stopping.", pg)
                            break
                        n = add_rows(raw)
                        if n == 0:
                            log.info("  Page %d: no new rows. Stopping.", pg)
                            break
                        if pg % 10 == 0:
                            log.info("  Page %d: +%d (total %d)",
                                     pg, n, len(all_members))
                    except Exception as exc:
                        log.error("  Page %d failed: %s", pg, exc)
                        break
                if len(all_members) > page_size:
                    log.info("Strategy C0 SUCCESS: %d members",
                             len(all_members))
                    return all_members
    except Exception as exc:
        log.warning("Strategy C0 failed: %s", exc)

    # ---- Strategy C: Click-through pagination (arrow → then numbers) ----
    log.info("Strategy C: click-through pagination")
    retry(lambda: page.goto(results_url, wait_until="domcontentloaded",
                            timeout=PAGE_LOAD_TIMEOUT))
    page.wait_for_timeout(3000)

    consecutive_empty = 0
    for page_num in range(2, 200):
        clicked = False
        for pat in [str(page_num), "\u2192", "\u203a", "\u00bb",
                    ">>", "Next", "next", ">"]:
            try:
                links = page.locator(f"a:text-is('{pat}')").all()
                for lnk in links:
                    if lnk.is_visible():
                        log.info("  Clicking '%s' for page %d",
                                 pat, page_num)
                        lnk.click(timeout=10000)
                        try:
                            page.wait_for_load_state("networkidle",
                                                     timeout=20000)
                        except Exception:
                            pass
                        page.wait_for_timeout(2000)
                        clicked = True
                        break
                if clicked:
                    break
            except Exception:
                continue
        if not clicked:
            log.info("  No pagination link for page %d. Done.", page_num)
            break

        raw = page.evaluate(JS_EXTRACT_TABLE)
        if not raw:
            log.info("  Page %d: no data. Done.", page_num)
            break
        n = add_rows(raw)
        log.info("  Page %d: +%d (total %d)",
                 page_num, n, len(all_members))
        if n == 0:
            consecutive_empty += 1
            if consecutive_empty >= 2:
                log.info("  2 consecutive empty pages. Done.")
                break
        else:
            consecutive_empty = 0

    log.info("Total members collected: %d", len(all_members))
    return all_members


# ---------------------------------------------------------------------------
# Phase 2 — Visit detail pages for Web Address + Country
# ---------------------------------------------------------------------------
def collect_detail_info(page, members):
    total = len(members)
    log.info("Phase 2: fetching detail info for %d members …", total)

    has_orgcd = sum(1 for m in members if m.get("_orgcd"))
    log.info("Members with orgcd: %d / %d", has_orgcd, total)
    for i, m in enumerate(members[:5]):
        log.info("  [%d] %s orgcd=%s url=%s", i, m["Name"],
                 m.get("_orgcd", ""),
                 (m.get("_detail_url", "") or "")[:120])

    # If no orgcds found, try clicking first member to discover pattern
    url_template = None
    if has_orgcd == 0:
        log.info("No orgcds — clicking first member to discover URL pattern")
        try:
            retry(lambda: page.goto(SEARCH_URL,
                                    wait_until="domcontentloaded",
                                    timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(3000)
            for sel in ["input[type='submit']", "button[type='submit']",
                        "input[value='Search']",
                        "button:has-text('Search')",
                        "input[type='image']"]:
                try:
                    btn = page.locator(sel).first
                    if btn.count() and btn.is_visible():
                        btn.click(timeout=10000)
                        page.wait_for_load_state("networkidle",
                                                 timeout=20000)
                        page.wait_for_timeout(2000)
                        break
                except Exception:
                    continue

            name = members[0]["Name"]
            log.info("Clicking member: %s", name)
            link = page.locator(f"a:text-is('{name}')").first
            if link.count():
                link.click(timeout=10000)
                page.wait_for_load_state("networkidle", timeout=20000)
                page.wait_for_timeout(2000)
                url = page.url
                log.info("Detail URL: %s", url)
                m = re.search(r'orgcd=(\d+)', url, re.IGNORECASE)
                if m:
                    url_template = re.sub(
                        r'orgcd=\d+', 'orgcd={orgcd}',
                        url, flags=re.IGNORECASE)
                    log.info("Template: %s", url_template)
        except Exception as exc:
            log.warning("Template discovery failed: %s", exc)

    found = 0
    skipped = 0
    for idx, member in enumerate(members):
        orgcd = member.get("_orgcd", "")
        detail_url = ""

        if orgcd:
            detail_url = build_detail_url(orgcd)
        elif url_template and orgcd:
            detail_url = url_template.format(orgcd=orgcd)
        else:
            stored = member.get("_detail_url", "")
            if (stored and stored.startswith("http")
                    and "javascript:" not in stored):
                detail_url = stored

        if not detail_url:
            skipped += 1
            if idx < 5:
                log.info("  [%d] %s — no URL, skipping",
                         idx, member["Name"])
            continue

        try:
            retry(lambda u=detail_url: page.goto(
                u, wait_until="domcontentloaded",
                timeout=PAGE_LOAD_TIMEOUT))
            page.wait_for_timeout(350)

            info = page.evaluate(JS_EXTRACT_DETAIL)
            raw_web = (info.get("webAddress") or "").strip()
            country = (info.get("country") or "").strip()

            member["Web Address"] = fix_url(raw_web) if raw_web else ""
            member["Country"] = country
            if raw_web:
                found += 1

            if idx < 3:
                page.screenshot(
                    path=os.path.join(OUTPUT_DIR,
                                      f"detail_{idx+1}.png"),
                    full_page=True)
                with open(os.path.join(OUTPUT_DIR,
                                       f"detail_{idx+1}.html"),
                          "w", encoding="utf-8") as f:
                    f.write(page.content())
                log.info("  [%d] %s — web=%s country=%s",
                         idx, member["Name"],
                         member["Web Address"] or "(none)",
                         country or "(none)")

        except Exception as exc:
            log.warning("  [%d] %s error: %s", idx, member["Name"], exc)
            skipped += 1

        if idx > 0 and idx % 100 == 0:
            log.info("  Progress: %d/%d — found %d web addresses",
                     idx, total, found)

        time.sleep(DETAIL_DELAY)

    log.info("Detail scrape done — web found: %d, skipped: %d",
             found, skipped)
    return found


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
        max_len = max((len(str(cell.value or "")) for cell in col),
                      default=10)
        ws.column_dimensions[col[0].column_letter].width = min(
            max_len + 2, 50)
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
            args=["--disable-blink-features=AutomationControlled",
                  "--no-sandbox"],
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
            "Object.defineProperty(navigator, 'webdriver', "
            "{get: () => false})"
        )
        page = context.new_page()

        # Phase 1: collect all members from list pages
        members = collect_all_members(page)
        if not members:
            log.error("No members found.")
            browser.close()
            return

        # Phase 2: visit detail pages for web address + country
        collect_detail_info(page, members)

        browser.close()

    # Phase 3: export
    export_csv(members)
    export_xlsx(members)

    log.info("=" * 60)
    log.info("DONE — Total members: %d", len(members))
    log.info("=" * 60)


if __name__ == "__main__":
    main()
