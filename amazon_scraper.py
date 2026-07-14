# amazon_scraper.py
# Requirements:
# pip install undetected-chromedriver selenium beautifulsoup4 pandas openpyxl lxml
#
# GitHub Actions automation version.
# Reads:  queries.xlsx   (column named 'url' or 'Link')
# Writes: output.xlsx    (selling_price, stars, reviews, stock_message, status, error)
#
# NOTE: All scraping/extraction logic below is UNCHANGED from the original
# Streamlit app. Only the UI layer (Streamlit) has been removed and replaced
# with plain file I/O + print()-based progress logging, and the fixed
# `version_main=149` in build_driver() has been made auto-detecting (see
# comment there) since GitHub's Ubuntu runners install whatever the current
# stable Chrome release is, not version 149.

import os
import re
import sys
import time
import random
import atexit
from bs4 import BeautifulSoup
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ========== CONFIG ==========
HEADLESS = True               # Must stay True on GitHub Actions (no display)
PAGELOAD_TIMEOUT = 60
WAIT_BODY_TIMEOUT = 25
DELAY = (1.5, 3.0)

INPUT_FILE = "queries.xlsx"
OUTPUT_FILE = "output.xlsx"

# CHANGED FOR GITHUB ACTIONS (large lists / timeout safety):
# With hundreds of URLs, a single Actions run can hit the job's time limit
# before the loop finishes. Since the original script only wrote the output
# file once at the very end, a mid-run timeout meant losing all progress.
# Two things fix that:
#   1. SAVE_EVERY: write output.xlsx to disk every N rows, not just at the end.
#   2. MAX_RUNTIME_SECONDS: stop scraping (and do a final save) a bit before
#      the workflow step's own timeout, so the "commit and push" step always
#      has a valid, up-to-date output.xlsx to push, run after run.
# Rows already marked status == "OK" from a previous run are skipped, so the
# NEXT scheduled run automatically continues where this one left off instead
# of re-scraping everything from row 1.
SAVE_EVERY = 20
MAX_RUNTIME_SECONDS = 50 * 60  # 50 minutes; keep below the workflow step timeout

# Regex patterns
CURR_RX = re.compile(r"(?:₹|Rs\.?)\s*([\d,]+)")
STAR_RX = re.compile(r"([0-5](?:\.\d)?)")
COUNTS_RX = re.compile(r"([\d,]+)")
AVAIL_PATTERNS = re.compile(r"(out of stock|currently unavailable|unavailable|only \d+ left)", re.I)


# ========== HELPERS ==========
def normalize_amazon_url(u: str) -> str:
    """Normalize URL or ASIN -> full product URL."""
    if not u:
        return u
    u = u.strip()
    if u.startswith("http://"):
        u = "https://" + u[len("http://"):]
    # Accept ASIN bare string
    if re.fullmatch(r"[A-Z0-9]{10}", u, re.I):
        return f"https://www.amazon.in/dp/{u}"
    return u


def to_float_amt(txt):
    """Extract numeric rupee amount from text, return float or None."""
    if not txt:
        return None
    m = CURR_RX.search(txt)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except:
        return None


def build_driver():
    """Create undetected_chromedriver driver with sane options.
    PERF NOTE: only non-logic, performance-related options were added here:
    - page_load_strategy='eager' -> Selenium considers navigation 'done' as soon as
      DOM is ready (interactive), without waiting for all sub-resources (images, etc.)
      to finish loading. This does NOT change what we read from the DOM/HTML since
      we already explicitly wait for <body> presence (existing logic) before parsing.
    - prefs to disable image loading -> images are never used by the scraping logic
      (only text/HTML nodes are parsed), so skipping their download just saves
      bandwidth/time and does not affect extraction results.
    """
    opts = uc.ChromeOptions()
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--lang=en-US")
    opts.add_argument("--remote-allow-origins=*")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    if HEADLESS:
        # If your chromedriver does not support --headless=new, change to "--headless"
        try:
            opts.add_argument("--headless=new")
        except Exception:
            opts.add_argument("--headless")
    # --- Pure performance options (do not affect parsing/extraction logic) ---
    opts.page_load_strategy = "eager"
    opts.add_argument("--blink-settings=imagesEnabled=false")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-notifications")
    opts.add_experimental_option(
        "prefs",
        {
            "profile.managed_default_content_settings.images": 2,  # block images
            "profile.default_content_setting_values.notifications": 2,
        },
    )
    # CHANGED FOR GITHUB ACTIONS:
    # If the workflow set CHROME_PATH (browser-actions/setup-chrome does),
    # point Chrome at that exact binary instead of relying on PATH lookup.
    # Harmless / not used at all when running locally, where CHROME_PATH is
    # simply unset.
    chrome_path = os.environ.get("CHROME_PATH")
    if chrome_path:
        opts.binary_location = chrome_path
    # ---------------------------------------------------------------------
    # CHANGED FOR GITHUB ACTIONS:
    # The original code pinned version_main=149 to match a specific local
    # Chrome install. On GitHub's Ubuntu runner the installed Chrome version
    # will not be 149 (and will drift upward over time), so a hardcoded
    # version causes:
    #   SessionNotCreatedException: This version of ChromeDriver only
    #   supports Chrome version 149, current browser version is X
    # Leaving version_main unset lets undetected_chromedriver auto-detect
    # the installed Chrome's major version and fetch a matching driver.
    # If you ever need to pin a version again (e.g. to match a specific
    # Chrome installed by the workflow), pass version_main=<int> here.
    d = uc.Chrome(options=opts)
    d.set_page_load_timeout(PAGELOAD_TIMEOUT)
    atexit.register(lambda dd=d: safe_quit(dd))
    return d


def safe_quit(d):
    try:
        d.quit()
    except Exception:
        pass


def safe_get(driver, url, attempts=3, sleep=2.0):
    """Attempt driver.get with retries."""
    last = None
    for _ in range(attempts):
        try:
            driver.get(url)
            return
        except Exception as e:
            last = e
            time.sleep(sleep)
    raise last


# ========== EXTRACTORS (conservative, UNCHANGED) ==========
def extract_price(soup: BeautifulSoup):
    """
    Conservative price extraction:
    1) check known price blocks
    2) check meta tags
    3) as last resort, only accept body price if context contains MRP/Inclusive text
    """
    selectors = [
        "span#priceblock_dealprice",
        "span#priceblock_ourprice",
        "div#corePrice_feature_div span.a-price span.a-offscreen",
        "span.a-price > span.a-offscreen",
        "div#corePriceDisplay_desktop_feature_div span.a-offscreen",
    ]
    for sel in selectors:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(" ", strip=True)
            amt = to_float_amt(txt)
            if amt is not None:
                return amt
    # meta fallbacks
    meta_attrs = [
        ("meta", {"property": "product:price:amount"}),
        ("meta", {"name": "twitter:data1"}),
        ("meta", {"property": "og:price:amount"}),
        ("meta", {"itemprop": "price"}),
    ]
    for tag, attrs in meta_attrs:
        el = soup.find(tag, attrs=attrs)
        if el and el.get("content"):
            try:
                return float(str(el["content"]).replace(",", ""))
            except:
                pass
    # body fallback but only if context suggests price info
    body = soup.get_text(" ", strip=True)
    if ("M.R.P" in body or "MRP" in body or "Inclusive of all taxes" in body or "M.R.P." in body):
        m = CURR_RX.search(body)
        if m:
            try:
                return float(m.group(1).replace(",", ""))
            except:
                pass
    return None  # explicit None when price not determinable


def extract_rating_block(soup: BeautifulSoup):
    """
    Extract stars and reviews conservatively.
    Returns (stars or "N/A", reviews or "N/A")
    """
    stars = None
    reviews = None
    # Preferred star nodes
    el = soup.select_one("#averageCustomerReviews .a-icon-alt")
    if el:
        m = STAR_RX.search(el.get_text(" ", strip=True))
        if m:
            try:
                stars = float(m.group(1))
            except:
                stars = None
    if stars is None:
        acr = soup.select_one("#acrPopover")
        if acr:
            title = acr.get("title") or acr.get("data-a-popover")
            if title:
                m = STAR_RX.search(title)
                if m:
                    try:
                        stars = float(m.group(1))
                    except:
                        stars = None
    if stars is None:
        el2 = soup.select_one("span[data-hook='rating-out-of-text']")
        if el2:
            m = STAR_RX.search(el2.get_text(" ", strip=True))
            if m:
                try:
                    stars = float(m.group(1))
                except:
                    stars = None
    # Reviews count preferred
    acr_count = soup.select_one("#acrCustomerReviewText")
    if acr_count:
        txt = acr_count.get_text(" ", strip=True)
        m = COUNTS_RX.search(txt.replace(",", ""))
        if m:
            try:
                reviews = int(m.group(1))
            except:
                reviews = None
    # alternate attribute
    if reviews is None:
        alt = soup.select_one('span[data-hook="total-review-count"]')
        if alt:
            m = COUNTS_RX.search(alt.get_text(" ", strip=True).replace(",", ""))
            if m:
                try:
                    reviews = int(m.group(1))
                except:
                    reviews = None
    return (stars if stars is not None else "N/A", reviews if reviews is not None else "N/A")


def extract_stock_message(soup: BeautifulSoup):
    """
    Return short availability message or "N/A".
    Looks for buybox/availability and common phrases.
    """
    blocks = [
        "#availability", "#availability_feature_div", "div#outOfStock", "#outOfStock", "#sellerProfileTriggerId"
    ]
    for sel in blocks:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(" ", strip=True)
            if txt:
                if AVAIL_PATTERNS.search(txt) or "See all buying options" in txt or "Temporarily out of stock" in txt:
                    return txt.strip()
    if soup.find(string=re.compile(r"See all buying options", re.I)):
        return "See all buying options"
    m = re.search(r"Only\s+\d+\s+left", soup.get_text(" ", strip=True), re.I)
    if m:
        return m.group(0)
    return "N/A"


def _detect_link_col(df):
    for cand in ("url", "Url", "URL", "link", "Link", "LinkUrl", "amazon_link", "Link "):
        if cand in df.columns:
            return cand
    return None


def _load_working_dataframe(input_path: str, output_path: str):
    """
    Builds the dataframe to work on for this run.

    - Always reads queries.xlsx fresh, so it stays the source of truth for
      which rows/columns (e.g. 'Seller SKU') should exist and in what order.
    - If output.xlsx already exists (from a previous run that got cut off,
      or a previous scheduled run), carries forward any row whose status was
      already "OK" by matching on the normalized URL, so those rows are NOT
      re-scraped. Everything else (new rows, previously FAILed rows, blank
      rows) will be (re)scraped this run.
    """
    if input_path.lower().endswith(".csv"):
        df = pd.read_csv(input_path)
    else:
        df = pd.read_excel(input_path)

    link_col = _detect_link_col(df)
    if not link_col:
        print(f"ERROR: '{input_path}' must contain a 'url' or 'Link' column.")
        sys.exit(1)

    for c in ["selling_price", "stars", "reviews", "stock_message", "status", "error"]:
        if c not in df.columns:
            df[c] = None

    if os.path.exists(output_path):
        try:
            prev = pd.read_excel(output_path)
            prev_link_col = _detect_link_col(prev)
            if prev_link_col:
                done_rows = {}
                for _, row in prev.iterrows():
                    if str(row.get("status", "")).strip() == "OK":
                        key = normalize_amazon_url(str(row[prev_link_col]))
                        done_rows[key] = row
                if done_rows:
                    carried = 0
                    for idx, raw in df[link_col].astype(str).items():
                        key = normalize_amazon_url(raw)
                        if key in done_rows:
                            prev_row = done_rows[key]
                            for c in ["selling_price", "stars", "reviews", "stock_message", "status", "error"]:
                                if c in prev_row:
                                    df.loc[idx, c] = prev_row[c]
                            carried += 1
                    print(f"Resuming: {carried} row(s) already OK from a previous run, skipping those.")
        except Exception as e:
            print(f"Could not read previous {output_path}, starting fresh ({e}).")

    return df, link_col


# ========== SCRAPER (file-based, GitHub Actions entrypoint, resume-aware) ==========
def scrape_amazon_file(input_path: str, output_path: str):
    """
    Reads queries.xlsx/csv with 'url'/'Link' column, scrapes each product page
    not already marked "OK" in output.xlsx, and writes/updates
    selling_price, stars, reviews, stock_message, status, error.

    All original extraction/retry/delay logic is UNCHANGED from the
    Streamlit version. What's new here (for GitHub Actions + large lists):
      - resumes from output.xlsx instead of re-scraping everything each run
      - saves progress to disk every SAVE_EVERY rows
      - stops (with a final save) after MAX_RUNTIME_SECONDS so a run never
        gets killed mid-way with nothing written
    """
    df, link_col = _load_working_dataframe(input_path, output_path)

    total = len(df)
    pending_mask = df["status"].astype(str) != "OK"
    pending_count = int(pending_mask.sum())
    print(f"{total} total row(s), {pending_count} pending (not yet OK).")

    if pending_count == 0:
        print("Nothing left to scrape — all rows already OK.")
        return df

    # build driver
    d = build_driver()
    wait = WebDriverWait(d, WAIT_BODY_TIMEOUT)

    # warmup
    try:
        d.get("https://www.amazon.in/")
        time.sleep(1.0)
    except Exception:
        pass

    start_time = time.time()
    processed = 0
    try:
        for count, (idx, raw) in enumerate(df[link_col].astype(str).items(), start=1):
            if str(df.loc[idx, "status"]) == "OK":
                continue  # already scraped in a previous run

            if time.time() - start_time > MAX_RUNTIME_SECONDS:
                print("Time budget reached for this run — stopping early and saving progress. "
                      "Remaining rows will be picked up on the next scheduled run.")
                break

            url = normalize_amazon_url(raw)
            if not (url and url.startswith("http")):
                df.loc[idx, ["status", "error"]] = ["FAIL", "Bad/empty URL"]
                print(f"[{count}/{total}] SKIP (bad/empty URL)")
                continue
            try:
                safe_get(d, url, attempts=3, sleep=2.0)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                # small wait to let dynamic parts load
                time.sleep(1.0 + random.random() * 0.8)
                html = d.page_source
                soup = BeautifulSoup(html, "lxml")
                price = extract_price(soup)
                stars, reviews = extract_rating_block(soup)
                stock_msg = extract_stock_message(soup)
                # If price is None leave blank; stars/reviews may be "N/A"
                df.loc[idx, ["selling_price", "stars", "reviews", "stock_message", "status", "error"]] = \
                    [price, stars, reviews, stock_msg, "OK", None]
                print(f"[{count}/{total}] OK -> {url}")
            except Exception as e:
                # Save error message short
                err = str(e)[:400]
                df.loc[idx, ["status", "error"]] = ["FAIL", err]
                print(f"[{count}/{total}] FAIL -> {url} ({err})")

            processed += 1
            if processed % SAVE_EVERY == 0:
                df.to_excel(output_path, index=False)
                print(f"Progress saved to {output_path} ({processed} row(s) processed this run).")

            # polite delay between requests
            time.sleep(random.uniform(*DELAY))
    finally:
        safe_quit(d)
        time.sleep(0.2)

    return df


# ========== ENTRYPOINT (replaces Streamlit main()) ==========
def main():
    out = scrape_amazon_file(INPUT_FILE, OUTPUT_FILE)
    out.to_excel(OUTPUT_FILE, index=False)
    remaining = int((out["status"].astype(str) != "OK").sum())
    print(f"Done for this run. Wrote {len(out)} row(s) to {OUTPUT_FILE}. "
          f"{remaining} row(s) still pending" + (" — next scheduled run will continue." if remaining else "."))


if __name__ == "__main__":
    main()
