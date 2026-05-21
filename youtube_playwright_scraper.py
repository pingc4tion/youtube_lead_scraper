from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError
import pandas as pd
import time
from datetime import datetime
from urllib.parse import quote_plus
import re
import sqlite3

OUTPUT_FILE = "youtube_leads_playwright.csv"
DB_FILE = "youtube_leads.db"
SHEET_FILE = "2026 Cold Email Outreach Sheet.xlsx"
SHEET_NAME = "Cold Email"
SHEET_HEADER_ROW = 11
NAV_TIMEOUT_MS = 20000
CHANNELS_PER_KEYWORD_LIMIT = 12
TOTAL_LEADS_LIMIT = 30
MAX_SUBSCRIBERS = 1_000

keywords = [
    "how i quit my job",
    "side hustle to full time",
    "behind the scenes of my business",
    "starting a business from scratch",
    "my entrepreneur journey"
]


def normalize_channel_url(url):
    if not url:
        return ""
    u = str(url).strip().lower()
    u = re.sub(r"^https?://", "", u)
    u = re.sub(r"^www\.", "", u)
    return u.rstrip("/")


def load_existing_urls():
    urls = set()
    try:
        df = pd.read_excel(SHEET_FILE, sheet_name=SHEET_NAME, header=SHEET_HEADER_ROW)
    except Exception as e:
        print(f"Could not read sheet for dedupe: {e}")
        return urls

    if "YouTube Link" not in df.columns:
        print("YouTube Link column not found in sheet")
        return urls

    for raw in df["YouTube Link"].dropna().tolist():
        norm = normalize_channel_url(raw)
        if norm:
            urls.add(norm)
    print(f"Loaded {len(urls)} existing leads from sheet for dedupe")
    return urls


def init_db():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            started_at TEXT NOT NULL,
            finished_at TEXT,
            keywords TEXT,
            lead_count INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            keyword TEXT,
            channel_name TEXT,
            channel_url TEXT,
            channel_url_normalized TEXT,
            raw_text TEXT,
            scraped_at TEXT NOT NULL,
            FOREIGN KEY (run_id) REFERENCES runs(id)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_leads_url ON leads(channel_url_normalized)")
    conn.commit()
    return conn


def start_run(conn, keywords_list):
    cur = conn.execute(
        "INSERT INTO runs (started_at, keywords, lead_count) VALUES (?, ?, 0)",
        (datetime.now().isoformat(timespec="seconds"), ",".join(keywords_list)),
    )
    conn.commit()
    return cur.lastrowid


def finish_run(conn, run_id, leads):
    now = datetime.now().isoformat(timespec="seconds")
    conn.executemany(
        """INSERT INTO leads
           (run_id, keyword, channel_name, channel_url, channel_url_normalized, raw_text, scraped_at)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        [
            (run_id, r["keyword"], r["channel_name"], r["channel_url"],
             normalize_channel_url(r["channel_url"]), r["raw_text"], now)
            for r in leads
        ],
    )
    conn.execute(
        "UPDATE runs SET finished_at = ?, lead_count = ? WHERE id = ?",
        (now, len(leads), run_id),
    )
    conn.commit()


def load_existing_urls_from_db(conn):
    rows = conn.execute("SELECT channel_url_normalized FROM leads").fetchall()
    return {r[0] for r in rows if r[0]}


all_data = []
existing_urls = set()


def parse_relative_age_days(text):
    if not text:
        return None

    clean_text = text.lower().strip()
    match = re.search(r"(\d+)\s+(hour|day|week|month|year)s?\s+ago", clean_text)
    if not match:
        if re.search(r"\b(minute|second)s?\s+ago\b", clean_text):
            return 0
        return None

    value = int(match.group(1))
    unit = match.group(2)

    if unit == "hour":
        return 0
    if unit == "day":
        return value
    if unit == "week":
        return value * 7
    if unit == "month":
        return value * 30
    if unit == "year":
        return value * 365
    return None


VIDEO_ITEM_SELECTOR = "ytd-rich-item-renderer, ytd-rich-grid-media, ytd-grid-video-renderer"
VIDEO_META_SELECTOR = (
    # legacy DOM (old A/B variant)
    "ytd-rich-grid-media #metadata-line span, "
    "ytd-rich-item-renderer #metadata-line span, "
    "ytd-grid-video-renderer #metadata-line span, "
    # current DOM: ytd-video-meta-block + inline-metadata-item class
    "ytd-rich-item-renderer span.inline-metadata-item, "
    "ytd-rich-grid-media span.inline-metadata-item, "
    "ytd-grid-video-renderer span.inline-metadata-item"
)


def _wait_for_grid(page):
    try:
        page.wait_for_selector(VIDEO_ITEM_SELECTOR, timeout=5000)
    except PWTimeoutError:
        pass


def _wait_for_video_metadata(page):
    """Wait until at least one metadata span contains 'ago' text."""
    try:
        page.wait_for_function(
            """() => {
                const spans = document.querySelectorAll(
                    'ytd-rich-item-renderer #metadata-line span, '
                    + 'ytd-rich-grid-media #metadata-line span, '
                    + 'ytd-grid-video-renderer #metadata-line span, '
                    + 'ytd-rich-item-renderer span.inline-metadata-item, '
                    + 'ytd-rich-grid-media span.inline-metadata-item, '
                    + 'ytd-grid-video-renderer span.inline-metadata-item'
                );
                for (const s of spans) {
                    if (s.textContent && /\\bago\\b/i.test(s.textContent)) return true;
                }
                return false;
            }""",
            timeout=10000,
        )
    except PWTimeoutError:
        pass
    # extra settle in case items render but spans are still hydrating
    try:
        page.wait_for_load_state("networkidle", timeout=3000)
    except PWTimeoutError:
        pass


def parse_subscriber_count(text):
    """Parse '12.5K subscribers' / '1.2M subscribers' / '900 subscribers' → int. Returns None if absent."""
    if not text:
        return None
    match = re.search(
        r"([\d.,]+)\s*([KMB])?\s*subscriber",
        text,
        re.IGNORECASE,
    )
    if not match:
        return None
    num_str = match.group(1).replace(",", "")
    try:
        num = float(num_str)
    except ValueError:
        return None
    suffix = (match.group(2) or "").upper()
    multiplier = {"K": 1_000, "M": 1_000_000, "B": 1_000_000_000}.get(suffix, 1)
    return int(num * multiplier)


def about_page_state(page, channel_url):
    """Returns (is_india, subscriber_count_or_None)."""
    page.goto(f"{channel_url}/about", wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    try:
        page.wait_for_selector("ytd-about-channel-renderer, #additional-info-container", timeout=5000)
    except PWTimeoutError:
        pass

    about_text = page.locator("body").inner_text(timeout=5000)
    is_india = bool(re.search(r"country[:\s]*\n?\s*india\b", about_text, re.IGNORECASE))
    sub_count = parse_subscriber_count(about_text)
    return is_india, sub_count


def videos_tab_state(page, channel_url):
    """One nav, returns (regular_video_count, latest_age_days_or_None)."""
    page.goto(f"{channel_url}/videos", wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    _wait_for_grid(page)

    regular_video_count = page.locator(VIDEO_ITEM_SELECTOR).count()
    if regular_video_count == 0:
        return 0, None

    _wait_for_video_metadata(page)
    # nudge to trigger lazy hydration of metadata
    try:
        page.mouse.wheel(0, 800)
        time.sleep(0.6)
    except Exception:
        pass

    age_days = None
    meta_spans = page.locator(VIDEO_META_SELECTOR)
    span_count = min(meta_spans.count(), 20)
    for i in range(span_count):
        try:
            text = meta_spans.nth(i).inner_text(timeout=2000)
        except PWTimeoutError:
            continue
        parsed = parse_relative_age_days(text)
        if parsed is not None:
            age_days = parsed
            break

    # fallback: scan body text if span-based read missed it
    if age_days is None:
        try:
            body_text = page.locator("body").inner_text(timeout=3000)
            parsed = parse_relative_age_days(body_text)
            if parsed is not None:
                age_days = parsed
        except PWTimeoutError:
            pass

    return regular_video_count, age_days


def has_shorts(page, channel_url):
    page.goto(f"{channel_url}/shorts", wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    _wait_for_grid(page)
    return page.locator(VIDEO_ITEM_SELECTOR).count() > 0


def passes_channel_filters(page, channel_url):
    try:
        is_india, sub_count = about_page_state(page, channel_url)
        if is_india:
            return False, "country=India"

        if sub_count is None:
            return False, "subscriber_count_unknown"

        if sub_count > MAX_SUBSCRIBERS:
            return False, f"subscribers={sub_count}>max={MAX_SUBSCRIBERS}"

        regular_count, age_days = videos_tab_state(page, channel_url)

        if regular_count == 0:
            if has_shorts(page, channel_url):
                return False, "shorts_only"
            return False, "no_videos"

        if age_days is None:
            return False, "latest_upload_unknown"

        if age_days >= 365:
            return False, "latest_upload_older_than_1_year"

        return True, ""
    except Exception as err:
        return False, f"filter_error: {err}"

def scrape_keyword(page, keyword):
    print(f"Scraping: {keyword}")

    url = f"https://www.youtube.com/results?search_query={quote_plus(keyword)}&sp=EgIQAg%253D%253D"
    page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
    try:
        page.wait_for_selector("ytd-channel-renderer", timeout=8000)
    except PWTimeoutError:
        print(f"No channel results loaded for {keyword!r}")
        return

    # Scroll until we have enough channel candidates or stop growing.
    channels = page.locator("ytd-channel-renderer")
    last_count = -1
    for _ in range(8):
        current = channels.count()
        if current >= CHANNELS_PER_KEYWORD_LIMIT * 2 or current == last_count:
            break
        last_count = current
        page.mouse.wheel(0, 3000)
        time.sleep(0.8)

    channel_count = channels.count()
    candidates = []

    for idx in range(channel_count):
        try:
            channel = channels.nth(idx)
            name = channel.locator("#channel-title").inner_text(timeout=3000).strip()
            href = channel.locator("#main-link").get_attribute("href")

            if not href:
                continue

            channel_url = "https://www.youtube.com" + href

            if normalize_channel_url(channel_url) in existing_urls:
                continue

            candidates.append({
                "name": name,
                "channel_url": channel_url
            })

        except Exception as e:
            print("Skipped one result:", e)

    for candidate in candidates[:CHANNELS_PER_KEYWORD_LIMIT]:
        if len(all_data) >= TOTAL_LEADS_LIMIT:
            return
        try:
            channel_url = candidate["channel_url"]
            passed, reason = passes_channel_filters(page, channel_url)
            if not passed:
                print(f"Skipped {channel_url}: {reason}")
                continue

            page.goto(channel_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            time.sleep(0.6)
            meta_text = page.locator("body").inner_text(timeout=5000)

            all_data.append({
                "keyword": keyword,
                "channel_name": candidate["name"],
                "channel_url": channel_url,
                "raw_text": meta_text,
            })

            existing_urls.add(normalize_channel_url(channel_url))

        except Exception as e:
            print("Skipped one result:", e)

def main():
    global existing_urls

    existing_urls |= load_existing_urls()
    db_conn = init_db()
    existing_urls |= load_existing_urls_from_db(db_conn)
    print(f"Total dedupe set size (sheet + db): {len(existing_urls)}")
    run_id = start_run(db_conn, keywords)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()

            for keyword in keywords:
                if len(all_data) >= TOTAL_LEADS_LIMIT:
                    break
                scrape_keyword(page, keyword)

            browser.close()
    finally:
        finish_run(db_conn, run_id, all_data)
        db_conn.close()
        print(f"Run {run_id} saved to {DB_FILE} ({len(all_data)} leads)")

    df = pd.DataFrame(all_data)

    if df.empty:
        print("No new leads found.")
        return

    try:
        df.to_csv(OUTPUT_FILE, index=False)
        saved_path = OUTPUT_FILE
    except PermissionError:
        fallback_file = f"youtube_leads_playwright_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(fallback_file, index=False)
        saved_path = fallback_file
        print(f"{OUTPUT_FILE} is currently in use. Saved to fallback file instead: {fallback_file}")

    try:
        print(df)
    except UnicodeEncodeError:
        print(df.to_string().encode("ascii", errors="replace").decode("ascii"))
    print(f"Saved to {saved_path}")


if __name__ == "__main__":
    main()