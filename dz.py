"""
dz.py — posts jobs listed in processed_jobs_algeria.csv to a Facebook Page.

The CSV is ALWAYS read from GitHub (raw URL below), never from a local copy,
so the poster sees whatever the scraper last pushed.

Every post is built from columns already in the CSV:
    Job Title, Company Name, Location, Short Description, Job Site URL
Only rows with Status == "posted" are used.
If "Job Site URL" is empty but the row has a WP ID and SITE_BASE_URL is set,
the link falls back to  SITE_BASE_URL/?p=<WP ID>  (WordPress redirects it to
the real permalink).

Env vars
    FB_PAGE_ID             numeric Page ID                      (secret)
    FB_PAGE_ACCESS_TOKEN   Page access token (not user token)   (secret)
    CSV_SOURCE             URL of the CSV   default: the projectfetcher/dz raw URL
    CSV_TOKEN              optional GitHub token (needed if the repo is private)
    SITE_BASE_URL          e.g. https://algeria.mimusjobs.com   (optional fallback)
    FB_MAX_POSTS_PER_RUN   default 15
    FB_MAX_AGE_DAYS        skip rows older than this, default 7

State: fb_posted_algeria.csv (Job ID, FB Post ID, Site Path, Timestamp) — remembers
what was already posted so nothing is posted twice. Commit it back to the repo.
"""
import csv
import io
import logging
import os
import re
import sys
import time
from collections import Counter
from datetime import datetime, timedelta
from urllib.parse import urlparse

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Config ───────────────────────────────────────────────────────────────────
DEFAULT_CSV_URL = "https://raw.githubusercontent.com/projectfetcher/dz/main/processed_jobs_algeria.csv"
CSV_SOURCE   = os.environ.get("CSV_SOURCE", "").strip() or DEFAULT_CSV_URL
CSV_TOKEN    = (os.environ.get("CSV_TOKEN", "") or os.environ.get("GITHUB_TOKEN", "")).strip()
SITE_BASE_URL = os.environ.get("SITE_BASE_URL", "").strip().rstrip("/")
STATE_FILE   = "fb_posted_algeria.csv"
STATE_COLS   = ["Job ID", "FB Post ID", "Site Path", "Timestamp"]

FB_PAGE_ID     = os.environ.get("FB_PAGE_ID", "").strip()
FB_PAGE_TOKEN  = os.environ.get("FB_PAGE_ACCESS_TOKEN", "").strip()
FB_API_VERSION = "v21.0"
FB_POST_DELAY_S      = 45
FB_MAX_POSTS_PER_RUN = int(os.environ.get("FB_MAX_POSTS_PER_RUN", "15") or 15)
FB_MAX_AGE_DAYS      = int(os.environ.get("FB_MAX_AGE_DAYS", "7") or 7)
SNIPPET_CHARS        = 220
HASHTAGS             = "#Algeria #Jobs #Hiring #Emploi"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("fb_poster")


# ── CSV loading (always from GitHub) ─────────────────────────────────────────
def _to_raw_github_url(url: str) -> str:
    """Accepts a github.com/.../blob/... link and converts it to the raw file URL."""
    m = re.match(r"https?://github\.com/([^/]+)/([^/]+)/blob/(.+)", url)
    return f"https://raw.githubusercontent.com/{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else url


def _read_source() -> str:
    url = _to_raw_github_url(CSV_SOURCE)
    if not url.lower().startswith("http"):
        raise ValueError(f"CSV_SOURCE must be a URL, got: {CSV_SOURCE!r}")

    attempts = []
    if CSV_TOKEN:
        attempts.append({"Authorization": f"token {CSV_TOKEN}"})
    attempts.append({})  # public repo / token rejected → try anonymously

    last_err = None
    for headers in attempts:
        try:
            r = requests.get(url, headers={**headers, "Cache-Control": "no-cache"}, timeout=30)
            r.raise_for_status()
            r.encoding = "utf-8"
            log.info(f"CSV fetched from GitHub ({len(r.text)} chars)")
            return r.text.lstrip("\ufeff")
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Could not download CSV from {url}: {last_err}")


def load_rows() -> list:
    text = _read_source()
    first_line = text.split("\n", 1)[0]
    delim = "\t" if first_line.count("\t") > first_line.count(",") else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    rows = []
    for raw in reader:
        rows.append({(k or "").strip(): (v or "").strip() for k, v in raw.items()})
    return rows


# ── State (what has already gone to Facebook) ────────────────────────────────
def load_state() -> tuple:
    ids, paths = set(), set()
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8", newline="") as f:
            for r in csv.DictReader(f):
                ids.add(r.get("Job ID", ""))
                paths.add(r.get("Site Path", ""))
    return ids, paths


def save_state(job_id: str, fb_id: str, site_path_: str):
    new_file = not os.path.exists(STATE_FILE)
    with open(STATE_FILE, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(STATE_COLS)
        w.writerow([job_id, fb_id, site_path_, datetime.now().isoformat()])


# ── Row filtering / text building ────────────────────────────────────────────
def job_link(row: dict) -> str:
    """Job Site URL from the CSV, or a ?p=<WP ID> fallback built from SITE_BASE_URL."""
    url = row.get("Job Site URL", "")
    if url.startswith("http"):
        return url
    wp_id = row.get("WP ID", "")
    if SITE_BASE_URL and wp_id:
        try:
            return f"{SITE_BASE_URL}/?p={int(float(wp_id))}"
        except ValueError:
            pass
    return ""


def site_path(url: str) -> str:
    """Path (+query) only — keeps the domain out of logs and the state file."""
    p = urlparse(url)
    path = p.path or url
    if p.query:
        path += "?" + p.query
    return path


def parse_ts(value: str):
    """Naive datetime, or None. Timezone info is stripped so comparisons never crash."""
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except Exception:
        return None


def skip_reason(row: dict, cutoff: datetime):
    """Returns None if the row is eligible, otherwise a short reason string."""
    if row.get("Status", "").lower() != "posted":
        return "status_not_posted"
    if not row.get("Job Title"):
        return "no_title"
    if not job_link(row):
        return "no_site_url"
    ts = parse_ts(row.get("Timestamp", ""))
    if ts and ts < cutoff:
        return "too_old"
    return None


def clean_location(text: str) -> str:
    """The site stores locations like 'سطيف سطيف الجزائر' — drop consecutive duplicate words."""
    out = []
    for tok in (text or "").split():
        if not out or out[-1] != tok:
            out.append(tok)
    return " ".join(out)


def short_description(text: str, limit: int = SNIPPET_CHARS) -> str:
    text = re.sub(r"\s+", " ", text or "").strip()
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(".,;:!?…")
    return cut + "…"


def build_message(row: dict, link: str) -> str:
    lines = [f"📢 {row['Job Title']}"]
    if row.get("Company Name"):
        lines.append(f"🏢 {row['Company Name']}")
    loc = clean_location(row.get("Location", ""))
    if loc:
        lines.append(f"📍 {loc}")
    snippet = short_description(row.get("Short Description", ""))
    if snippet:
        lines += ["", snippet]
    lines += ["", f"👉 Full details & how to apply: {link}", "", HASHTAGS]
    return "\n".join(lines)


# ── Facebook ─────────────────────────────────────────────────────────────────
def post_to_facebook(message: str, link: str) -> tuple:
    """Returns (fb_post_id | None, status) where status is 'ok' | 'retry' | 'rate_limit' | 'bad_token'."""
    endpoint = f"https://graph.facebook.com/{FB_API_VERSION}/{FB_PAGE_ID}/feed"
    payload = {"message": message, "link": link, "access_token": FB_PAGE_TOKEN}
    for attempt in range(3):
        try:
            r = requests.post(endpoint, data=payload, timeout=30)
            data = r.json() if r.content else {}
            if r.status_code == 200 and data.get("id"):
                return data["id"], "ok"
            err = data.get("error", {})
            code = err.get("code")
            log.error(f"Facebook error (attempt {attempt+1}) code={code} "
                      f"subcode={err.get('error_subcode')}: {err.get('message', r.text[:200])}")
            if code == 190:
                return None, "bad_token"
            if code in (4, 17, 32, 613):
                return None, "rate_limit"
        except Exception as e:
            log.error(f"Facebook request failed (attempt {attempt+1}): {e}")
        time.sleep(3 * 2 ** attempt)
    return None, "retry"


# ── Main ─────────────────────────────────────────────────────────────────────
def main() -> int:
    if not (FB_PAGE_ID and FB_PAGE_TOKEN):
        log.error("FB_PAGE_ID / FB_PAGE_ACCESS_TOKEN not set — nothing posted.")
        return 1

    try:
        rows = load_rows()
    except Exception as e:
        log.error(f"Could not read CSV from GitHub: {e}")
        return 1

    cutoff = datetime.now() - timedelta(days=FB_MAX_AGE_DAYS)
    done_ids, done_paths = load_state()

    todo, seen_paths = [], set()
    skipped = Counter()
    for r in sorted(rows, key=lambda r: r.get("Timestamp", "")):
        reason = skip_reason(r, cutoff)
        if reason:
            skipped[reason] += 1
            continue
        link = job_link(r)
        p = site_path(link)
        if r.get("Job ID") in done_ids or p in done_paths:
            skipped["already_posted"] += 1
            continue
        if p in seen_paths:
            skipped["duplicate_in_csv"] += 1
            continue
        seen_paths.add(p)
        todo.append((r, link, p))

    log.info(f"CSV rows: {len(rows)} | new jobs to post: {len(todo)} | cap this run: {FB_MAX_POSTS_PER_RUN}")
    if skipped:
        log.info("Skipped rows by reason: " + ", ".join(f"{k}={v}" for k, v in skipped.most_common()))

    posted = 0
    for r, link, path in todo:
        if posted >= FB_MAX_POSTS_PER_RUN:
            log.info("Per-run cap reached — the rest will go out next run.")
            break
        msg = build_message(r, link)

        fb_id, status = post_to_facebook(msg, link)
        if status == "ok":
            save_state(r.get("Job ID", ""), fb_id, path)
            posted += 1
            log.info(f"✅ posted '{r['Job Title']}' → {fb_id}  ({path})")
            time.sleep(FB_POST_DELAY_S)
        elif status == "bad_token":
            log.error("Facebook token invalid/expired — generate a new Page access token.")
            return 1
        elif status == "rate_limit":
            log.warning("Facebook rate limit hit — stopping; will resume next run.")
            break
        else:
            log.warning(f"Skipped '{r['Job Title']}' this run (will retry next run).")

    log.info(f"Done. Posted {posted} job(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
