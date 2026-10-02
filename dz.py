"""
fb_poster.py — posts jobs listed in processed_jobs_algeria.csv to a Facebook Page.

NO scraping. Every post is built from columns already in the CSV:
    Job Title, Company Name, Location, Short Description, Job Site URL
Only rows with Status == "posted" AND a Job Site URL are used.

Env vars
    FB_PAGE_ID             numeric Page ID                      (secret)
    FB_PAGE_ACCESS_TOKEN   Page access token (not user token)   (secret)
    CSV_SOURCE             local path or URL of the CSV         default: processed_jobs_algeria.csv
    CSV_TOKEN              optional GitHub token (private repo, when CSV_SOURCE is a URL)
    FB_MAX_POSTS_PER_RUN   default 15
    FB_MAX_AGE_DAYS        skip rows older than this, default 7
    FB_DRY_RUN             "1" = print posts, don't send, don't save state

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
from datetime import datetime, timedelta
from urllib.parse import urlparse

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Config ───────────────────────────────────────────────────────────────────
CSV_SOURCE   = os.environ.get("CSV_SOURCE", "processed_jobs_algeria.csv").strip()
CSV_TOKEN    = os.environ.get("CSV_TOKEN", "").strip()
STATE_FILE   = "fb_posted_algeria.csv"
STATE_COLS   = ["Job ID", "FB Post ID", "Site Path", "Timestamp"]

FB_PAGE_ID     = os.environ.get("FB_PAGE_ID", "").strip()
FB_PAGE_TOKEN  = os.environ.get("FB_PAGE_ACCESS_TOKEN", "").strip()
FB_API_VERSION = "v21.0"
FB_POST_DELAY_S      = 45
FB_MAX_POSTS_PER_RUN = int(os.environ.get("FB_MAX_POSTS_PER_RUN", "15"))
FB_MAX_AGE_DAYS      = int(os.environ.get("FB_MAX_AGE_DAYS", "7"))
FB_DRY_RUN           = os.environ.get("FB_DRY_RUN", "").strip() == "1"
SNIPPET_CHARS        = 220
HASHTAGS             = "#Algeria #Jobs #Hiring #Emploi"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("fb_poster")


# ── CSV loading ──────────────────────────────────────────────────────────────
def _to_raw_github_url(url: str) -> str:
    """Accepts a github.com/.../blob/... link and converts it to the raw file URL."""
    m = re.match(r"https?://github\.com/([^/]+)/([^/]+)/blob/(.+)", url)
    return f"https://raw.githubusercontent.com/{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else url


def _read_source() -> str:
    if CSV_SOURCE.lower().startswith("http"):
        headers = {"Authorization": f"token {CSV_TOKEN}"} if CSV_TOKEN else {}
        r = requests.get(_to_raw_github_url(CSV_SOURCE), headers=headers, timeout=30)
        r.raise_for_status()
        r.encoding = "utf-8"
        return r.text.lstrip("\ufeff")
    with open(CSV_SOURCE, encoding="utf-8-sig", newline="") as f:
        return f.read()


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


def save_state(job_id: str, fb_id: str, site_path: str):
    new_file = not os.path.exists(STATE_FILE)
    with open(STATE_FILE, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(STATE_COLS)
        w.writerow([job_id, fb_id, site_path, datetime.now().isoformat()])


# ── Row filtering / text building ────────────────────────────────────────────
def site_path(url: str) -> str:
    """Path only (e.g. /job/barista/) — keeps the domain out of logs and the state file."""
    return urlparse(url).path or url


def parse_ts(value: str):
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def is_eligible(row: dict, cutoff: datetime) -> bool:
    if row.get("Status", "").lower() != "posted":
        return False
    if not row.get("Job Title") or not row.get("Job Site URL", "").startswith("http"):
        return False
    ts = parse_ts(row.get("Timestamp", ""))
    if ts and ts < cutoff:
        return False
    return True


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


def build_message(row: dict) -> str:
    lines = [f"📢 {row['Job Title']}"]
    if row.get("Company Name"):
        lines.append(f"🏢 {row['Company Name']}")
    loc = clean_location(row.get("Location", ""))
    if loc:
        lines.append(f"📍 {loc}")
    snippet = short_description(row.get("Short Description", ""))
    if snippet:
        lines += ["", snippet]
    lines += ["", f"👉 Full details & how to apply: {row['Job Site URL']}", "", HASHTAGS]
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
            log.error(f"Facebook error (attempt {attempt+1}): {err.get('message', r.text[:200])}")
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
    if not FB_DRY_RUN and not (FB_PAGE_ID and FB_PAGE_TOKEN):
        log.error("FB_PAGE_ID / FB_PAGE_ACCESS_TOKEN not set — nothing posted.")
        return 1

    try:
        rows = load_rows()
    except Exception as e:
        log.error(f"Could not read CSV source: {e}")
        return 1

    cutoff = datetime.now() - timedelta(days=FB_MAX_AGE_DAYS)
    done_ids, done_paths = load_state()

    todo, seen_paths = [], set()
    for r in sorted(rows, key=lambda r: r.get("Timestamp", "")):
        if not is_eligible(r, cutoff):
            continue
        p = site_path(r["Job Site URL"])
        if r.get("Job ID") in done_ids or p in done_paths or p in seen_paths:
            continue
        seen_paths.add(p)
        todo.append(r)

    log.info(f"CSV rows: {len(rows)} | new jobs to post: {len(todo)} "
             f"| cap this run: {FB_MAX_POSTS_PER_RUN}{' | DRY RUN' if FB_DRY_RUN else ''}")

    posted = 0
    for r in todo:
        if posted >= FB_MAX_POSTS_PER_RUN:
            log.info("Per-run cap reached — the rest will go out next run.")
            break
        msg, path = build_message(r), site_path(r["Job Site URL"])

        if FB_DRY_RUN:
            print("\n" + "─" * 60 + f"\n{msg}\n")
            posted += 1
            continue

        fb_id, status = post_to_facebook(msg, r["Job Site URL"])
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
