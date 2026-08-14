"""
Daily Job Search Automation
----------------------------
Scrapes career pages listed in companies_config.json, keeps jobs posted
within the last N days that match your role keywords, and writes/updates
an Excel sheet with the results. Safe to re-run daily — it dedupes by
job link so you won't get repeat rows.

Usage:
    python main.py
"""

import json
import re
import time
import uuid
import logging
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "companies_config.json"
OUTPUT_FILE = BASE_DIR / "job_search_results.xlsx"
LOG_FILE = BASE_DIR / "scraper.log"

DAYS_THRESHOLD = 15
MAX_EXPERIENCE_YEARS = 4

# Edit this list to match the roles you're targeting.
KEYWORDS = [
    "full stack", "fullstack", "full-stack",
    "backend", "back end", "back-end",
    "software engineer", "software developer",
    "java developer", "spring boot", "java",
    "react", "reactjs", "sde",
]

# Matched against each job's location field. Add/remove cities as needed.
INDIA_LOCATION_KEYWORDS = [
    "india", "bengaluru", "bangalore", "mumbai", "pune", "hyderabad",
    "chennai", "delhi", "ncr", "gurgaon", "gurugram", "noida", "kolkata",
    "ahmedabad", "kochi", "coimbatore",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36"
}

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger().addHandler(console)
log = logging.getLogger(__name__)


# ----------------------------------------------------------------------
# ATS DETECTION
# ----------------------------------------------------------------------
def strip_html(html: str) -> str:
    if not html:
        return ""
    try:
        return BeautifulSoup(html, "html.parser").get_text(separator=" ", strip=True)
    except Exception:
        return html


def detect_ats(url: str) -> str:
    host = urlparse(url).netloc.lower()
    path = urlparse(url).path.lower()
    if "greenhouse.io" in host:
        return "greenhouse"
    if "lever.co" in host:
        return "lever"
    if "ashbyhq.com" in host:
        return "ashby"
    if "myworkdayjobs.com" in host:
        return "workday"
    if "smartrecruiters.com" in host:
        return "smartrecruiters"
    if "oraclecloud.com" in host and ("candidateexperience" in path or "/sites/" in path):
        return "oraclecloud"
    if re.search(r"/sites/[^/]+/jobs", path):
        # Oracle Fusion Recruiting's Candidate Experience URL shape, seen even
        # on custom/vanity domains (e.g. careers.oracle.com) that proxy a
        # backend *.oraclecloud.com tenant rather than expose it in the URL.
        return "oraclecloud"
    return "generic"


def extract_slug(url: str) -> str:
    """Pulls the company token out of a career page URL, e.g.
    https://boards.greenhouse.io/acme -> 'acme'
    https://jobs.lever.co/acme        -> 'acme'
    """
    path = urlparse(url).path.strip("/")
    return path.split("/")[0] if path else ""


# ----------------------------------------------------------------------
# FETCHERS — one per ATS type. Each returns a list of
# {title, link, posted_date (datetime), source}
# ----------------------------------------------------------------------
def fetch_greenhouse(company, slug):
    url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true"
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json().get("jobs", []):
        try:
            posted = datetime.fromisoformat(j["updated_at"].replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            posted = None
        jobs.append({
            "title": j.get("title", ""),
            "link": j.get("absolute_url", ""),
            "posted_date": posted,
            "source": company,
            "location": (j.get("location") or {}).get("name", ""),
            "description": strip_html(j.get("content", "")),
        })
    return jobs


def fetch_lever(company, slug):
    url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json():
        ts = j.get("createdAt")
        posted = datetime.fromtimestamp(ts / 1000) if ts else None
        categories = j.get("categories") or {}
        description = j.get("descriptionPlain") or strip_html(j.get("description", ""))
        lists_text = " ".join(
            strip_html(section.get("content", "")) for section in (j.get("lists") or [])
        )
        jobs.append({
            "title": j.get("text", ""),
            "link": j.get("hostedUrl", ""),
            "posted_date": posted,
            "source": company,
            "location": categories.get("location", ""),
            "description": strip_html(description + " " + lists_text),
        })
    return jobs


def fetch_ashby(company, slug):
    url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    jobs = []
    for j in r.json().get("jobs", []):
        try:
            posted = datetime.fromisoformat(j["publishedAt"].replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            posted = None
        jobs.append({
            "title": j.get("title", ""),
            "link": j.get("jobUrl", ""),
            "posted_date": posted,
            "source": company,
            "location": j.get("location", ""),
            "description": strip_html(j.get("descriptionHtml", "")),  # often blank in list view
        })
    return jobs


def resolve_oracle_site_number(careers_url: str) -> str:
    """A /sites/<segment> path that starts with 'CX' IS the site number.
    Otherwise it's a vanity slug and the real value has to be pulled out
    of the careers page HTML.
    """
    path = urlparse(careers_url).path
    match = re.search(r"/sites/([^/?#]+)", path, re.IGNORECASE)
    if match and match.group(1).upper().startswith("CX"):
        return match.group(1)

    r = requests.get(careers_url, headers=HEADERS, timeout=20)
    if "ORA_CX_SITE_NUMBER" in r.cookies:
        return r.cookies["ORA_CX_SITE_NUMBER"]
    html = r.text
    for pattern in (
        r"""siteNumber\s*[:=]\s*['"]([^'"]+)['"]""",
        r'"siteNumber"\s*:\s*"([^"]+)"',
        r"siteNumber=(CX_[^&\"'\s]+)",
    ):
        found = re.search(pattern, html, re.IGNORECASE)
        if found:
            return found.group(1)
    return "CX_1"  # common default for production tenants


def fetch_oracle_cloud(company, careers_url, api_domain=None):
    """Oracle Cloud HCM (Fusion Recruiting) — used by JPMorgan Chase, Oracle
    itself, and many large enterprises. Undocumented but public: this is the
    same endpoint the candidate-facing search box calls, no login needed.

    Two quirks that matter:
    - limit/offset must live INSIDE the 'finder' param string, not as
      normal query params, or every page returns identical rows.
    - the top-level 'hasMore' flag is unreliable; paginate against
      TotalJobsCount instead.

    Some tenants (e.g. Oracle's own careers.oracle.com) front a backend
    *.oraclecloud.com tenant with a custom/vanity domain. This defaults to
    calling the API on the careers_url's own domain, which works when that
    domain proxies the API too. If a company's fetch comes back empty,
    check the real backend host via browser DevTools (Network tab, look for
    'recruitingCEJobRequisitions' requests) and set "api_domain" in that
    company's config entry to override.
    """
    domain = api_domain or urlparse(careers_url).netloc
    public_domain = urlparse(careers_url).netloc  # always use this for links people click
    site_path_match = re.search(r"/sites/([^/?#]+)", urlparse(careers_url).path, re.IGNORECASE)
    site_path = site_path_match.group(1) if site_path_match else "jobsearch"
    site_number = resolve_oracle_site_number(careers_url)

    oracle_headers = {
        "ora-irc-cx-userid": str(uuid.uuid4()),
        "ora-irc-language": "en",
        "content-type": "application/vnd.oracle.adf.resourceitem+json;charset=utf-8",
    }
    listings_path = "/hcmRestApi/resources/latest/recruitingCEJobRequisitions"
    expand = "requisitionList.workLocation,requisitionList.secondaryLocations"

    jobs = []
    offset = 0
    limit = 200
    total = None

    while True:
        finder = f"findReqs;siteNumber={site_number},limit={limit},offset={offset}"
        url = f"https://{domain}{listings_path}?onlyData=true&expand={expand}&finder={finder}"
        r = requests.get(url, headers=oracle_headers, timeout=20)
        r.raise_for_status()
        try:
            data = r.json()
        except ValueError:
            log.error(
                f"[{company}] Oracle API at {domain} returned non-JSON (status {r.status_code}). "
                f"This usually means the API isn't hosted on this domain — try setting "
                f"'api_domain' in the config to the real backend host (check DevTools -> Network "
                f"tab on the careers page for 'recruitingCEJobRequisitions' requests). "
                f"First 200 chars of response: {r.text[:200]!r}"
            )
            break
        items = data.get("items") or []
        if not items:
            break
        req_list = items[0].get("requisitionList") or []
        if not req_list:
            break

        for j in req_list:
            posted = None
            raw_date = j.get("PostedDate")
            if raw_date:
                try:
                    posted = datetime.fromisoformat(raw_date.replace("Z", "+00:00")).replace(tzinfo=None)
                except Exception:
                    posted = None
            job_id = j.get("Id")
            jobs.append({
                "title": j.get("Title", ""),
                "link": f"https://{public_domain}/hcmUI/CandidateExperience/en/sites/{site_path}/job/{job_id}",
                "posted_date": posted,
                "source": company,
                "location": j.get("PrimaryLocation", ""),
                "description": "",  # fetched on demand via get_oracle_job_details()
                "_oracle_domain": domain,
                "_oracle_site_number": site_number,
                "_oracle_job_id": job_id,
            })

        if total is None:
            total = items[0].get("TotalJobsCount") or 0
        offset += len(req_list)
        if total and offset >= total:
            break
        if not total and len(req_list) < limit:
            break
        time.sleep(0.3)

    return jobs


def get_oracle_job_details(domain, site_number, job_id) -> str:
    """Fetch full description text for one Oracle Cloud job. Only call this
    for jobs that already passed the keyword/date/location filters, since
    it's one extra request per job."""
    path = "/hcmRestApi/resources/latest/recruitingCEJobRequisitionDetails"
    finder = f'ById;Id="{job_id}",siteNumber={site_number}'
    url = f"https://{domain}{path}?expand=all&onlyData=true&finder={finder}"
    oracle_headers = {
        "ora-irc-cx-userid": str(uuid.uuid4()),
        "ora-irc-language": "en",
        "content-type": "application/vnd.oracle.adf.resourceitem+json;charset=utf-8",
    }
    try:
        r = requests.get(url, headers=oracle_headers, timeout=20)
        r.raise_for_status()
        items = r.json().get("items") or []
        if not items:
            return ""
        d = items[0]
        parts = [
            d.get("ExternalDescriptionStr", ""),
            d.get("ExternalQualificationsStr", ""),
            d.get("ExternalResponsibilitiesStr", ""),
        ]
        return strip_html(" ".join(p for p in parts if p))
    except Exception:
        return ""


def fetch_description_fallback(job) -> str:
    """Generic fallback: fetch the individual job posting page and strip
    HTML. Works when the job detail page itself is server-rendered (true
    for most Greenhouse/Lever/Ashby job pages even though their search
    boards are JS-driven). Won't work for JS-only pages (e.g. Workday)."""
    link = job.get("link")
    if not link:
        return ""
    try:
        r = requests.get(link, headers=HEADERS, timeout=15)
        r.raise_for_status()
        return strip_html(r.text)
    except Exception:
        return ""


def fetch_generic(company, url):
    """
    Fallback for custom career sites (and Workday, which needs JS rendering).
    Plain requests + BeautifulSoup can't see JS-rendered content, so this
    will only work for simple server-rendered pages out of the box.
    For JS-heavy sites, see fallback_playwright.py — you'll likely need to
    add a couple of site-specific CSS selectors there per company.
    """
    log.warning(f"[{company}] Using generic fallback — verify results manually, "
                f"this ATS type often needs custom selectors.")
    return []


ATS_FETCHERS = {
    "greenhouse": fetch_greenhouse,
    "lever": fetch_lever,
    "ashby": fetch_ashby,
}

# ATS types that need the full careers URL rather than just a slug.
URL_BASED_FETCHERS = {
    "oraclecloud": fetch_oracle_cloud,
}


def fetch_jobs_for_company(company_entry):
    name = company_entry["name"]
    url = company_entry["url"]
    ats = company_entry.get("ats") or detect_ats(url)

    try:
        if ats in URL_BASED_FETCHERS:
            if ats == "oraclecloud":
                return URL_BASED_FETCHERS[ats](name, url, company_entry.get("api_domain"))
            return URL_BASED_FETCHERS[ats](name, url)
        elif ats in ATS_FETCHERS:
            slug = company_entry.get("slug") or extract_slug(url)
            return ATS_FETCHERS[ats](name, slug)
        else:
            return fetch_generic(name, url)
    except Exception as e:
        log.error(f"[{name}] Failed to fetch ({ats}): {e}")
        return []


# ----------------------------------------------------------------------
# FILTERING
# ----------------------------------------------------------------------
def matches_keywords(title: str) -> bool:
    t = title.lower()
    return any(k in t for k in KEYWORDS)


def within_date_range(posted_date) -> bool:
    if posted_date is None:
        # Unknown date — include it so you can manually check, rather than
        # silently dropping a possibly-relevant job.
        return True
    return datetime.now() - posted_date <= timedelta(days=DAYS_THRESHOLD)


def is_india_location(location: str) -> bool:
    if not location:
        # Unknown location — include it so you can manually check, rather
        # than silently dropping a possibly-relevant job.
        return True
    loc = location.lower()
    return any(k in loc for k in INDIA_LOCATION_KEYWORDS)


# Best-effort patterns for "X years", "X-Y years", "X+ years" style phrasing.
_EXP_RANGE = re.compile(r"(\d{1,2})\s*(?:-|to)\s*(\d{1,2})\s*\+?\s*(?:years?|yrs?)", re.I)
_EXP_PLUS = re.compile(r"(\d{1,2})\s*\+\s*(?:years?|yrs?)", re.I)
_EXP_PLAIN = re.compile(r"(\d{1,2})\s*(?:years?|yrs?)(?:\s+of)?\s+experience", re.I)


def extract_min_experience_years(text: str):
    """Best-effort regex scan of a job description for the minimum years of
    experience mentioned. Returns None if nothing recognizable is found —
    this is a heuristic on free text, not a structured field, so treat it
    as a helpful filter, not a guarantee. Worth spot-checking results."""
    if not text:
        return None
    candidates = []
    for m in _EXP_RANGE.finditer(text):
        candidates.append(int(m.group(1)))
    for m in _EXP_PLUS.finditer(text):
        candidates.append(int(m.group(1)))
    for m in _EXP_PLAIN.finditer(text):
        candidates.append(int(m.group(1)))
    return min(candidates) if candidates else None


def matches_experience(min_years) -> bool:
    if min_years is None:
        # Couldn't determine it from the text — include it flagged as
        # "Unknown", rather than silently dropping a possibly-relevant job.
        return True
    return min_years <= MAX_EXPERIENCE_YEARS


# ----------------------------------------------------------------------
# EXCEL OUTPUT (dedupe by link, append new rows)
# ----------------------------------------------------------------------
COLUMNS = ["Date Found", "Company", "Job Title", "Location", "Experience", "Posted Date", "Link"]


def load_existing_links():
    if not OUTPUT_FILE.exists():
        return set(), None
    wb = load_workbook(OUTPUT_FILE)
    ws = wb.active
    links = set()
    link_col = COLUMNS.index("Link")
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and len(row) > link_col and row[link_col]:
            links.add(row[link_col])
    return links, wb


def write_results(new_jobs):
    existing_links, wb = load_existing_links()
    if wb is None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Jobs"
        ws.append(COLUMNS)
        for i, col in enumerate(COLUMNS, 1):
            ws.column_dimensions[get_column_letter(i)].width = 60 if col == "Link" else 20
    else:
        ws = wb.active

    added = 0
    for job in new_jobs:
        if job["link"] in existing_links or not job["link"]:
            continue
        posted_str = job["posted_date"].strftime("%Y-%m-%d") if job["posted_date"] else "Unknown"
        ws.append([
            datetime.now().strftime("%Y-%m-%d"),
            job["source"],
            job["title"],
            job.get("location") or "Unknown",
            job.get("experience_note", "Unknown"),
            posted_str,
            job["link"],
        ])
        existing_links.add(job["link"])
        added += 1

    wb.save(OUTPUT_FILE)
    return added


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    if not CONFIG_FILE.exists():
        log.error(f"Config file not found: {CONFIG_FILE}. Copy companies_config.example.json "
                  f"to companies_config.json and fill in your company URLs.")
        return

    companies = json.loads(CONFIG_FILE.read_text())
    log.info(f"Loaded {len(companies)} companies from config.")

    matched_jobs = []
    for company in companies:
        jobs = fetch_jobs_for_company(company)
        for job in jobs:
            # Cheap filters first — title keywords, posting date, location —
            # before touching anything that needs an extra network call.
            if not matches_keywords(job["title"]):
                continue
            if not within_date_range(job["posted_date"]):
                continue
            if not is_india_location(job.get("location", "")):
                continue

            # Experience is free text, so only fetch/parse it for jobs that
            # already survived the filters above.
            description = job.get("description") or ""
            if not description:
                if job.get("_oracle_job_id"):
                    description = get_oracle_job_details(
                        job["_oracle_domain"], job["_oracle_site_number"], job["_oracle_job_id"]
                    )
                else:
                    description = fetch_description_fallback(job)

            min_years = extract_min_experience_years(description)
            if not matches_experience(min_years):
                continue

            job["experience_note"] = f"{min_years}+ yrs" if min_years is not None else "Unknown - verify"
            matched_jobs.append(job)
        time.sleep(1)  # be polite between requests

    added = write_results(matched_jobs)
    log.info(f"Done. {len(matched_jobs)} matching jobs found this run, "
             f"{added} new rows added to {OUTPUT_FILE.name}.")


if __name__ == "__main__":
    main()