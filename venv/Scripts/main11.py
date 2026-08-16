"""
Daily Job Search Automation
----------------------------
Scrapes career pages listed in companies_config.json, keeps jobs posted
within the last N days that match your role keywords, and writes results
to a dated output folder (output/YYYY-MM-DD/job_search_results.xlsx).

Cross-day dedup: a persistent seen_jobs.json (at the top level, NOT inside
a dated folder) tracks every job link ever written across ALL past runs,
not just yesterday's. This matters because several integrated companies
don't expose a posted-date field at all, so without persistent tracking
the same job would resurface in every single run indefinitely. Companies
fetch in parallel (ThreadPoolExecutor) since these are independent,
I/O-bound network calls -- notably faster wall-clock time than the old
sequential loop, especially with 15+ companies.

Usage:
    python main.py
"""

import json
import re
import time
import uuid
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse, quote

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "companies_config.json"

# Every run's results land in their own dated folder, so history is never
# overwritten and each day is easy to review independently.
TODAY_STR = datetime.now().strftime("%Y-%m-%d")
TODAY_DIR = BASE_DIR / "output" / TODAY_STR
TODAY_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = TODAY_DIR / "job_search_results.xlsx"
LOG_FILE = TODAY_DIR / "scraper.log"

# Persistent, NOT dated -- this is what makes cross-day dedup work. Grows
# forever unless pruned; see prune_seen_jobs() if you want a rolling window.
SEEN_JOBS_FILE = BASE_DIR / "seen_jobs.json"

DAYS_THRESHOLD = 2
MAX_EXPERIENCE_YEARS = 4
MAX_WORKERS = 10  # companies fetched concurrently; raise/lower to taste

# Transient-failure retry, mainly for unattended scheduled runs where
# there's no one around to notice a company came back empty for a day.
MAX_FETCH_RETRIES = 2
RETRY_BACKOFF_SECONDS = 8
TRANSIENT_STATUS_CODES = {500, 502, 503, 504}

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

# Broad terms used to page through keyword-driven search platforms (BP,
# SAP, Nutanix, Citi/Barclays) that don't expose a "list everything" mode.
PAGINATION_KEYWORDS = ["backend", "full stack", "java", "react", "software engineer"]

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
    if "search-results" in path:
        # Weak heuristic for Phenom People (CareerConnect) sites, e.g.
        # careers.mastercard.com. Many custom sites also use this path
        # segment, so set "ats": "phenom" explicitly in the config when
        # you know a company uses Phenom rather than relying on this.
        return "phenom"
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


def extract_balanced_json(text: str, start_marker: str):
    """Find start_marker, then balanced-brace parse the following {...}
    object out of raw JS/HTML text and json.loads it. Returns None if the
    marker isn't found or the extracted text isn't valid JSON."""
    idx = text.find(start_marker)
    if idx == -1:
        return None
    brace_start = text.find("{", idx)
    if brace_start == -1:
        return None
    depth = 0
    in_str = False
    esc = False
    i = brace_start
    while i < len(text):
        c = text[i]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
        else:
            if c == '"':
                in_str = True
            elif c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    break
        i += 1
    try:
        return json.loads(text[brace_start:i + 1])
    except Exception:
        return None


# (uses shared PAGINATION_KEYWORDS defined near the top config section)


def parse_phenom_page(html: str):
    """Extract the server-rendered job batch + site bootstrap config from a
    Phenom People (CareerConnect) search-results page. Verified structure:
    the page embeds `phApp.ddo = {...}` containing
    ddo.eagerLoadRefineSearch.data.jobs (first batch) and .totalHits."""
    ddo = extract_balanced_json(html, "phApp.ddo = {") or {}
    search = ddo.get("eagerLoadRefineSearch") or {}
    jobs = (search.get("data") or {}).get("jobs") or []
    return {"jobs": jobs, "total_hits": search.get("totalHits")}


def phenom_job_to_dict(j, company):
    posted = None
    raw_date = j.get("postedDate") or j.get("dateCreated")
    if raw_date:
        try:
            posted = datetime.fromisoformat(raw_date.replace("Z", "+00:00")).replace(tzinfo=None)
        except Exception:
            posted = None
    return {
        "title": j.get("title", ""),
        # applyUrl often points straight at the real backend ATS (Workday,
        # etc.) — that's fine, it's still the correct link for the user.
        "link": j.get("applyUrl") or "",
        "posted_date": posted,
        "source": company,
        "location": j.get("cityStateCountry") or j.get("location") or "",
        # Only a short teaser is available from the search results, not the
        # full JD — the experience-years regex will miss postings that state
        # requirements outside this teaser. Best-effort, like elsewhere.
        "description": j.get("descriptionTeaser", ""),
        "_job_id": j.get("jobId") or j.get("reqId"),
    }


def fetch_phenom(company, careers_url):
    """Phenom People / CareerConnect career sites (e.g. Mastercard). Job
    data for the first batch is embedded server-side in the search-results
    page's HTML — confirmed via real captured traffic, no session/CSRF
    needed for that part.

    Getting beyond the first ~10 jobs is best-effort: this tries
    keyword-scoped GET requests (?keywords=...&size=50), a common Phenom
    convention, but it's NOT confirmed working for every tenant. If a
    company keeps returning the same first page regardless of keyword,
    query-param filtering isn't supported there and you're only seeing
    that tenant's default top results, not full coverage — worth checking
    scraper.log's "total openings" line against how many you're getting.
    """
    jobs_by_id = {}

    def collect(url):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
        except Exception as e:
            log.warning(f"[{company}] Phenom page fetch failed for {url}: {e}")
            return None
        parsed = parse_phenom_page(r.text)
        for j in parsed["jobs"]:
            jd = phenom_job_to_dict(j, company)
            key = jd["_job_id"] or jd["link"]
            if key:
                jobs_by_id[key] = jd
        return parsed

    base = collect(careers_url)
    if base is None:
        return []
    if base.get("total_hits"):
        log.info(f"[{company}] Phenom site reports {base['total_hits']} total openings "
                 f"company-wide; pulling a best-effort keyword-scoped sample, not the full set.")

    sep = "&" if "?" in careers_url else "?"
    for kw in PAGINATION_KEYWORDS:
        collect(f"{careers_url}{sep}keywords={quote(kw)}&size=50")
        time.sleep(0.5)

    return list(jobs_by_id.values())


def fetch_amazon(company, careers_url):
    """Amazon's own in-house job search API (amazon.jobs). Clean public
    JSON endpoint, no auth, full descriptions and qualifications included.
    Verified against real captured traffic."""
    jobs = []
    offset = 0
    limit = 100
    total = None
    while True:
        url = (
            "https://www.amazon.jobs/en/search.json"
            f"?loc_query=India&country=IND&offset={offset}&result_limit={limit}&sort=recent"
        )
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        data = r.json()
        batch = data.get("jobs") or []
        if not batch:
            break
        if total is None:
            total = data.get("hits") or 0
        for j in batch:
            posted = None
            raw_date = j.get("posted_date")  # e.g. "July  8, 2026"
            if raw_date:
                try:
                    posted = datetime.strptime(re.sub(r"\s+", " ", raw_date).strip(), "%B %d, %Y")
                except Exception:
                    posted = None
            jobs.append({
                "title": j.get("title", ""),
                "link": f"https://www.amazon.jobs{j.get('job_path', '')}",
                "posted_date": posted,
                "source": company,
                "location": j.get("normalized_location", ""),
                "description": strip_html(
                    (j.get("description_short") or "") + " " +
                    (j.get("basic_qualifications") or "") + " " +
                    (j.get("preferred_qualifications") or "")
                ),
            })
        offset += len(batch)
        if total and offset >= total:
            break
        if len(batch) < limit:
            break
        time.sleep(0.5)
    return jobs


def fetch_bp(company, careers_url):
    """BP careers uses Algolia search (production_bp_jobs index). The API
    key captured is a search-only key meant for frontend/browser use (this
    is standard Algolia practice, not a leaked secret) -- verified against
    real captured traffic."""
    app_id = "UM59DWRPA1"
    api_key = "0f3e2a2d4ca48d7cfeba5e0ea876e172"
    index = "production_bp_jobs"
    url = f"https://{app_id.lower()}-dsn.algolia.net/1/indexes/*/queries"
    headers = {
        "content-type": "application/x-www-form-urlencoded",
        "x-algolia-api-key": api_key,
        "x-algolia-application-id": app_id,
        # BP's Algolia key is domain-restricted and checks these — without
        # them the request gets a 403 (confirmed: this was the actual bug).
        "Origin": "https://careers.bp.com",
        "Referer": "https://careers.bp.com/",
        "User-Agent": HEADERS["User-Agent"],
    }

    jobs = []
    for kw in PAGINATION_KEYWORDS:
        page = 0
        while True:
            body = {
                "requests": [{
                    "indexName": index,
                    "params": f"query={quote(kw)}&hitsPerPage=50&page={page}",
                }]
            }
            r = requests.post(url, headers=headers, data=json.dumps(body), timeout=20)
            r.raise_for_status()
            result = r.json()["results"][0]
            hits = result.get("hits", [])
            if not hits:
                break
            for h in hits:
                job_code = (h.get("job_code") or [""])[0]
                posted = None
                raw_date = (h.get("posting_date") or [None])[0]
                if raw_date:
                    try:
                        posted = datetime.strptime(raw_date, "%Y-%m-%d")
                    except Exception:
                        posted = None
                summary = strip_html((h.get("summary") or [""])[0]) if isinstance(h.get("summary"), list) else ""
                jobs.append({
                    "title": h.get("title", ""),
                    "link": f"https://careers.bp.com/job-description/{job_code}" if job_code else "",
                    "posted_date": posted,
                    "source": company,
                    "location": ", ".join(h.get("location") or []),
                    "description": summary,  # only available via _highlightResult snippet, may be truncated
                    "_id": h.get("objectID"),
                })
            page += 1
            if page >= result.get("nbPages", 1):
                break
            time.sleep(0.3)
    # dedupe across keyword queries
    seen = {}
    for j in jobs:
        seen[j.get("_id") or j["link"]] = j
    return list(seen.values())


def fetch_cohesity(company, careers_url):
    """Cohesity's careers page (AEM-backed) returns ALL open jobs grouped
    by department in a single request -- no pagination needed. Verified
    against real captured traffic. Real ATS is Workday underneath, same
    pattern as Mastercard: job links go straight to the Workday apply page.

    Requires a Referer header matching the public careers page -- AEM
    backends commonly enforce this as a lightweight same-origin check.
    (Confirmed root cause of a run that silently returned 0 jobs: this
    header was missing.)"""
    url = "https://www.cohesity.com/bin/cohesity/open-positions/"
    headers = {**HEADERS, "Referer": "https://www.cohesity.com/careers/open-positions/"}
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    data = r.json()
    by_dept = data.get("job_data") or {}
    if not by_dept:
        log.warning(f"[{company}] Response had no 'job_data' — first 200 chars: {r.text[:200]!r}")

    jobs = []
    for dept, dept_jobs in by_dept.items():
        for j in dept_jobs:
            jobs.append({
                "title": j.get("title", ""),
                "link": j.get("jobUrl", ""),
                "posted_date": None,  # not provided by this endpoint
                "source": company,
                "location": j.get("primaryLocation") or j.get("country") or "",
                "description": "",  # not provided; would need per-job Workday fetch
                "_id": j.get("req_id") or j.get("JobID"),
            })
    return jobs


def fetch_radancy(company, careers_url):
    """Radancy / TalentBrew career sites (e.g. Citi, Barclays). Same
    backend platform, but tenants customize the frontend template
    differently -- verified against real captured traffic for both, and
    they use genuinely different HTML structures:
      - Citi:     <li class="sr-job-item"><a class="sr-job-item__link">...
                  no posted-date field on this tenant.
      - Barclays: <div class="list-item"><a class="job-title--link">...
                  DOES show a posted date, but year-less ("09 Aug") --
                  assumed to be the current year, rolled back one year if
                  that would put it in the future.
    Tries both selector sets per page; uses whichever matches."""
    base = f"{urlparse(careers_url).scheme}://{urlparse(careers_url).netloc}"
    results_url = f"{base}/search-jobs/results"
    headers = {**HEADERS, "X-Requested-With": "XMLHttpRequest", "Referer": careers_url}

    def parse_barclays_date(text):
        text = text.strip()
        if not text:
            return None
        try:
            dt = datetime.strptime(f"{text} {datetime.now().year}", "%d %b %Y")
            if dt > datetime.now() + timedelta(days=3):
                dt = dt.replace(year=dt.year - 1)
            return dt
        except Exception:
            return None

    jobs_by_id = {}
    for kw in PAGINATION_KEYWORDS:
        page = 1
        while True:
            params = {
                "ActiveFacetID": 0, "CurrentPage": page, "RecordsPerPage": 50,
                "TotalContentResults": "", "Distance": 50, "RadiusUnitType": 0,
                "Keywords": kw, "Location": "", "ShowRadius": "False",
                "IsPagination": "False" if page == 1 else "True",
                "CustomFacetName": "", "FacetTerm": "", "FacetType": 0,
                "SearchResultsModuleName": "Search Results",
                "SearchFiltersModuleName": "Search Filters",
                "SortCriteria": 5, "SortDirection": 1, "SearchType": 5,
                "PostalCode": "", "ResultsType": 0,
                "fc": "", "fl": "", "fcf": "", "afc": "", "afl": "", "afcf": "",
                "TotalContentPages": "NaN",
            }
            r = requests.get(results_url, headers=headers, params=params, timeout=20)
            r.raise_for_status()
            try:
                data = r.json()
            except ValueError:
                log.warning(f"[{company}] Radancy search returned non-JSON, stopping pagination.")
                break
            soup = BeautifulSoup(data.get("results", ""), "html.parser")

            found_any = False
            # Template A: Citi-style
            for item in soup.select("li.sr-job-item"):
                link_tag = item.select_one("a.sr-job-item__link")
                if not link_tag:
                    continue
                found_any = True
                job_id = link_tag.get("data-job-id", "")
                loc_tag = item.select_one(".sr-job-location")
                jobs_by_id[job_id or link_tag.get("href", "")] = {
                    "title": link_tag.get_text(strip=True),
                    "link": base + link_tag.get("href", ""),
                    "posted_date": None,  # not exposed on this template
                    "source": company,
                    "location": loc_tag.get_text(strip=True) if loc_tag else "",
                    "description": "",
                }
            # Template B: Barclays-style
            for item in soup.select("div.list-item"):
                link_tag = item.select_one("a.job-title--link")
                if not link_tag:
                    continue
                found_any = True
                job_id = link_tag.get("data-job-id", "")
                loc_tag = item.select_one(".job-location")
                date_tag = item.select_one(".job-date span")
                jobs_by_id[job_id or link_tag.get("href", "")] = {
                    "title": link_tag.get_text(strip=True),
                    "link": base + link_tag.get("href", ""),
                    "posted_date": parse_barclays_date(date_tag.get_text()) if date_tag else None,
                    "source": company,
                    "location": loc_tag.get_text(strip=True) if loc_tag else "",
                    "description": "",
                }

            if not found_any:
                break
            page += 1
            if page > 20:  # safety cap
                break
            time.sleep(0.3)
    return list(jobs_by_id.values())


def fetch_sap(company, careers_url):
    """SAP's SuccessFactors Career Site Builder. Server-rendered HTML table,
    paginated via startrow=. Verified against real captured traffic. No
    posted-date field is displayed on listing rows (only used internally
    for the "referencedate" sort), so posted_date is left None."""
    base = f"{urlparse(careers_url).scheme}://{urlparse(careers_url).netloc}"
    jobs_by_link = {}
    for kw in PAGINATION_KEYWORDS:
        startrow = 0
        while True:
            url = (
                f"{base}/search/?q={quote(kw)}&sortColumn=referencedate"
                f"&sortDirection=desc&startrow={startrow}"
            )
            r = requests.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, "html.parser")
            rows = soup.select("tr.data-row")
            if not rows:
                break
            for row in rows:
                link_tag = row.select_one("a.jobTitle-link")
                loc_tag = row.select_one(".jobLocation")
                if not link_tag:
                    continue
                href = link_tag.get("href", "")
                jobs_by_link[href] = {
                    "title": link_tag.get_text(strip=True),
                    "link": base + href if href.startswith("/") else href,
                    "posted_date": None,  # not exposed on listing rows
                    "source": company,
                    "location": loc_tag.get_text(strip=True) if loc_tag else "",
                    "description": "",
                }
            startrow += 25
            if startrow > 500:  # safety cap
                break
            time.sleep(0.3)
    return list(jobs_by_link.values())


def fetch_doordash(company, careers_url):
    """DoorDash's careers site (custom WordPress theme, Cloudflare-
    protected). Plain server-rendered HTML, and location filtering works
    server-side via a ?location= query param -- verified against real
    captured traffic. No posted-date on listing cards.

    CAVEAT: this site sits behind Cloudflare. A prior run got a 403. The
    extra browser-matching headers below may help with header-based bot
    rules, but Cloudflare can also fingerprint at the TLS-handshake level,
    which the `requests` library cannot replicate -- if 403s persist after
    this change, that's very likely why, and would need a real browser
    (e.g. Playwright) to reliably get past, not just better headers."""
    parsed = urlparse(careers_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    doordash_headers = {
        **HEADERS,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    jobs_by_link = {}
    page = 1
    while True:
        url = f"{base}/job-search/?keyword=&location=India&spage={page}"
        r = requests.get(url, headers=doordash_headers, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        items = soup.select("div.job-item")
        if not items:
            break
        for item in items:
            title_tag = item.select_one(".title-container .value a")
            loc_tag = item.select_one(".location-container .value-secondary")
            dept_tag = item.select_one(".department-container .value-secondary")
            if not title_tag:
                continue
            href = title_tag.get("href", "")
            jobs_by_link[href] = {
                "title": title_tag.get_text(strip=True),
                "link": href,
                "posted_date": None,  # not shown on listing cards
                "source": company,
                "location": loc_tag.get_text(strip=True) if loc_tag else "",
                "description": dept_tag.get_text(strip=True) if dept_tag else "",
            }
        page += 1
        if page > 15:  # safety cap
            break
        time.sleep(0.3)
    return list(jobs_by_link.values())


def fetch_deloitte(company, careers_url):
    """Deloitte USI careers (legacy Oracle Taleo-based portal). Plain
    server-rendered HTML, paginated via jobOffset= -- verified against
    real captured traffic (595 total India openings at capture time). No
    posted-date field available on listing cards."""
    parsed = urlparse(careers_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    jobs_by_link = {}
    offset = 0
    page_size = 10  # confirmed value from real captured traffic
    max_jobs = 150  # safety cap -- raise if you want deeper coverage
    while offset < max_jobs:
        url = (f"{base}/en_US/careersUSI/SearchJobs/India"
               f"?listFilterMode=1&jobSort=relevancy&jobRecordsPerPage={page_size}&jobOffset={offset}")
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        articles = soup.select("article.article--result")
        if not articles:
            break
        for a in articles:
            link_tag = a.select_one("h3 a")
            if not link_tag:
                continue
            spans = a.select(".article__header__text__subtitle span")
            location = spans[-1].get_text(strip=True) if spans else ""
            jobs_by_link[link_tag.get("href", "")] = {
                "title": link_tag.get_text(strip=True),
                "link": link_tag.get("href", ""),
                "posted_date": None,  # not shown on listing cards
                "source": company,
                "location": location,
                "description": "",
            }
        offset += page_size
        time.sleep(0.3)
    return list(jobs_by_link.values())


def fetch_ubs_brassring(company, careers_url):
    """UBS's BrassRing (Kenexa/IBM) career site. Requires an
    EncryptedSessionValue token embedded in the initial page load's
    bootstrap data (HTML-entity-encoded), which must be extracted and
    replayed in the search POST -- verified against real captured traffic.
    A single unfiltered request already returns all India openings (19 at
    capture time, JobsCount matched exactly), so no pagination needed.
    Real fields (verified): jobtitle, formtext23 (location), lastupdated
    (DD-Mon-YYYY), jobdescription, all nested inside a Questions array.

    KNOWN BLOCKER (unresolved): the live request also failed with a 500,
    separate from the EncryptedSessionValue token. The real POST requires
    an additional 'RFT' header whose value does NOT match any static token
    findable in the page HTML (confirmed by direct comparison against a
    real capture) -- it looks like a client-side-computed anti-bot
    fingerprint, not something a plain HTTP request can derive. This would
    likely need a real browser (e.g. Playwright) to solve properly rather
    than further guessing at a formula."""
    parsed = urlparse(careers_url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    qs = dict(x.split("=") for x in parsed.query.split("&") if "=" in x)
    partner_id = qs.get("partnerid") or qs.get("PartnerId") or "25008"
    site_id = qs.get("siteid") or qs.get("SiteId") or "5012"

    r = requests.get(careers_url, headers=HEADERS, timeout=20)
    r.raise_for_status()
    clean_html = r.text.replace("\\", "")
    m = re.search(r"EncryptedSessionValue&quot;:&quot;([^&]+)&quot;", clean_html)
    if not m:
        log.warning(f"[{company}] Could not find EncryptedSessionValue token on the careers "
                    f"page -- BrassRing may have changed its bootstrap format.")
        return []
    token = m.group(1)

    body = {
        "PartnerId": partner_id, "SiteId": site_id, "Keyword": "", "Location": "India",
        "KeywordCustomSolrFields": "FORMTEXT2,FORMTEXT21,AutoReq,Department,JobTitle",
        "LocationCustomSolrFields": "FORMTEXT2,FORMTEXT23,Location",
        "TurnOffHttps": False, "Latitude": 0, "Longitude": 0,
        "PowerSearchOptions": {"PowerSearchOption": []},
        "encryptedsessionvalue": token,
    }
    headers = {
        **HEADERS, "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest", "Referer": careers_url,
        "Origin": base,
    }
    try:
        r2 = requests.post(f"{base}/TgNewUI/Search/Ajax/MatchedJobs", headers=headers,
                            data=json.dumps(body), timeout=20)
        r2.raise_for_status()
    except requests.exceptions.HTTPError as e:
        log.error(f"[{company}] UBS search POST failed ({e}). This site requires an 'RFT' "
                  f"anti-bot header this script cannot compute (see fetch_ubs_brassring "
                  f"docstring) -- likely needs a real browser to solve, not a header tweak.")
        return []
    job_list = ((r2.json().get("Jobs") or {}).get("Job")) or []

    jobs = []
    for j in job_list:
        answers = {q["QuestionName"]: q.get("Value", "") for q in j.get("Questions", [])}
        posted = None
        raw_date = answers.get("lastupdated")  # e.g. "14-Aug-2026"
        if raw_date:
            try:
                posted = datetime.strptime(raw_date, "%d-%b-%Y")
            except Exception:
                posted = None
        jobs.append({
            "title": answers.get("jobtitle", ""),
            "link": j.get("Link", ""),
            "posted_date": posted,
            "source": company,
            "location": answers.get("formtext23", ""),
            "description": strip_html(answers.get("jobdescription", "")),
        })
    return jobs


def fetch_nutanix(company, careers_url):
    """Nutanix's Umbraco-based job board. Plain server-rendered HTML with
    clean ?search=&page= pagination. Verified against real captured
    traffic. No posted-date shown on listing cards, so posted_date is
    left None."""
    base = f"{urlparse(careers_url).scheme}://{urlparse(careers_url).netloc}"
    jobs_by_id = {}
    for kw in PAGINATION_KEYWORDS:
        page = 1
        while True:
            url = f"{base}/en/jobs/?search={quote(kw)}&page={page}&origin=global"
            r = requests.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            soup = BeautifulSoup(r.text, "html.parser")
            cards = soup.select("div.card-job")
            if not cards:
                break
            for card in cards:
                link_tag = card.select_one("a.js-view-job")
                loc_tag = card.select_one(".job-meta-location")
                if not link_tag:
                    continue
                job_id = card.get("data-id", "")
                href = link_tag.get("href", "")
                jobs_by_id[job_id or href] = {
                    "title": link_tag.get_text(strip=True),
                    "link": base + href if href.startswith("/") else href,
                    "posted_date": None,  # not shown on listing cards
                    "source": company,
                    "location": loc_tag.get_text(strip=True) if loc_tag else "",
                    "description": "",
                }
            page += 1
            if page > 20:  # safety cap
                break
            time.sleep(0.3)
    return list(jobs_by_id.values())


def _extract_google_ds_block(html: str, key: str):
    """Google career pages embed search results via AF_initDataCallback
    blocks like `{key: 'ds:1', hash: '2', data: [...]}`. This finds the
    named block and balanced-bracket parses its `data` array."""
    idx = html.find(f"key: '{key}'")
    if idx == -1:
        return None
    data_start = html.find("data:", idx) + len("data:")
    start = html.find("[", data_start)
    if start == -1:
        return None
    depth, in_str, esc, i = 0, False, False, start
    while i < len(html):
        c = html[i]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
        else:
            if c == '"':
                in_str = True
            elif c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
                if depth == 0:
                    break
        i += 1
    try:
        return json.loads(html[start:i + 1])
    except Exception:
        return None


def fetch_google(company, careers_url):
    """Google Careers. Job data is embedded server-side in the page via an
    internal `AF_initDataCallback` / 'ds:1' block (positional array, not
    named fields -- verified against real captured traffic by inspecting
    each index). Real total-count field confirmed (305 at capture time),
    so this paginates via &page=N -- NOT confirmed to work beyond page 1
    since only page 1 was captured; if a page comes back with the same
    jobs as the previous one, that means this parameter guess is wrong
    and you're only getting page 1 repeated.

    Field mapping (by position, verified against real data):
      1=title, 2=apply link, 3/4/10/19=[None, description HTML] pairs,
      9=locations list of [display_name, address_lines, city, zip, state,
      country], 12/13/14=[epoch_seconds, nanos] timestamps (uses the max
      of these as posted_date -- exact semantic of each isn't documented,
      but the latest one is the safest "freshness" signal)."""
    jobs_by_id = {}
    seen_ids_per_page = None
    page = 0
    while True:
        sep = "&" if "?" in careers_url else "?"
        url = careers_url if page == 0 else f"{careers_url}{sep}page={page}"
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        data = _extract_google_ds_block(r.text, "ds:1")
        if not data or not data[0]:
            break
        jobs_raw = data[0]
        ids_this_page = {j[0] for j in jobs_raw}
        if seen_ids_per_page is not None and ids_this_page == seen_ids_per_page:
            log.info(f"[{company}] Page {page} returned identical results to the previous "
                     f"page -- &page= pagination isn't supported here, stopping.")
            break
        seen_ids_per_page = ids_this_page

        for j in jobs_raw:
            locs = j[9] or []
            location = "; ".join(l[0] for l in locs if l and l[0])
            timestamps = [t[0] for t in (j[12], j[13], j[14]) if t]
            posted = datetime.fromtimestamp(max(timestamps)) if timestamps else None
            desc_parts = []
            for idx in (3, 4, 10, 19):
                field = j[idx] if idx < len(j) else None
                if field and len(field) > 1 and field[1]:
                    desc_parts.append(field[1])
            jobs_by_id[j[0]] = {
                "title": j[1],
                "link": j[2],
                "posted_date": posted,
                "source": company,
                "location": location,
                "description": strip_html(" ".join(desc_parts)),
            }

        total = data[2] if len(data) > 2 else None
        page_size = data[3] if len(data) > 3 else len(jobs_raw)
        page += 1
        if total and len(jobs_by_id) >= total:
            break
        if page > 30:  # safety cap
            break
        time.sleep(0.3)
    return list(jobs_by_id.values())


def fetch_microsoft(company, careers_url):
    """Microsoft Careers. Clean dedicated JSON search API -- verified
    against real captured traffic. Full description requires a per-job
    detail call (position_details), so this only fetches that for jobs
    that already survive the cheap filters (title/date/location), same
    pattern as Oracle Cloud's description fetching."""
    base = "https://apply.careers.microsoft.com"
    jobs = []
    start = 0
    page_size = 10  # confirmed value from real captured traffic
    total = None
    while True:
        url = (f"{base}/api/pcsx/search?domain=microsoft.com&query=&location=India"
               f"&start={start}&sort_by=distance&filter_include_remote=1")
        r = requests.get(url, headers={**HEADERS, "Accept": "application/json, text/plain, */*"},
                          timeout=20)
        r.raise_for_status()
        data = r.json().get("data") or {}
        positions = data.get("positions") or []
        if not positions:
            break
        if total is None:
            total = data.get("count") or 0
        for p in positions:
            posted = None
            ts = p.get("postedTs") or p.get("creationTs")
            if ts:
                try:
                    posted = datetime.fromtimestamp(ts)
                except Exception:
                    posted = None
            jobs.append({
                "title": p.get("name", ""),
                "link": base + p.get("positionUrl", ""),
                "posted_date": posted,
                "source": company,
                "location": "; ".join(p.get("locations") or []),
                "description": "",  # fetched on demand below for filtered survivors
                "_ms_position_id": p.get("id"),
            })
        start += len(positions)
        if total and start >= total:
            break
        time.sleep(0.3)
    return jobs


def get_microsoft_job_description(position_id) -> str:
    """On-demand full description fetch for one Microsoft job. Only call
    this for jobs that already passed the cheap filters."""
    url = (f"https://apply.careers.microsoft.com/api/pcsx/position_details"
           f"?position_id={position_id}&domain=microsoft.com&hl=en")
    try:
        r = requests.get(url, headers={**HEADERS, "Accept": "application/json, text/plain, */*"},
                          timeout=20)
        r.raise_for_status()
        return strip_html((r.json().get("data") or {}).get("jobDescription", ""))
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
    "phenom": fetch_phenom,
    "amazon": fetch_amazon,
    "bp_algolia": fetch_bp,
    "cohesity": fetch_cohesity,
    "radancy": fetch_radancy,
    "sap_successfactors": fetch_sap,
    "nutanix": fetch_nutanix,
    "doordash": fetch_doordash,
    "deloitte": fetch_deloitte,
    "ubs_brassring": fetch_ubs_brassring,
    "google": fetch_google,
    "microsoft": fetch_microsoft,
}


def fetch_jobs_for_company(company_entry):
    name = company_entry["name"]
    url = company_entry["url"]
    ats = company_entry.get("ats") or detect_ats(url)

    def _do_fetch():
        if ats in URL_BASED_FETCHERS:
            if ats == "oraclecloud":
                return URL_BASED_FETCHERS[ats](name, url, company_entry.get("api_domain"))
            return URL_BASED_FETCHERS[ats](name, url)
        elif ats in ATS_FETCHERS:
            slug = company_entry.get("slug") or extract_slug(url)
            return ATS_FETCHERS[ats](name, slug)
        else:
            return fetch_generic(name, url)

    # Retry only what looks transient (server hiccups, network blips) --
    # not permanent failures like a 403 block or an application-level bug,
    # where retrying just burns time for the same guaranteed outcome.
    last_exception = None
    for attempt in range(MAX_FETCH_RETRIES + 1):
        try:
            return _do_fetch()
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else None
            last_exception = e
            if status in TRANSIENT_STATUS_CODES and attempt < MAX_FETCH_RETRIES:
                log.warning(f"[{name}] Got HTTP {status} (likely transient), retrying "
                            f"({attempt + 1}/{MAX_FETCH_RETRIES}) in {RETRY_BACKOFF_SECONDS}s...")
                time.sleep(RETRY_BACKOFF_SECONDS)
                continue
            break
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
            last_exception = e
            if attempt < MAX_FETCH_RETRIES:
                log.warning(f"[{name}] {e.__class__.__name__}, retrying "
                            f"({attempt + 1}/{MAX_FETCH_RETRIES}) in {RETRY_BACKOFF_SECONDS}s...")
                time.sleep(RETRY_BACKOFF_SECONDS)
                continue
            break
        except Exception as e:
            last_exception = e
            break  # not a transient-looking error type — don't retry

    log.error(f"[{name}] Failed to fetch ({ats}) after "
              f"{MAX_FETCH_RETRIES + 1} attempt(s): {last_exception}")
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
def load_seen_jobs() -> set:
    """All job links ever written on any previous day. Empty set on first
    ever run (file won't exist yet)."""
    if not SEEN_JOBS_FILE.exists():
        return set()
    try:
        return set(json.loads(SEEN_JOBS_FILE.read_text()))
    except Exception:
        log.warning(f"Couldn't parse {SEEN_JOBS_FILE.name}, starting with empty history.")
        return set()


def save_seen_jobs(seen: set):
    SEEN_JOBS_FILE.write_text(json.dumps(sorted(seen)))


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
def process_company(company: dict) -> tuple:
    """Fetch + filter one company. Returns (name, raw_count, matched_jobs).
    Runs inside a worker thread -- must not touch shared mutable state
    other than through return values (it doesn't)."""
    name = company["name"]
    jobs = fetch_jobs_for_company(company)
    log.info(f"[{name}] Fetched {len(jobs)} raw job(s) before filtering.")
    if not jobs:
        log.warning(f"[{name}] Zero jobs fetched — check this company's fetcher/selectors "
                    f"even though no exception was raised (a silent 0 often means a changed "
                    f"page structure, a blocked/anti-bot response, or missing headers).")

    matched = []
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
            elif job.get("_ms_position_id"):
                description = get_microsoft_job_description(job["_ms_position_id"])
            else:
                description = fetch_description_fallback(job)

        min_years = extract_min_experience_years(description)
        if not matches_experience(min_years):
            continue

        job["experience_note"] = f"{min_years}+ yrs" if min_years is not None else "Unknown - verify"
        matched.append(job)

    log.info(f"[{name}] {len(matched)} job(s) matched all filters.")
    return name, len(jobs), matched


def main():
    if not CONFIG_FILE.exists():
        log.error(f"Config file not found: {CONFIG_FILE}. Copy companies_config.example.json "
                  f"to companies_config.json and fill in your company URLs.")
        return

    companies = json.loads(CONFIG_FILE.read_text())
    log.info(f"Loaded {len(companies)} companies from config.")
    log.info(f"Today's output folder: {TODAY_DIR}")

    # Companies are independent, I/O-bound (network) work, so fetch them
    # concurrently rather than one at a time. Log lines from different
    # companies will interleave in the log -- that's expected.
    matched_jobs = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_company, c): c["name"] for c in companies}
        for future in as_completed(futures):
            name = futures[future]
            try:
                _, _, matched = future.result()
                matched_jobs.extend(matched)
            except Exception as e:
                log.error(f"[{name}] Unhandled error during fetch/filter: {e}")

    # Cross-day dedup: drop anything already written on a previous day
    # before it ever reaches today's file.
    seen = load_seen_jobs()
    before = len(matched_jobs)
    new_jobs = [j for j in matched_jobs if j["link"] and j["link"] not in seen]
    skipped_repeats = before - len(new_jobs)
    if skipped_repeats:
        log.info(f"Skipped {skipped_repeats} job(s) already seen on a previous day.")

    added = write_results(new_jobs)

    seen.update(j["link"] for j in new_jobs)
    save_seen_jobs(seen)

    log.info(f"Done. {len(new_jobs)} new matching job(s) this run (of {before} total matched "
             f"before cross-day dedup), {added} row(s) written to {OUTPUT_FILE}.")


if __name__ == "__main__":
    main()