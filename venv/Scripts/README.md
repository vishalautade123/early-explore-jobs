# Daily Job Search Automation

Scrapes your target companies' career pages every morning, keeps roles
posted in the last 15 days matching your keywords (Full Stack /
React / Spring Boot / Backend), and writes them into
`job_search_results.xlsx`. Safe to re-run daily — it won't duplicate rows.

## 1. Setup (one-time)

```bash
cd job_scraper
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

pip install -r requirements.txt
```

Copy the example config and fill in your real companies:
```bash
cp companies_config.example.json companies_config.json
```

## 2. Building your 40-50 company list

For each company's career page, you need to figure out which ATS
(Applicant Tracking System) they use — this determines whether the
script can pull clean, reliable data.

**How to check a company's career page:**
1. Open their "Careers" / "Jobs" page in your browser.
2. Look at the URL once you click into job listings, or open browser
   DevTools → Network tab → reload the page and look for requests to:
   - `boards-api.greenhouse.io` or a URL containing `greenhouse.io` → **Greenhouse**
   - `api.lever.co` or `jobs.lever.co` → **Lever**
   - `api.ashbyhq.com` or `jobs.ashbyhq.com` → **Ashby**
   - `myworkdayjobs.com` → **Workday** (JS-rendered, harder — see note below)
   - anything else → **Custom/generic** (needs manual handling)

Many mid-size tech companies and startups use Greenhouse or Lever — those
are the easiest and most reliable to pull from directly via their public
JSON APIs (no scraping fragility, they're built for this).

3. In `companies_config.json`, add an entry:
   ```json
   { "name": "CompanyName", "url": "https://boards.greenhouse.io/companyname" }
   ```
   The script auto-detects the ATS and slug from the URL, so for
   Greenhouse/Lever/Ashby you usually just need `name` and `url`.

**For custom/Workday sites:** the generic fetcher currently returns
nothing — these need JS rendering (Playwright) and page-specific CSS
selectors, since every custom site is laid out differently. My suggestion:
build your list with Greenhouse/Lever/Ashby companies first (should cover
a good chunk of 40-50 targets), run it for a week, then we add Playwright
support and custom selectors for the remaining handful of must-have
companies that use something else. Trying to genericize scraping for
arbitrary custom HTML upfront usually wastes time on pages that change
their layout anyway.

## 3. Run it

```bash
python main.py
```

Check `job_search_results.xlsx` and `scraper.log` (errors per company
are logged there, so you can see which companies failed to fetch).

## 4. Schedule it to run every morning

### Windows (Task Scheduler)
1. Open Task Scheduler → Create Basic Task.
2. Name it "Job Search Scraper", trigger: Daily, set your preferred time.
3. Action: "Start a program".
   - Program/script: `C:\path\to\job_scraper\venv\Scripts\python.exe`
   - Add arguments: `main.py`
   - Start in: `C:\path\to\job_scraper`
4. Finish. It'll now run daily even if you're not logged in (check "Run
   whether user is logged on or not" in the task's Properties if needed).

### Mac/Linux (cron)
```bash
crontab -e
```
Add a line (runs daily at 8:00 AM):
```
0 8 * * * cd /path/to/job_scraper && venv/bin/python main.py
```

## 5. Tuning

- `DAYS_THRESHOLD` in `main.py` — change from 15 if you want a different window.
- `KEYWORDS` in `main.py` — add/remove terms to widen or narrow role matching.
- Jobs with an unknown posted date are included by default (rather than
  silently dropped) so you can eyeball them — remove that behavior in
  `within_date_range()` if you'd rather they be excluded.
