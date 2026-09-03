#!/usr/bin/env python3
"""
job_monitor.py — one-file job monitor + email notifier.

What it does, every time it runs:
  1. Reads config.json (your roles/keywords + the list of companies).
  2. Visits each company's career page and pulls out job-ish links.
     - Companies whose source_type is "greenhouse" or "lever" use those
       ATS's public JSON APIs (fast + reliable + full descriptions).
     - Everything else uses a conservative generic HTML scraper
       (title + link only — good enough to catch new postings, but it
       cannot read full job descriptions on JS-heavy sites like Google,
       Microsoft or Workday-based portals; see README.md "Known limits").
  3. Scores every job against your profile (roles/keywords/locations/
     experience, minus excluded terms).
  4. Compares against jobs.xlsx from the last run to find what's NEW or
     CHANGED, and skips anything already seen.
  5. Emails you a digest of the new/changed relevant jobs (only if there
     is something to report, unless send_empty_report is true).
  6. Saves the updated job list back to jobs.xlsx.

Run it directly:  python job_monitor.py
Nothing here needs to be edited by hand except config.json and the
environment variables described in README.md.
"""

import os
import re
import sys
import json
import time
import hashlib
import logging
import smtplib
import unicodedata
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from urllib.parse import urljoin

import requests
import pandas as pd
from bs4 import BeautifulSoup

# --------------------------------------------------------------------------
# CONSTANTS
# --------------------------------------------------------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
JOB_DATABASE_FILE = os.path.join(BASE_DIR, "jobs.xlsx")

REQUEST_TIMEOUT = 20
MAX_RETRIES = 2

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/125.0 Safari/537.36 JobMonitor/2.0"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

EXCEL_COLUMNS = [
    "Job_ID", "Company", "Job_Title", "Location", "Description_Snippet",
    "Job_URL", "Source_Type", "Discovered_Date", "Last_Seen_Date",
    "Content_Hash", "Relevance_Score", "Relevance_Reasons",
    "Visited", "Applied", "Notes",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("JobMonitor")

SESSION = requests.Session()
SESSION.headers.update(HTTP_HEADERS)


# --------------------------------------------------------------------------
# JOB RECORD
# --------------------------------------------------------------------------

@dataclass
class Job:
    company: str = ""
    job_title: str = ""
    location: str = ""
    description: str = ""
    job_url: str = ""
    source_type: str = ""
    job_id: str = field(default="", init=True)
    content_hash: str = ""
    relevance_score: float = 0.0
    relevance_reasons: str = ""


def normalize_text(value):
    if value is None:
        return ""
    value = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", value).strip()


def compute_identity(job: Job):
    key = "|".join([
        normalize_text(job.company).lower(),
        normalize_text(job.job_title).lower(),
        normalize_text(job.job_url).lower(),
    ])
    job.job_id = hashlib.sha256(key.encode("utf-8")).hexdigest()[:20]

    content = "|".join([
        normalize_text(job.job_title),
        normalize_text(job.location),
        normalize_text(job.description)[:500],
    ])
    job.content_hash = hashlib.sha256(content.encode("utf-8")).hexdigest()[:20]
    return job


# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------

def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


# --------------------------------------------------------------------------
# HTTP HELPER
# --------------------------------------------------------------------------

def http_get(url, params=None):
    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            resp.raise_for_status()
            return resp
        except requests.exceptions.RequestException as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
    raise last_error


# --------------------------------------------------------------------------
# CONNECTORS (one function per source_type -> list[Job])
# --------------------------------------------------------------------------

def fetch_greenhouse(company_cfg):
    """source_type='greenhouse', needs 'board_token' (from boards.greenhouse.io/<token>)."""
    token = company_cfg.get("board_token")
    if not token:
        raise ValueError("greenhouse source needs a 'board_token' in config.json")
    url = f"https://boards-api.greenhouse.io/v1/boards/{token}/jobs?content=true"
    data = http_get(url).json()
    jobs = []
    for item in data.get("jobs", []):
        desc = BeautifulSoup(item.get("content", ""), "lxml").get_text(" ", strip=True)
        jobs.append(Job(
            company=company_cfg["name"],
            job_title=item.get("title", ""),
            location=(item.get("location") or {}).get("name", ""),
            description=desc,
            job_url=item.get("absolute_url", ""),
            source_type="greenhouse",
        ))
    return jobs


def fetch_lever(company_cfg):
    """source_type='lever', needs 'lever_site' (from jobs.lever.co/<site>)."""
    site = company_cfg.get("lever_site")
    if not site:
        raise ValueError("lever source needs a 'lever_site' in config.json")
    url = f"https://api.lever.co/v0/postings/{site}?mode=json"
    data = http_get(url).json()
    jobs = []
    for item in data:
        categories = item.get("categories", {})
        jobs.append(Job(
            company=company_cfg["name"],
            job_title=item.get("text", ""),
            location=categories.get("location", ""),
            description=item.get("descriptionPlain", "") or item.get("description", ""),
            job_url=item.get("hostedUrl", ""),
            source_type="lever",
        ))
    return jobs


IGNORED_LINK_TERMS = [
    "home", "login", "sign in", "sign up", "privacy", "cookie", "contact us",
    "about us", "terms", "faq", "help", "accessibility", "sitemap",
]
JOB_LIKE_TERMS = [
    "job", "career", "engineer", "developer", "analyst", "scientist",
    "consultant", "intern", "trainee", "specialist", "manager", "associate",
    "programmer", "architect", "administrator", "designer",
]


def fetch_generic(company_cfg):
    """
    Best-effort connector for any plain career page. Works well for
    server-rendered pages; on JS-only single-page apps (common for large
    MNC portals) it may return few or zero results — that's a limitation
    of not running a real browser, not a bug. See README.md.
    """
    career_url = company_cfg["career_url"]
    resp = http_get(career_url)
    soup = BeautifulSoup(resp.text, "lxml")

    jobs = []
    seen_urls = set()

    for link in soup.find_all("a", href=True):
        title = link.get_text(" ", strip=True)
        href = link["href"].strip()
        if not title or not href or len(title) < 4 or len(title) > 140:
            continue

        title_lower = title.lower()
        if any(term in title_lower for term in IGNORED_LINK_TERMS):
            continue
        if not any(term in title_lower for term in JOB_LIKE_TERMS):
            continue

        absolute_url = urljoin(career_url, href)
        if absolute_url in seen_urls:
            continue
        seen_urls.add(absolute_url)

        jobs.append(Job(
            company=company_cfg["name"],
            job_title=title,
            location="",
            description="",
            job_url=absolute_url,
            source_type="generic",
        ))

    return jobs


def fetch_adzuna(company_cfg):
    """
    source_type='adzuna'. Aggregates real postings from Naukri, Indeed, Shine
    and thousands of company sites via Adzuna's free developer API.
    Needs ADZUNA_APP_ID and ADZUNA_APP_KEY env vars (free signup, no card
    needed): https://developer.adzuna.com/
    Config fields: 'query' (search term), 'location' (city), optional
    'country' (defaults to 'in' for India).
    """
    app_id = os.environ.get("ADZUNA_APP_ID")
    app_key = os.environ.get("ADZUNA_APP_KEY")
    if not app_id or not app_key:
        raise ValueError("ADZUNA_APP_ID / ADZUNA_APP_KEY env vars are not set")

    query = company_cfg.get("query", "")
    location = company_cfg.get("location", "")
    country = company_cfg.get("country", "in")

    url = f"https://api.adzuna.com/v1/api/jobs/{country}/search/1"
    params = {
        "app_id": app_id,
        "app_key": app_key,
        "what": query,
        "where": location,
        "results_per_page": 30,
        "content-type": "application/json",
    }
    data = http_get(url, params=params).json()

    jobs = []
    for item in data.get("results", []):
        jobs.append(Job(
            company=item.get("company", {}).get("display_name", "Unknown"),
            job_title=item.get("title", ""),
            location=item.get("location", {}).get("display_name", ""),
            description=item.get("description", ""),
            job_url=item.get("redirect_url", ""),
            source_type="adzuna",
        ))
    return jobs


def fetch_arbeitnow(company_cfg):
    """
    source_type='arbeitnow'. Free, no API key needed. Global tech-job board,
    good for remote roles. Config field: optional 'query' to filter by
    keyword client-side (the API itself doesn't support search params).
    """
    url = "https://www.arbeitnow.com/api/job-board-api"
    data = http_get(url).json()
    query = company_cfg.get("query", "").lower()

    jobs = []
    for item in data.get("data", []):
        title = item.get("title", "")
        if query and query not in title.lower():
            continue
        jobs.append(Job(
            company=item.get("company_name", "Unknown"),
            job_title=title,
            location=item.get("location", "") or ("Remote" if item.get("remote") else ""),
            description=item.get("description", ""),
            job_url=item.get("url", ""),
            source_type="arbeitnow",
        ))
    return jobs


SOURCE_CONNECTORS = {
    "greenhouse": fetch_greenhouse,
    "lever": fetch_lever,
    "generic": fetch_generic,
    "adzuna": fetch_adzuna,
    "arbeitnow": fetch_arbeitnow,
}


def discover_all_jobs(companies_cfg):
    all_jobs = []
    status = []

    for company_cfg in companies_cfg:
        if not company_cfg.get("enabled", True):
            continue

        name = company_cfg["name"]
        source_type = company_cfg.get("source_type", "generic")
        connector = SOURCE_CONNECTORS.get(source_type, fetch_generic)

        try:
            jobs = connector(company_cfg)
            all_jobs.extend(jobs)
            status.append((name, "OK", len(jobs), ""))
            logger.info(f"{name:30s} -> {len(jobs):3d} postings found")
        except Exception as e:
            status.append((name, "FAILED", 0, str(e)))
            logger.warning(f"{name:30s} -> FAILED ({e})")
            continue  # never let one bad company stop the run

    return all_jobs, status


# --------------------------------------------------------------------------
# RELEVANCE SCORING
# --------------------------------------------------------------------------

def clean_for_match(value):
    value = normalize_text(value).lower()
    return re.sub(r"[^a-z0-9+#./& -]", " ", value)


def contains_term(text, term):
    text = clean_for_match(text)
    term = clean_for_match(term)
    if not text or not term:
        return False
    if re.match(r"^[a-z0-9]+$", term.replace(" ", "")):
        return bool(re.search(rf"\b{re.escape(term)}\b", text))
    return term in text


EXPERIENCE_PATTERN = re.compile(
    r'(\d{1,2})\s*(?:\+|to|-)?\s*\d{0,2}\s*\+?\s*(?:years?|yrs?)\b',
    re.IGNORECASE,
)


def passes_mandatory_filters(job: Job, profile: dict):
    """
    Hard pass/fail gate. A job must clear ALL of these to be considered at
    all — none of this contributes to the score, it's strictly qualify /
    disqualify. Returns (passes: bool, reason: str).
    """
    title = job.job_title
    full_text = " ".join([job.job_title, job.location, job.description])

    # 1. Exclude keywords — checked across title + location + description,
    #    not just the title (this was the original bug: "5+ years" written
    #    in a description was invisible to a title-only check).
    for term in profile.get("exclude_keywords", []):
        if contains_term(full_text, term):
            return False, f"excluded (matched '{term}')"

    # 2. Numeric experience cutoff — catches "5+ years", "3-5 years" etc.
    #    even when that exact phrase isn't in your exclude_keywords list.
    #    Only fires when a number is actually present, so postings with no
    #    stated experience (common on scraped listings) aren't punished.
    max_years = profile.get("max_experience_years")
    if max_years is not None:
        for m in EXPERIENCE_PATTERN.finditer(full_text):
            lower_bound = int(m.group(1))
            if lower_bound > max_years:
                return False, f"excluded (needs {m.group(0).strip()}, above your {max_years}-year limit)"

    # 3. Role — job title must contain at least one of your target roles.
    roles = profile.get("job_roles", [])
    if roles and not any(contains_term(title, r) for r in roles):
        return False, "no matching role in title"

    # 4. Location — only enforced when the posting actually states a
    #    location; scraped postings with a blank location field aren't
    #    penalized for a connector limitation that isn't the job's fault.
    locations = profile.get("preferred_locations", [])
    if locations and job.location.strip() and not any(
        contains_term(full_text, loc) for loc in locations
    ):
        return False, "no matching location"

    return True, "passed all mandatory filters"


def score_job(job: Job, profile: dict):
    """
    Score reflects ONLY keyword (skill) matches, used purely to rank
    results that already passed passes_mandatory_filters(). Role, location
    and experience do not affect this number — they're gates, not points.
    """
    haystack = " ".join([job.job_title, job.location, job.description])
    matched = [kw for kw in profile.get("keywords", []) if contains_term(haystack, kw)]

    job.relevance_score = round(min(len(matched) * 10, 100), 1)
    job.relevance_reasons = (
        ", ".join(f"skill:{k}" for k in matched[:10]) if matched
        else "no specific skill keywords matched (still passed role/location/experience filters)"
    )
    return job


# --------------------------------------------------------------------------
# EXCEL PERSISTENCE + DEDUPE
# --------------------------------------------------------------------------

def load_database():
    if not os.path.exists(JOB_DATABASE_FILE):
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        df = pd.read_excel(JOB_DATABASE_FILE, engine="openpyxl")
    except Exception as e:
        logger.warning(f"Could not read {JOB_DATABASE_FILE}: {e}. Starting fresh.")
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    for col in EXCEL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


def classify_and_merge(scored_jobs, existing_df):
    """Returns (new_jobs, updated_jobs, merged_dataframe)."""
    existing_by_id = {
        str(row["Job_ID"]): row.to_dict()
        for _, row in existing_df.iterrows()
        if str(row.get("Job_ID", "")).strip()
    }

    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    new_jobs, updated_jobs, rows = [], [], []

    for job in scored_jobs:
        old = existing_by_id.get(job.job_id)
        row = {
            "Job_ID": job.job_id,
            "Company": job.company,
            "Job_Title": job.job_title,
            "Location": job.location,
            "Description_Snippet": (job.description or "")[:300],
            "Job_URL": job.job_url,
            "Source_Type": job.source_type,
            "Discovered_Date": old["Discovered_Date"] if old else now,
            "Last_Seen_Date": now,
            "Content_Hash": job.content_hash,
            "Relevance_Score": job.relevance_score,
            "Relevance_Reasons": job.relevance_reasons,
            "Visited": old.get("Visited", "No") if old else "No",
            "Applied": old.get("Applied", "No") if old else "No",
            "Notes": old.get("Notes", "") if old else "",
        }
        rows.append(row)

        if old is None:
            new_jobs.append(job)
        elif str(old.get("Content_Hash", "")) != job.content_hash:
            updated_jobs.append(job)

    merged_df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    return new_jobs, updated_jobs, merged_df


def save_database(df):
    df.to_excel(JOB_DATABASE_FILE, index=False, engine="openpyxl")
    logger.info(f"Saved {len(df)} job records -> {JOB_DATABASE_FILE}")


# --------------------------------------------------------------------------
# EMAIL
# --------------------------------------------------------------------------

def _html_escape(value):
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def build_email_body_plain(new_jobs, updated_jobs, run_status, profile_name):
    """Plain-text fallback for email clients that don't render HTML."""
    lines = [f"Job Monitor digest for {profile_name}", "=" * 50, ""]

    def section(title, jobs):
        out = [title, "-" * len(title), ""]
        for i, job in enumerate(sorted(jobs, key=lambda j: -j.relevance_score), 1):
            out += [
                f"{i}. [{job.relevance_score:.0f}] {job.company} - {job.job_title}",
                f"   Location: {job.location or 'Not specified'}",
                f"   Why: {job.relevance_reasons or 'n/a'}",
                f"   Apply: {job.job_url}",
                "",
            ]
        return out

    if new_jobs:
        lines += section("NEW MATCHING JOBS", new_jobs)
    if updated_jobs:
        lines += section("UPDATED MATCHING JOBS", updated_jobs)
    if not new_jobs and not updated_jobs:
        lines += ["No new or updated matching jobs this run.", ""]

    failed = [s for s in run_status if s[1] == "FAILED"]
    if failed:
        lines += ["", "Companies that could not be checked this run:", "-" * 40]
        lines += [f"  - {name}: {err}" for name, _, _, err in failed]

    return "\n".join(lines)


def build_email_body_html(new_jobs, updated_jobs, run_status, profile_name):
    """Clean HTML digest — this is what most inboxes (Gmail, Outlook, etc.) will show."""

    def job_card(job):
        return f"""
        <tr>
          <td style="padding:14px 16px;border:1px solid #e2e2e2;border-radius:8px;
                     display:block;margin-bottom:10px;">
            <div style="font-size:15px;font-weight:600;color:#111;">
              {_html_escape(job.job_title)}
            </div>
            <div style="font-size:13px;color:#555;margin-top:2px;">
              {_html_escape(job.company)} &middot; {_html_escape(job.location or "Location not specified")}
            </div>
            <div style="font-size:12px;color:#888;margin-top:6px;">
              Match score: <b>{job.relevance_score:.0f}</b> &middot; {_html_escape(job.relevance_reasons or "n/a")}
            </div>
            <div style="margin-top:10px;">
              <a href="{_html_escape(job.job_url)}"
                 style="display:inline-block;padding:8px 14px;background:#111;color:#fff;
                        text-decoration:none;border-radius:6px;font-size:13px;">
                View &amp; Apply
              </a>
            </div>
          </td>
        </tr>"""

    def section(title, jobs):
        cards = "".join(job_card(j) for j in sorted(jobs, key=lambda j: -j.relevance_score))
        return f"""
        <h3 style="font-size:14px;text-transform:uppercase;letter-spacing:0.04em;
                   color:#111;margin:28px 0 10px;">{_html_escape(title)}</h3>
        <table role="presentation" width="100%" style="border-collapse:separate;border-spacing:0 10px;">
          {cards}
        </table>"""

    body_sections = []
    if new_jobs:
        body_sections.append(section(f"New matching jobs ({len(new_jobs)})", new_jobs))
    if updated_jobs:
        body_sections.append(section(f"Updated matching jobs ({len(updated_jobs)})", updated_jobs))
    if not new_jobs and not updated_jobs:
        body_sections.append(
            '<p style="color:#555;font-size:14px;">No new or updated matching jobs this run.</p>'
        )

    failed = [s for s in run_status if s[1] == "FAILED"]
    failed_html = ""
    if failed:
        items = "".join(
            f'<li style="margin-bottom:4px;">{_html_escape(name)} '
            f'<span style="color:#999;">— {_html_escape(err[:120])}</span></li>'
            for name, _, _, err in failed
        )
        failed_html = f"""
        <details style="margin-top:26px;">
          <summary style="font-size:12px;color:#888;cursor:pointer;">
            {len(failed)} companies could not be checked this run
          </summary>
          <ul style="font-size:12px;color:#999;margin-top:8px;padding-left:18px;">
            {items}
          </ul>
        </details>"""

    return f"""\
<!DOCTYPE html>
<html>
  <body style="margin:0;padding:0;background:#f5f5f5;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;">
    <table role="presentation" width="100%" style="background:#f5f5f5;padding:24px 0;">
      <tr>
        <td align="center">
          <table role="presentation" width="600" style="background:#fff;border-radius:10px;
                 padding:28px 28px 20px;">
            <tr>
              <td>
                <div style="font-size:18px;font-weight:700;color:#111;">Job Monitor</div>
                <div style="font-size:13px;color:#888;margin-top:2px;">Daily digest for {_html_escape(profile_name)}</div>
                {"".join(body_sections)}
                {failed_html}
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>"""


def send_email(subject, plain_body, html_body):
    sender = os.environ.get("GMAIL_ADDRESS")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    recipient = os.environ.get("RECIPIENT_EMAIL", sender)
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    if not sender or not app_password:
        logger.error(
            "GMAIL_ADDRESS / GMAIL_APP_PASSWORD not set — skipping email. "
            "See README.md for how to set these."
        )
        return False

    msg = MIMEMultipart("alternative")
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    # Attach plain text first, HTML second — email clients render the last
    # part they understand, so HTML-capable clients show the pretty version
    # and plain-text-only clients fall back cleanly.
    msg.attach(MIMEText(plain_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender, app_password)
        server.sendmail(sender, [recipient], msg.as_string())

    logger.info(f"Email sent to {recipient}")
    return True


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------

def run():
    config = load_config()
    profile = config["profile"]
    email_cfg = config.get("email", {})

    logger.info(f"Job Monitor starting for profile: {profile['name']}")

    existing_df = load_database()
    logger.info(f"Existing records: {len(existing_df)}")

    discovered_jobs, run_status = discover_all_jobs(config["companies"])
    logger.info(f"Total postings discovered: {len(discovered_jobs)}")

    # de-duplicate within this run + compute identity
    unique = {}
    for job in discovered_jobs:
        compute_identity(job)
        unique[job.job_id] = job
    discovered_jobs = list(unique.values())

    # score
    for job in discovered_jobs:
        score_job(job, profile)

    # mandatory gates: role, location, experience, exclusions — pass/fail only
    qualified_jobs = []
    for job in discovered_jobs:
        ok, reason = passes_mandatory_filters(job, profile)
        if ok:
            qualified_jobs.append(job)
    logger.info(f"Passed mandatory filters (role/location/experience/exclude): {len(qualified_jobs)}")

    relevant_jobs = [
        j for j in qualified_jobs
        if j.relevance_score >= profile.get("min_relevance_score", 30)
    ]
    logger.info(f"Relevant postings (keyword score >= {profile.get('min_relevance_score', 30)}): {len(relevant_jobs)}")

    new_jobs, updated_jobs, merged_df = classify_and_merge(relevant_jobs, existing_df)
    save_database(merged_df)

    should_send = bool(new_jobs or updated_jobs) or email_cfg.get("send_empty_report", False)
    if should_send:
        subject = f"Job Monitor: {len(new_jobs)} new, {len(updated_jobs)} updated"
        plain_body = build_email_body_plain(new_jobs, updated_jobs, run_status, profile["name"])
        html_body = build_email_body_html(new_jobs, updated_jobs, run_status, profile["name"])
        send_email(subject, plain_body, html_body)
    else:
        logger.info("Nothing new to report — no email sent.")

    logger.info("Run complete.")


if __name__ == "__main__":
    run()
