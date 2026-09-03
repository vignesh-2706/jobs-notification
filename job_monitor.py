#!/usr/bin/env python3
"""
job_monitor.py — one-file job monitor + email notifier, multi-person.

What it does, every time it runs:
  1. Reads one config file (a person's roles/keywords/companies/recipient
     email) from the configs/ folder.
  2. Visits each company's career page / API and pulls out job-ish links.
  3. Scores every job against that person's profile (hard filters for
     role/location/experience/exclusions; score reflects keyword matches).
  4. Compares against that person's own history file in data/ to find
     what's NEW, and skips anything already seen.
  5. Emails that person a digest of new relevant jobs.
  6. Saves the updated job list back to their history file.

Run it for one person directly:
    python job_monitor.py configs/vignesh.json

Run it for everyone (what the GitHub Action does):
    for f in configs/*.json; do python job_monitor.py "$f"; done

Only the configs/*.json files need editing per-person (roles, keywords,
recipient email). GMAIL_ADDRESS / GMAIL_APP_PASSWORD / ADZUNA_* stay as a
single shared set of repo secrets — friends never touch those.
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
from dataclasses import dataclass, field
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
DATA_DIR = os.path.join(BASE_DIR, "data")
DEFAULT_CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

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

def load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def database_path_for(config_path):
    """Each person gets their own history file, named after their config,
    so dedupe never mixes across people: configs/vignesh.json ->
    data/vignesh_jobs.xlsx"""
    stem = os.path.splitext(os.path.basename(config_path))[0]
    os.makedirs(DATA_DIR, exist_ok=True)
    return os.path.join(DATA_DIR, f"{stem}_jobs.xlsx")


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
#
# Every connector now takes (company_cfg, profile) so that source types
# which search by keyword (adzuna, jooble) can loop over the person's own
# job_roles list instead of needing one config entry per role. That means
# configs/*.json only needs ONE "Adzuna - Bangalore" style entry per
# location, not one per role — the roles already live in profile.job_roles
# and this is the single place that reads them for querying.
# --------------------------------------------------------------------------

def fetch_greenhouse(company_cfg, profile):
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


def fetch_lever(company_cfg, profile):
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


MAX_DETAIL_FETCHES_PER_COMPANY = 25  # keeps runtime sane even if a page has many links


def fetch_generic(company_cfg, profile):
    """
    Best-effort connector for any plain career page. Works well for
    server-rendered pages; on JS-only single-page apps (common for large
    MNC portals) it may return few or zero results — that's a limitation
    of not running a real browser, not a bug. See README.md.

    Also fetches each individual posting's page (capped at
    MAX_DETAIL_FETCHES_PER_COMPANY) to pull real description text, so the
    experience-year filter and keyword scoring have something to check
    besides a bare title.
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

    # Fetch real page content for each candidate. Any fetch failure is
    # silently tolerated — that job just keeps its empty description.
    for job in jobs[:MAX_DETAIL_FETCHES_PER_COMPANY]:
        try:
            detail_resp = http_get(job.job_url)
            job.description = BeautifulSoup(detail_resp.text, "lxml").get_text(" ", strip=True)
        except Exception:
            pass

    return jobs


def fetch_adzuna(company_cfg, profile):
    """
    source_type='adzuna'. Aggregates real postings from Naukri, Indeed, Shine
    and thousands of company sites via Adzuna's free developer API.
    Needs ADZUNA_APP_ID and ADZUNA_APP_KEY env vars (free signup, no card
    needed): https://developer.adzuna.com/

    Config fields: 'location' (city), optional 'country' (defaults to
    'in'). Queries are built from profile.job_roles automatically — one
    API call per role — so you only need one company entry per location,
    not one per role.
    """
    app_id = os.environ.get("ADZUNA_APP_ID")
    app_key = os.environ.get("ADZUNA_APP_KEY")
    if not app_id or not app_key:
        raise ValueError("ADZUNA_APP_ID / ADZUNA_APP_KEY env vars are not set")

    location = company_cfg.get("location", "")
    country = company_cfg.get("country", "in")
    roles = profile.get("job_roles") or [company_cfg.get("query", "")]

    jobs = []
    seen_urls = set()
    for role in roles:
        url = f"https://api.adzuna.com/v1/api/jobs/{country}/search/1"
        params = {
            "app_id": app_id,
            "app_key": app_key,
            "what": role,
            "where": location,
            "results_per_page": 30,
            "content-type": "application/json",
        }
        try:
            data = http_get(url, params=params).json()
        except Exception as e:
            logger.warning(f"  Adzuna query '{role}' failed: {e}")
            continue

        for item in data.get("results", []):
            job_url = item.get("redirect_url", "")
            if job_url in seen_urls:
                continue
            seen_urls.add(job_url)
            jobs.append(Job(
                company=item.get("company", {}).get("display_name", "Unknown"),
                job_title=item.get("title", ""),
                location=item.get("location", {}).get("display_name", ""),
                description=item.get("description", ""),
                job_url=job_url,
                source_type="adzuna",
            ))
    return jobs


def fetch_arbeitnow(company_cfg, profile):
    """
    source_type='arbeitnow'. Free, no API key needed. Global tech-job board,
    good for remote roles. Pulls the full board — role matching happens
    later in passes_mandatory_filters() against profile.job_roles, so no
    per-role config or query field is needed here.
    """
    url = "https://www.arbeitnow.com/api/job-board-api"
    data = http_get(url).json()

    jobs = []
    for item in data.get("data", []):
        jobs.append(Job(
            company=item.get("company_name", "Unknown"),
            job_title=item.get("title", ""),
            location=item.get("location", "") or ("Remote" if item.get("remote") else ""),
            description=item.get("description", ""),
            job_url=item.get("url", ""),
            source_type="arbeitnow",
        ))
    return jobs


def fetch_jooble(company_cfg, profile):
    """
    source_type='jooble'. Aggregates from Naukri, Indeed, TimesJobs, and
    thousands of company sites — strong India coverage, similar role to
    Adzuna but a different underlying index (worth having both).
    Needs JOOBLE_API_KEY env var (free signup, no card needed):
    https://jooble.org/api/about

    Config field: 'location' (city — leave blank or use a broad value like
    "India" if city-level matching returns 0). Like Adzuna, queries are
    built from profile.job_roles — one call per role, one config entry per
    location.
    """
    api_key = os.environ.get("JOOBLE_API_KEY")
    if not api_key:
        raise ValueError("JOOBLE_API_KEY env var is not set")

    location = company_cfg.get("location", "")
    roles = profile.get("job_roles") or [company_cfg.get("query", "")]
    url = f"https://jooble.org/api/{api_key}"

    jobs = []
    seen_urls = set()
    for role in roles:
        payload = {"keywords": role, "location": location}
        try:
            resp = SESSION.post(url, json=payload, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            logger.warning(f"  Jooble query '{role}' failed: {e}")
            continue

        if "errorMessage" in data:
            logger.warning(f"Jooble API error for '{role}': {data['errorMessage']}")

        for item in data.get("jobs", []):
            link = item.get("link", "")
            if link in seen_urls:
                continue
            seen_urls.add(link)
            jobs.append(Job(
                company=item.get("company", "Unknown"),
                job_title=item.get("title", ""),
                location=item.get("location", ""),
                description=item.get("snippet", ""),
                job_url=link,
                source_type="jooble",
            ))
    return jobs


def fetch_remoteok(company_cfg, profile):
    """
    source_type='remoteok'. Free, no API key needed. Remote-first tech
    jobs board. Pulls the full board — role matching happens later in
    passes_mandatory_filters(), so no query field is needed here.

    Known issue: RemoteOK sometimes rate-limits/blocks requests from cloud
    datacenter IPs (which is what GitHub Actions runners use). If the
    diagnostic log below consistently shows raw_items=1, that's why.
    """
    url = "https://remoteok.com/api"
    data = http_get(url).json()

    logger.info(f"  [diagnostic] RemoteOK raw_items={len(data)}")

    jobs = []
    for item in data:
        title = item.get("position") or item.get("title", "")
        if not title:
            continue  # first element of RemoteOK's response is a legal notice, not a job
        jobs.append(Job(
            company=item.get("company", "Unknown"),
            job_title=title,
            location=item.get("location", "") or "Remote",
            description=item.get("description", ""),
            job_url=item.get("url", ""),
            source_type="remoteok",
        ))
    return jobs


SOURCE_CONNECTORS = {
    "greenhouse": fetch_greenhouse,
    "lever": fetch_lever,
    "generic": fetch_generic,
    "adzuna": fetch_adzuna,
    "arbeitnow": fetch_arbeitnow,
    "jooble": fetch_jooble,
    "remoteok": fetch_remoteok,
}


def discover_all_jobs(companies_cfg, profile):
    all_jobs = []
    status = []

    for company_cfg in companies_cfg:
        if not company_cfg.get("enabled", True):
            continue

        name = company_cfg["name"]
        source_type = company_cfg.get("source_type", "generic")
        connector = SOURCE_CONNECTORS.get(source_type, fetch_generic)

        try:
            jobs = connector(company_cfg, profile)
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

# Catches "<Any Role> II", "<Any Role> III", "<Any Role> 2", "<Any Role> 3",
# etc. at the END of a job title — the generic pattern behind seniority
# suffixes like "Software Engineer II", "Site Reliability Engineer II",
# "Software Development Engineer III" — WITHOUT needing to hand-list every
# possible role name it could be attached to. "I" and "1" are deliberately
# NOT matched here since those usually denote entry-level, not senior.
LEVEL_SUFFIX_PATTERN = re.compile(
    r'\b(?:ii|iii|iv|v|vi|2|3|4|5|6)\s*$',
    re.IGNORECASE,
)


def passes_mandatory_filters(job: Job, profile: dict):
    """
    Hard pass/fail gate. A job must clear ALL of these to be considered at
    all — none of this contributes to the score, it's strictly qualify /
    disqualify. Returns (passes: bool, reason: str).

    Policy: if experience isn't mentioned anywhere in the title/location/
    description, that's treated as a PASS (unstated is not the same as
    disqualifying) — only an EXPLICIT requirement above your limit
    excludes a job.
    """
    title = job.job_title
    full_text = " ".join([job.job_title, job.location, job.description])

    # 1a. Numeric experience cutoff — catches "5+ years", "3-5 years" etc.
    max_years = profile.get("max_experience_years")
    if max_years is not None:
        for m in EXPERIENCE_PATTERN.finditer(full_text):
            lower_bound = int(m.group(1))
            if lower_bound > max_years:
                return False, f"excluded (needs {m.group(0).strip()}, above your {max_years}-year limit)"

    # 1b. Title-level suffix — catches "<Role> II", "<Role> III" etc.
    if LEVEL_SUFFIX_PATTERN.search(title.strip()):
        return False, f"excluded (title suggests a senior level: '{title.strip()}')"

    # 2. Exclude keywords — checked across title + location + description.
    for term in profile.get("exclude_keywords", []):
        if contains_term(full_text, term):
            return False, f"excluded (matched '{term}')"

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


def deep_verify_experience(job: Job, profile: dict):
    """
    Second-pass, stricter experience check — run ONLY on the small final
    shortlist (jobs that already passed every other gate and are about to
    be emailed), not on every discovered posting. Fetches the real posting
    page when possible and re-runs the same regex against the full text.

    Skipped for adzuna/jooble, whose job_url is a tracking/redirect link
    that returns 403 to non-browser requests — the earlier check against
    their API description field is the best available signal there.

    Fails open: if the page can't be fetched, the job is kept rather than
    dropped — we only exclude on a positive, explicit match.
    Returns (keep: bool, reason: str).
    """
    max_years = profile.get("max_experience_years")
    if max_years is None or not job.job_url:
        return True, "no max_experience_years set or no URL to verify"

    NOT_DIRECTLY_FETCHABLE = {"adzuna", "jooble"}
    if job.source_type in NOT_DIRECTLY_FETCHABLE:
        return True, f"{job.source_type} URLs are tracking redirects, not fetchable"

    try:
        resp = http_get(job.job_url)
        full_text = BeautifulSoup(resp.text, "lxml").get_text(" ", strip=True)
    except Exception as e:
        logger.info(f"  Deep-check skipped for '{job.job_title}' @ {job.company} (couldn't fetch page: {e})")
        return True, "fetch failed, kept"

    for m in EXPERIENCE_PATTERN.finditer(full_text):
        lower_bound = int(m.group(1))
        if lower_bound > max_years:
            reason = f"full posting requires '{m.group(0).strip()}', above your {max_years}-year limit"
            logger.info(f"  Deep-check EXCLUDED '{job.job_title}' @ {job.company}: {reason}")
            return False, reason

    return True, "verified against full posting, within limit"


# --------------------------------------------------------------------------
# EXCEL PERSISTENCE + DEDUPE
# --------------------------------------------------------------------------

def load_database(db_path):
    if not os.path.exists(db_path):
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        df = pd.read_excel(db_path, engine="openpyxl")
    except Exception as e:
        logger.warning(f"Could not read {db_path}: {e}. Starting fresh.")
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


def save_database(df, db_path):
    df.to_excel(db_path, index=False, engine="openpyxl")
    logger.info(f"Saved {len(df)} job records -> {db_path}")


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


def send_email(subject, plain_body, html_body, recipient):
    sender = os.environ.get("GMAIL_ADDRESS")
    app_password = os.environ.get("GMAIL_APP_PASSWORD")
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))

    if not sender or not app_password:
        logger.error(
            "GMAIL_ADDRESS / GMAIL_APP_PASSWORD not set — skipping email. "
            "See README.md for how to set these."
        )
        return False
    if not recipient:
        logger.error("No recipient email configured for this profile — skipping email.")
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

def run(config_path):
    config = load_config(config_path)
    profile = config["profile"]
    email_cfg = config.get("email", {})
    db_path = database_path_for(config_path)

    # Each person's own email address lives in THEIR config file (not a
    # secret — it's just an address, not a credential). Falls back to
    # RECIPIENT_EMAIL env var for backward compatibility with single-user
    # setups that haven't added this field yet.
    recipient = email_cfg.get("recipient_email") or os.environ.get("RECIPIENT_EMAIL")

    logger.info(f"Job Monitor starting for profile: {profile['name']}  (config: {config_path})")

    existing_df = load_database(db_path)
    logger.info(f"Existing records: {len(existing_df)}")

    discovered_jobs, run_status = discover_all_jobs(config["companies"], profile)
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

    # Deep experience re-check — only on this small final shortlist, since
    # it fetches each job's actual page (not just the aggregator snippet)
    # to catch experience requirements the short teaser text missed.
    if profile.get("max_experience_years") is not None and relevant_jobs:
        logger.info(f"Deep-verifying experience on {len(relevant_jobs)} shortlisted job(s) against full postings...")
        verified_jobs = []
        for job in relevant_jobs:
            keep, reason = deep_verify_experience(job, profile)
            if keep:
                verified_jobs.append(job)
        logger.info(f"Passed deep experience verification: {len(verified_jobs)} (of {len(relevant_jobs)})")
        relevant_jobs = verified_jobs

    new_jobs, updated_jobs, merged_df = classify_and_merge(relevant_jobs, existing_df)
    save_database(merged_df, db_path)

    should_send = bool(new_jobs or updated_jobs) or email_cfg.get("send_empty_report", False)
    if should_send:
        subject = f"Job Monitor: {len(new_jobs)} new, {len(updated_jobs)} updated"
        plain_body = build_email_body_plain(new_jobs, updated_jobs, run_status, profile["name"])
        html_body = build_email_body_html(new_jobs, updated_jobs, run_status, profile["name"])
        send_email(subject, plain_body, html_body, recipient)
    else:
        logger.info("Nothing new to report — no email sent.")

    logger.info(f"Run complete for {profile['name']}.")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CONFIG_FILE
    run(target)
