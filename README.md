# Job Monitor — automated job-alert emails (multi-person)

Checks career pages + Adzuna/Arbeitnow every day, scores postings against
each person's own roles/keywords, and emails only new matches. One repo
can now run this for **you and any number of friends** — they don't set
up anything; they just get a config file with their name on it.

## How it works

- `configs/` holds one JSON file per person (their roles, keywords,
  locations, experience limit, and their email address).
- `data/` holds one history file per person, auto-created, used to avoid
  re-emailing the same posting twice.
- The GitHub Action loops over every file in `configs/` and runs the
  whole pipeline once per person, in a single workflow run.
- Only **you** (the repo owner) hold the shared secrets — your Gmail
  sending account and Adzuna API keys. Friends never touch secrets,
  never fork anything, never see the repo at all if you don't want them to.

## One-time setup (you, the repo owner)

**1. Get a Gmail "App Password"**
   - Turn on 2-Step Verification: https://myaccount.google.com/security
   - Create an app password: https://myaccount.google.com/apppasswords

**2. Get free Adzuna API keys** (optional but recommended — much cleaner
   results than career-page scraping): https://developer.adzuna.com/

**3. Add repo secrets** (Settings → Secrets and variables → Actions):

   | Secret name | Value |
   |---|---|
   | `GMAIL_ADDRESS` | your Gmail address |
   | `GMAIL_APP_PASSWORD` | the 16-character app password |
   | `ADZUNA_APP_ID` | from the Adzuna dashboard |
   | `ADZUNA_APP_KEY` | from the Adzuna dashboard |

   Notice there's no `RECIPIENT_EMAIL` secret anymore — each person's
   destination address now lives in *their own* config file instead,
   since an email address isn't a credential.

**4. Turn it on** — Actions tab → enable workflows if prompted → "Job
   Monitor" → "Run workflow" to test.

## Adding a friend (takes you two minutes, they do nothing)

1. Copy `configs/_template.json` to `configs/<their-name>.json`.
2. Fill in their `job_roles`, `keywords`, `preferred_locations`,
   `max_experience_years`, `exclude_keywords`, and `email.recipient_email`
   — base this on their resume/target roles.
3. Commit it.

That's it — next run, they're included automatically, with their own
separate history file so their matches never mix with anyone else's.
To stop alerting someone, just delete their file from `configs/`.

## How matching works

Every posting must pass ALL of these to be considered at all (hard
filters, not scoring):
- Doesn't contain anything in `exclude_keywords` (title, description, or
  location — checked everywhere, not just the title)
- Doesn't require more than `max_experience_years` (a number-detecting
  filter catches phrasing like "5+ years" or "3-5 years" even if that
  exact phrase isn't in your `exclude_keywords` list)
- Job title contains at least one of `job_roles`
- Location matches one of `preferred_locations` (only enforced when the
  posting actually states a location)

Only jobs that pass all of the above get a **score**, and that score is
based purely on how many `keywords` (skills) it matches — used only to
rank/sort, not to decide inclusion. `min_relevance_score` is the minimum
keyword score to bother reporting.

## Known limits (please read)

- **Company career pages, not one unified job board.** Large companies
  running JavaScript single-page apps (Google, Microsoft, Amazon, most
  Workday portals) often return few or zero results from a plain HTTP
  scraper — that's a structural limitation, not a bug.
- **Adzuna/Arbeitnow are the more reliable sources** — they're real
  aggregator APIs (not scraped), with full descriptions, so filters work
  best on those. Consider leaning on them more than the `generic`
  company connectors if accuracy matters more than coverage.
- Some `generic` company sites will show up as "FAILED" in the run log
  (403 Forbidden, DNS errors, bot detection) — this is expected for a
  handful of sites and doesn't affect the others.
- Treat this as a first-pass filter, not a replacement for occasionally
  checking the big-tech career pages yourself.

## Running it locally instead (optional)

```bash
pip install -r requirements.txt
export GMAIL_ADDRESS="you@gmail.com"
export GMAIL_APP_PASSWORD="xxxxxxxxxxxxxxxx"
export ADZUNA_APP_ID="..."
export ADZUNA_APP_KEY="..."
python job_monitor.py configs/vignesh.json
```

Run it once per config file to check everyone, or loop over `configs/*.json`
like the GitHub Action does.

## Files in this project

| File/folder | Purpose |
|---|---|
| `configs/*.json` | One file per person — the only thing that changes when adding/editing someone |
| `configs/_template.json` | Blank starting point for a new person (ignored by the workflow) |
| `data/*_jobs.xlsx` | Auto-generated per-person history/dedupe database |
| `job_monitor.py` | The whole program, takes a config path as its one argument |
| `requirements.txt` | Python dependencies |
| `.github/workflows/job-monitor.yml` | Daily automation, loops over every config, one commit at the end |
