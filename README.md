# Job Monitor — automated job-alert emails

Checks ~40 company career pages every 12 hours, scores each posting against
your roles/keywords, and emails you only the new or changed matches.
Runs for free on GitHub's servers — your laptop doesn't need to be on.

## One-time setup (5–10 minutes)

**1. Get a Gmail "App Password" (not your normal password)**
   - Turn on 2-Step Verification on your Google account, if it isn't already:
     https://myaccount.google.com/security
   - Create an app password: https://myaccount.google.com/apppasswords
   - Copy the 16-character password it gives you.

**2. Put this project in a GitHub repo**
   - Create a new **private** repo on GitHub and push these files to it
     (or use GitHub's "upload files" button in the browser — no git needed).

**3. Add your secrets** (Settings → Secrets and variables → Actions → New repository secret)
   | Secret name | Value |
   |---|---|
   | `GMAIL_ADDRESS` | the Gmail address you made the app password for |
   | `GMAIL_APP_PASSWORD` | the 16-character app password |
   | `RECIPIENT_EMAIL` | where you want alerts sent (can be the same address) |

**4. Edit `config.json`** — this is the only file you should need to touch:
   - `profile.name` — your name (shows in the email subject/body)
   - `profile.job_roles` — job titles you're targeting
   - `profile.keywords` — skills/tech that should count in your favor
   - `profile.preferred_locations` — cities or "Remote"
   - `profile.exclude_keywords` — anything that should disqualify a posting
     (e.g. "senior", "5+ years")
   - `profile.min_relevance_score` — raise this (out of 100) for fewer,
     more targeted emails; lower it to see more borderline matches

**5. Turn it on**
   - Go to the **Actions** tab of your repo → enable workflows if prompted.
   - Click **"Job Monitor" → "Run workflow"** to test it right now.
   - After that it runs automatically every day (see the `cron` line in
     `.github/workflows/job-monitor.yml` to change the time — it's in UTC).

That's it — it's now a genuine "one click" system: click **Run workflow**
any time you want an on-demand check, otherwise it just emails you daily.

## Giving this to friends

Each friend should **make their own copy of the repo** (GitHub's "Use this
template" or "Fork" button), then only needs to:
1. Add their own three secrets (step 3 above), using their own Gmail app password.
2. Edit `profile` in `config.json` with their own name/roles/keywords.

They should **not** reuse your secrets or edit your repo directly — each
person's copy keeps its own `jobs.xlsx` history and sends to their own inbox.

## Running it locally instead (optional)

If you'd rather run it on your own machine/cron job instead of GitHub Actions:

```bash
pip install -r requirements.txt
export GMAIL_ADDRESS="you@gmail.com"
export GMAIL_APP_PASSWORD="xxxxxxxxxxxxxxxx"
export RECIPIENT_EMAIL="you@gmail.com"
python job_monitor.py
```

Add that to cron (Linux/Mac) or Task Scheduler (Windows) to automate it
without GitHub.

## How matching works

Every posting gets points for: matching one of your `job_roles` (+35),
each matching `keywords` term (+6), each matching `preferred_locations`
(+10), each matching `experience_keywords` (+8) — capped at 100. Any
`exclude_keywords` match in the title disqualifies it outright. Only
postings scoring at or above `min_relevance_score` are ever emailed or
saved. `jobs.xlsx` is the running history — it's how the script knows
what's *new* vs. what it already told you about.

## Known limits (please read before relying on this)

- **Company sites, not a single job board.** Each company's own career
  page has a different structure. This script uses:
  - **Greenhouse / Lever** connectors for companies on those ATSs — these
    use official public JSON APIs and are reliable, with full descriptions.
    (None of the 40 starter companies are pre-configured this way — see
    "Adding a Greenhouse/Lever company" below.)
  - A **generic** connector for everything else, which reads the visible
    HTML of the career page. This works fine on simpler, server-rendered
    pages, but **large companies whose career sites are JavaScript
    single-page apps (Google, Microsoft, Amazon, most Workday-based
    portals) will often return few or zero results**, because this script
    doesn't run a real browser. That's a real limitation, not a bug — a
    browser-automation connector (e.g. Playwright) could be added later
    for those specific sites if it becomes worth the added complexity/cost.
- **No login, no CAPTCHA bypass, no rate-limit evasion.** The script only
  reads public pages, politely, with retries — nothing that requires an
  account.
- Treat this as a **first-pass filter**, not a substitute for periodically
  checking the career pages yourself, especially for the big-tech names.

### Adding a Greenhouse/Lever company (better accuracy)

If you find a company's career page is actually hosted on Greenhouse
(URL like `boards.greenhouse.io/<token>`) or Lever
(`jobs.lever.co/<site>`), change its entry in `config.json`, e.g.:

```json
{"name": "Example Co", "source_type": "greenhouse", "board_token": "examplecoco", "enabled": true}
```
```json
{"name": "Example Co", "source_type": "lever", "lever_site": "examplecoco", "enabled": true}
```

## Files in this project

| File | Purpose |
|---|---|
| `config.json` | Your profile + the company list — the only file you normally edit |
| `job_monitor.py` | The whole program |
| `jobs.xlsx` | Auto-generated history of everything seen so far (dedupe database) |
| `requirements.txt` | Python dependencies |
| `.github/workflows/job-monitor.yml` | Daily automation + manual "Run workflow" button |
