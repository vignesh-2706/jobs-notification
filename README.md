````markdown
# Job Monitor

> Automated, configurable job discovery and email notification pipeline built with Python and GitHub Actions.

Job Monitor is a lightweight automation system that collects job postings from multiple sources, filters them against configurable candidate preferences, ranks relevant opportunities, and sends email alerts only for new matches.

The system is designed to support **multiple candidates from a single repository**, with each candidate having an independent configuration and job-history database.

---

## ✨ Overview

Searching for jobs across multiple company career pages and job boards can be repetitive and time-consuming.

This project automates that workflow:

```text
                    ┌──────────────────────┐
                    │ Candidate Config     │
                    │ roles / skills /      │
                    │ locations / experience│
                    └──────────┬───────────┘
                               │
                               ▼
┌────────────────┐     ┌──────────────────────┐
│ Career Pages   │────▶│                      │
└────────────────┘     │    Job Monitor       │
                       │                      │
┌────────────────┐     │  Collect → Filter →  │
│ Adzuna API     │────▶│  Score → Deduplicate │
└────────────────┘     └──────────┬───────────┘
                                  │
┌────────────────┐                │
│ Arbeitnow API  │────────────────┘
└────────────────┘                │
                                  ▼
                         ┌─────────────────┐
                         │ Email Notification│
                         └─────────────────┘

              GitHub Actions runs the pipeline daily
````

The system can run automatically every day using **GitHub Actions**, without requiring a local machine or server to remain running.

---

## 🚀 Key Features

* **Multi-source job collection**

  * Company career pages
  * Adzuna API
  * Arbeitnow API

* **Configurable candidate profiles**

  * Target job roles
  * Skills / keywords
  * Preferred locations
  * Maximum experience level
  * Excluded keywords
  * Notification email

* **Rule-based job filtering**

  * Role matching
  * Location matching
  * Experience filtering
  * Exclusion rules

* **Keyword-based relevance scoring**

  * Jobs that pass the hard filters are ranked based on matching skills and keywords.

* **Duplicate prevention**

  * Previously processed jobs are stored in a per-candidate history file.
  * The same posting is not repeatedly emailed.

* **Automated execution**

  * GitHub Actions runs the complete pipeline on a scheduled basis.

* **Multi-person support**

  * One repository can manage multiple independent candidate profiles.
  * Each candidate has their own configuration and job history.

* **Centralized secret management**

  * API credentials and the email sender credentials are stored as GitHub Actions secrets.
  * Candidate configuration does not contain API credentials.

---

## 🏗️ Architecture

The application follows a simple pipeline-oriented architecture:

```text
                    GitHub Actions
                          │
                          ▼
                 ┌─────────────────┐
                 │ Load candidate  │
                 │ configuration   │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Job Collectors  │
                 │                 │
                 │ • Career Pages  │
                 │ • Adzuna        │
                 │ • Arbeitnow     │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Normalize Jobs  │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Hard Filtering  │
                 │                 │
                 │ Role            │
                 │ Location        │
                 │ Experience      │
                 │ Exclusions      │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Relevance Score │
                 │ Keyword matches │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Deduplication   │
                 │ Job history     │
                 └────────┬────────┘
                          │
                          ▼
                 ┌─────────────────┐
                 │ Email Alerts    │
                 └─────────────────┘
```

---

## 🔎 Matching Strategy

The matching pipeline intentionally separates **hard filtering** from **relevance scoring**.

### 1. Hard Filters

A job must first satisfy all required conditions.

#### Role

The job title must contain at least one configured target role.

Example:

```json
"job_roles": [
  "Software Engineer",
  "Backend Engineer",
  "Python Developer"
]
```

#### Location

If a posting specifies a location, it must match one of the candidate's preferred locations.

#### Experience

The system filters out jobs requiring more experience than the candidate's configured limit.

It handles common patterns such as:

```text
5+ years
3-5 years
4 years of experience
```

#### Excluded Keywords

Jobs containing configured exclusion terms are removed before scoring.

These checks can be applied across the available:

* Job title
* Description
* Location

---

### 2. Relevance Scoring

Only jobs that pass the hard filters are scored.

The relevance score is based on the number of configured skills/keywords appearing in the job posting.

For example:

```text
Candidate keywords:

Python
FastAPI
REST APIs
AWS
SQL
Docker
GitHub Actions
```

A posting matching:

```text
Python
REST APIs
SQL
AWS
```

receives a higher relevance score than one matching only:

```text
Python
```

The score is used to **rank matching jobs**, rather than replacing the hard eligibility filters.

---

## 👥 Multi-Person Configuration

The repository supports multiple independent candidate profiles.

```text
configs/
├── _template.json
├── vignesh.json
├── friend_1.json
└── friend_2.json
```

Each configuration defines that candidate's:

* Target roles
* Skills
* Locations
* Experience limit
* Excluded keywords
* Email destination

Example:

```json
{
  "job_roles": [
    "Software Engineer",
    "Backend Engineer"
  ],
  "keywords": [
    "Python",
    "FastAPI",
    "SQL",
    "AWS"
  ],
  "preferred_locations": [
    "Bengaluru",
    "Hyderabad"
  ],
  "max_experience_years": 2,
  "exclude_keywords": [
    "Senior",
    "Staff",
    "Principal"
  ],
  "email": {
    "recipient_email": "candidate@example.com"
  }
}
```

The workflow automatically discovers the configuration files and processes each candidate independently.

---

## 🗂️ Project Structure

```text
jobs-notification/
│
├── .github/
│   └── workflows/
│       └── job-monitor.yml
│
├── configs/
│   ├── _template.json
│   └── <candidate>.json
│
├── data/
│   └── <candidate>_jobs.xlsx
│
├── job_monitor.py
├── requirements.txt
└── README.md
```

### Components

| Component            | Purpose                                                                   |
| -------------------- | ------------------------------------------------------------------------- |
| `job_monitor.py`     | Core collection, filtering, scoring, deduplication and notification logic |
| `configs/`           | Candidate-specific matching configuration                                 |
| `data/`              | Per-candidate job history used for deduplication                          |
| `.github/workflows/` | Scheduled GitHub Actions automation                                       |
| `requirements.txt`   | Python dependencies                                                       |

---

## ⚙️ Automation with GitHub Actions

The system is designed to run without a continuously running server.

GitHub Actions:

1. Starts the scheduled workflow.
2. Loads the repository secrets.
3. Discovers candidate configuration files.
4. Runs the monitoring pipeline for each candidate.
5. Collects and filters job postings.
6. Removes previously processed postings.
7. Sends email notifications for new matches.
8. Updates the job history.

This turns the project into a small **serverless scheduled automation pipeline**.

---

## 🔐 Configuration & Secrets

Sensitive credentials are not stored directly in the source code.

The workflow uses GitHub Actions Secrets for:

```text
GMAIL_ADDRESS
GMAIL_APP_PASSWORD
ADZUNA_APP_ID
ADZUNA_APP_KEY
```

Candidate-specific configuration contains only matching preferences and the destination email address.

> **Security note:** Never commit API keys, Gmail app passwords, access tokens, or other credentials to the repository.

---

## 🛠️ Technology Stack

### Backend / Automation

* Python
* REST APIs
* JSON
* Excel-based job history

### Automation / CI/CD

* GitHub Actions
* Scheduled workflows
* Environment variables
* GitHub Actions Secrets

### Data Sources

* Adzuna API
* Arbeitnow API
* Company career pages

### Email

* Gmail SMTP
* App Password authentication

---

## 💻 Running Locally

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure environment variables

Linux/macOS:

```bash
export GMAIL_ADDRESS="you@gmail.com"
export GMAIL_APP_PASSWORD="xxxxxxxxxxxxxxxx"
export ADZUNA_APP_ID="your_app_id"
export ADZUNA_APP_KEY="your_app_key"
```

Windows PowerShell:

```powershell
$env:GMAIL_ADDRESS="you@gmail.com"
$env:GMAIL_APP_PASSWORD="xxxxxxxxxxxxxxxx"
$env:ADZUNA_APP_ID="your_app_id"
$env:ADZUNA_APP_KEY="your_app_key"
```

### 3. Run the monitor

```bash
python job_monitor.py configs/vignesh.json
```

The same program can be executed for any candidate configuration.

---

## 📈 Why This Project?

This project was built to solve a practical automation problem while applying several software engineering concepts:

* API integration
* Data normalization
* Rule-based filtering
* Text matching
* Relevance scoring
* Deduplication
* Configuration-driven architecture
* Scheduled automation
* CI/CD workflows
* Secret management
* Automated email delivery

Rather than building a one-off job scraper, the system was designed around **reusability and configuration**, allowing the same pipeline to serve different candidate profiles without changing the core application.

---

## 🔮 Possible Future Improvements

Potential extensions include:

* Browser automation for JavaScript-heavy career portals
* More job-source integrations
* Persistent database storage instead of spreadsheet-based history
* Semantic/embedding-based job matching
* Resume-aware relevance scoring
* Web dashboard for candidate configurations and job history
* Retry and failure monitoring
* Job-status tracking
* Improved notification summaries

---

## ⚠️ Limitations

This project is intended as a **first-pass job discovery and filtering system**, not a replacement for checking company career pages directly.

Some company career sites use JavaScript-heavy applications, bot protection, or other mechanisms that make simple HTTP-based collection unreliable.

Aggregator APIs such as Adzuna and Arbeitnow generally provide more structured job descriptions and therefore work better with the filtering and scoring pipeline.

The system may therefore miss jobs from certain career portals.

---

## 👤 Author

**Vignesh G**

M.Tech — Computer Science & Information Security
Software / Platform Engineering | Agentic AI | Automation

GitHub: [@vignesh-2706](https://github.com/vignesh-2706)

---

## 📄 License

This project is intended primarily as a personal engineering project and demonstration of software automation, API integration, and CI/CD practices.

````

### A few changes I'd make before you publish

There are **three things I'd check carefully** in the repository before recruiters start looking at it:

**1. Make absolutely sure `data/` contains no personal information.**

Your README says `data/*_jobs.xlsx` is generated per person, and the repository currently exposes a `data/` directory.

If those Excel files contain things like:

- personal email addresses
- job-search history
- friends' names
- private job preferences
- recruiter/company information you don't want public

don't commit them.

I'd preferably make the generated data directory contain only:

```text
data/
└── .gitkeep
````

and add:

```gitignore
data/*
!data/.gitkeep
```

Then the GitHub Action generates the actual files during execution.

**2. Be careful with `configs/*.json`.**

Because your design puts `email.recipient_email` in each person's config, a public repository could expose everyone's email address. Your README currently explicitly describes this architecture.

For a **public recruiter-facing repo**, I'd strongly consider keeping only:

```text
configs/
└── _template.json
```

public, and having your private deployment repository contain the actual candidate configurations.

That gives you a much cleaner public architecture:

```text
PUBLIC REPO
├── job_monitor.py
├── _template.json
├── workflow
├── requirements.txt
└── README.md

PRIVATE CONFIG
├── vignesh.json
├── friend1.json
└── friend2.json
```

**3. Don't advertise “friends” in the first sentence.**

The engineering feature is actually **configuration-driven multi-profile automation**. That's much more impressive professionally than:

> “you and any number of friends”

A recruiter sees:

> **Configuration-driven, multi-profile job monitoring pipeline**

and immediately understands the engineering concept.

---

### One more resume-related point

This project is actually worth keeping on your resume, but I would describe it as:

> **Built a configuration-driven job monitoring pipeline using Python and GitHub Actions, integrating external job APIs, rule-based filtering, relevance scoring, deduplication, and automated email notifications.**

That gives you legitimate keywords for **Python + APIs + automation + CI/CD + GitHub Actions + software engineering**, without pretending that you have deep Docker/Kubernetes/AWS experience.

Your public repository can then demonstrate those claims directly. [Job Monitor — GitHub repository](https://github.com/vignesh-2706/jobs-notification?utm_source=chatgpt.com)
