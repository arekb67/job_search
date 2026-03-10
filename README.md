# Job Search Pipeline

An automated two-phase pipeline that scrapes job listings from company career pages and scores them against your CV using the Claude API.

Built for job seekers targeting specific roles in specific industries — currently configured for trading infrastructure and platform roles in London, but fully customisable via `config.yaml`.

## How It Works

### Phase 1 — Web Scraper (`job_scraper_2stage.py`)

1. **Stage 1 (Quick Scan):** Visits career pages for all target companies and extracts job titles. Each title is scored against your target roles using keyword matching.
2. **Stage 2 (Deep Validation):** Fetches full job descriptions for the top candidates and validates them against location (e.g. London/Remote only) and domain (e.g. trading infrastructure only). Rejects jobs that don't match.

**Output files** (saved to `search_results/` with date suffix):
- `pass1_quick_results-YYYYMMDD.xlsx` — all jobs found, with title-based scores
- `pass2_top20_detailed-YYYYMMDD.xlsx` — top validated jobs with full descriptions
- `failed_companies-YYYYMMDD.xlsx` — companies whose career pages couldn't be scraped
- `failed_web_searches-YYYYMMDD.xlsx` — consolidated error log
- `rejected_jobs-YYYYMMDD.xlsx` — jobs rejected by location/domain validation

### Phase 2 — CV Scorer (`phase2_scorer.py`)

Takes the Phase 1 output and sends each job description to the Claude API along with your CV. Claude scores each job across five dimensions:

| Dimension | What it measures |
|---|---|
| Technical Match | How well the required skills align with your background |
| Role Match | How closely the role type fits your target positions |
| Seniority Match | Whether the seniority level is appropriate for your experience |
| Industry Match | How relevant the company/industry is to your targets |
| Location Match | Whether the job location meets your preferences |

Each job gets an overall score and a recommendation: **Apply Immediately**, **Strong Consider**, **Consider**, or **Skip**.

**Output:** `search_results/phase2_scored_results-YYYYMMDD.xlsx`

## Prerequisites

- **Python 3.10+**
- **Anthropic API key** (for Phase 2 CV scoring) — get one at [console.anthropic.com](https://console.anthropic.com/)

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/arekb67/job_search.git
cd job_search
```

### 2. Create a virtual environment

```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Create a `.env` file

Create a file called `.env` in the project root with your Anthropic API key:

```
ANTHROPIC_API_KEY=your_api_key_here
```

### 5. Add your CV

Place your CV as a `.docx` file in the project root. The filename should match the pattern configured in `config.yaml` (default: `CV - Arek Baranowski*.docx`). Update this pattern to match your own CV filename.

### 6. Customise `config.yaml`

Edit `config.yaml` to configure:

- **Target job titles** and their relevance scores
- **Target companies** and their career page URLs
- **Excluded job titles** (e.g. software engineer, data scientist)
- **Preferred locations** and location verification rules
- **Domain verification keywords** (to filter out irrelevant roles)
- **Scoring thresholds** for recommendations
- **Candidate profile** (your target roles and technical skills)
- **Claude model** and API settings

The file is heavily commented — each section explains what it does and how to edit it.

## Usage

### Run the full pipeline (Phase 1 + Phase 2)

```bash
./run_job_search_now.sh
```

### Run phases individually

```bash
# Phase 1 only (web scraping)
./run_job_search_now.sh --phase1

# Phase 2 only (Claude API scoring) — requires Phase 1 output
./run_job_search_now.sh --phase2
```

### Run Python scripts directly

```bash
# Phase 1
python3 job_scraper_2stage.py

# Phase 2 (auto-detects today's Phase 1 output)
python3 phase2_scorer.py

# Phase 2 with a specific date
python3 phase2_scorer.py --date 20260303

# Phase 2 with a specific input file
python3 phase2_scorer.py --input search_results/pass2_top20_detailed-20260303.xlsx
```

### Schedule daily runs (macOS only)

The project includes macOS `launchd` scripts to run the pipeline automatically on weekdays at 10:00 AM:

```bash
# Start the daily schedule
./start_scheduler.sh

# Stop the daily schedule
./stop_scheduler.sh
```

## Project Structure

```
job_search/
├── job_scraper_2stage.py      # Phase 1: web scraper
├── phase2_scorer.py           # Phase 2: Claude API CV scorer
├── config_loader.py           # Shared config loader
├── config.yaml                # All user-configurable settings
├── requirements.txt           # Python dependencies
├── run_job_search_now.sh      # Run pipeline on demand
├── run_daily_job_search.sh    # Pipeline runner (used by scheduler)
├── start_scheduler.sh         # Start macOS daily schedule
├── stop_scheduler.sh          # Stop macOS daily schedule
├── backup_job_search.sh       # Create versioned backups
├── Job_Scraper_User_Guide.docx # Detailed user guide
├── .env                       # API keys (not in repo)
├── search_results/            # Output files (not in repo)
└── logs/                      # Run logs (not in repo)
```

## Costs

Phase 2 uses the Claude API, which is a paid service. Each run scores ~20 jobs, using roughly 100k–200k tokens depending on job description length. Check [Anthropic's pricing](https://www.anthropic.com/pricing) for current rates.

## Customising for Your Own Job Search

To adapt this pipeline for a different role or industry:

1. **`config.yaml` > `target_job_titles`** — replace with your target job titles and scores
2. **`config.yaml` > `target_companies`** — replace with companies you want to monitor
3. **`config.yaml` > `exclude_keywords`** — update the list of roles to skip
4. **`config.yaml` > `candidate_profile`** — update with your target roles and skills
5. **`config.yaml` > `domain_verification`** — update the required/excluded domain keywords
6. **`config.yaml` > `preferred_locations`** and **`location_verification`** — set your location preferences
7. **`.env`** — add your own Anthropic API key
8. Place your own CV (`.docx` format) in the project root and update the `cv_filename_pattern` in `config.yaml`

## Documentation

For a comprehensive guide covering installation, daily workflow, output file details, troubleshooting, and the scoring algorithm, see `Job_Scraper_User_Guide.docx`.

## Licence

This project is provided as-is for personal use.
