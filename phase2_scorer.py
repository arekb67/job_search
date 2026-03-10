#!/usr/bin/env python3
"""
PHASE 2: LLM-POWERED CV-MATCHING JOB SCORER
============================================

Uses the Anthropic Claude API (Sonnet) to analyse each job description against
Arek Baranowski's CV across five dimensions:
  - Technical Match
  - Role Match
  - Seniority Match
  - Industry Match
  - Location Match

Produces: search_results/phase2_scored_results-DATE.xlsx

REQUIREMENTS:
    pip install anthropic pandas openpyxl python-docx pyyaml

SETUP:
    1. Ensure config.yaml exists in the same folder as this script
    2. Add your API key to the .env file:
        ANTHROPIC_API_KEY=your_api_key_here

USAGE:
    python3 phase2_scorer.py                     # auto-detects today's pass2 file
    python3 phase2_scorer.py --date 20260223     # specify a date explicitly
    python3 phase2_scorer.py --input pass2.xlsx  # specify input file directly
"""

import pandas as pd
import json
import os
import sys
import glob
import time
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Load all settings from config.yaml
from config_loader import load_config
config = load_config()

# ============================================================================
# CONFIGURATION (loaded from config.yaml — edit that file, not this one)
# ============================================================================

# File paths
WORKING_DIR = os.path.dirname(os.path.abspath(__file__))
RESULTS_DIR = os.path.join(WORKING_DIR, "search_results")
CV_FILENAME_PATTERN = config['scorer_settings']['cv_filename_pattern']


# Candidate profile
YEARS_OF_EXPERIENCE = config['scorer_settings']['years_of_experience']
TARGET_ROLES = config['candidate_profile']['target_roles']
TECHNICAL_SKILLS = config['candidate_profile']['technical_skills']

# Industry classification
INDUSTRY_TIERS = config['industry_tiers']

# Location scoring
LOCATION_SCORES = config['location_scores']

# Recommendation thresholds and colours
RECOMMENDATION_THRESHOLDS = config['recommendations']['thresholds']
RECOMMENDATION_COLOURS = config['recommendations']['colours']

# API settings
CLAUDE_MODEL = config['scorer_settings']['claude_model']
MAX_TOKENS = config['scorer_settings']['max_tokens']
API_RETRY_ATTEMPTS = config['scorer_settings']['api_retry_attempts']
API_RETRY_DELAY = config['scorer_settings']['api_retry_delay']
API_RATE_LIMIT_PAUSE = config['scorer_settings']['api_rate_limit_pause']
JOB_DESCRIPTION_MAX_CHARS = config['scorer_settings']['job_description_max_chars']

# ============================================================================
# END OF CONFIGURATION — you should not need to edit below this line
# ============================================================================

# Load environment variables (for ANTHROPIC_API_KEY)
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(WORKING_DIR, ".env"))
except ImportError:
    pass


# ============================================================================
# CV FILE DETECTION
# ============================================================================

def find_latest_cv():
    """
    Find the most recent CV file matching CV_FILENAME_PATTERN.
    Picks the file with the highest date in the filename (YYYYMMDD).
    Falls back to most recently modified file if no date found.
    """
    import re as _re
    pattern = os.path.join(WORKING_DIR, CV_FILENAME_PATTERN)
    matches = glob.glob(pattern)

    if not matches:
        return None

    # Try to extract YYYYMMDD dates from filenames and pick the highest
    dated_files = []
    for filepath in matches:
        filename = os.path.basename(filepath)
        date_match = _re.search(r'(\d{8})', filename)
        if date_match:
            dated_files.append((date_match.group(1), filepath))

    if dated_files:
        # Sort by date string descending, return the most recent
        dated_files.sort(key=lambda x: x[0], reverse=True)
        return dated_files[0][1]

    # No dates in filenames — fall back to most recently modified
    return max(matches, key=os.path.getmtime)


CV_PATH = find_latest_cv()
if CV_PATH is None:
    print(f"ERROR: No CV file found matching '{CV_FILENAME_PATTERN}'")
    print(f"  Looked in: {WORKING_DIR}")
    print(f"  Expected format: CV - Arek Baranowski YYYYMMDD.docx")
    sys.exit(1)


# ============================================================================
# CV TEXT EXTRACTION
# ============================================================================

def extract_cv_text():
    """Extract text from the CV .docx file."""
    try:
        from docx import Document as DocxDocument
        doc = DocxDocument(CV_PATH)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except ImportError:
        import subprocess
        result = subprocess.run(
            ["pandoc", CV_PATH, "-t", "plain"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            return result.stdout
        print("ERROR: Could not extract CV text. Install python-docx: pip install python-docx")
        sys.exit(1)


# ============================================================================
# FILE DISCOVERY
# ============================================================================

def find_pass2_file(date_str=None):
    """Find the pass2_top20_detailed file for the given date (or today)."""
    if date_str is None:
        date_str = datetime.now().strftime('%Y%m%d')

    dated_file = os.path.join(RESULTS_DIR, f"pass2_top20_detailed-{date_str}.xlsx")
    if os.path.exists(dated_file):
        return dated_file

    undated_file = os.path.join(RESULTS_DIR, "pass2_top20_detailed.xlsx")
    if os.path.exists(undated_file):
        return undated_file

    pattern = os.path.join(RESULTS_DIR, "pass2_top20_detailed-*.xlsx")
    files = glob.glob(pattern)
    if files:
        files.sort(reverse=True)
        return files[0]

    return None


# ============================================================================
# CLAUDE API — JOB SCORING
# ============================================================================

def build_scoring_prompt(cv_text, job_title, job_company, job_location, job_description):
    """Build the prompt for Claude to score a single job against the CV."""

    # Build dynamic sections from configuration variables
    skills_text = ", ".join(TECHNICAL_SKILLS)
    roles_text = ", ".join(TARGET_ROLES)

    industry_lines = []
    for tier_key in sorted(INDUSTRY_TIERS.keys()):
        tier = INDUSTRY_TIERS[tier_key]
        industry_lines.append(f"{tier_key.title()} (score {tier['score_range']}): {tier['description']}")
    industry_text = ". ".join(industry_lines)

    location_lines = []
    for loc_type, score_range in LOCATION_SCORES.items():
        location_lines.append(f"{loc_type} = {score_range}")
    location_text = ". ".join(location_lines)

    return f"""You are an expert career adviser analysing job-to-CV fit. You must score the following job vacancy against the candidate's CV.

## CANDIDATE CV

{cv_text}

## JOB VACANCY

**Company:** {job_company}
**Title:** {job_title}
**Location:** {job_location}
**Description:**
{job_description[:JOB_DESCRIPTION_MAX_CHARS]}

## SCORING INSTRUCTIONS

Score this job across exactly five dimensions (0-100 each). Be rigorous and specific — reference concrete CV details and job requirements in your analysis.

1. **Technical Match (0-100):** How well do the required technical skills align with the candidate's background? Consider: {skills_text}.

2. **Role Match (0-100):** How closely does the role type fit the candidate's target positions? Target roles: {roles_text}. Penalise pure IC/software engineering roles heavily.

3. **Seniority Match (0-100):** Is the seniority level appropriate for {YEARS_OF_EXPERIENCE}+ years of experience with leadership background (Head of, Director, VP, Senior)? Penalise junior/mid-level roles. Consider years-of-experience requirements stated in the job description.

4. **Industry Match (0-100):** How relevant is the company/role to target industries? {industry_text}. Below 50: non-financial or irrelevant sector.

5. **Location Match (0-100):** {location_text}. Check the ACTUAL job location from the description, not just the metadata — some jobs list London in metadata but are actually based elsewhere (look for USD salaries, US cities in the description).

## RESPONSE FORMAT

Respond with ONLY a valid JSON object (no markdown fencing, no explanation outside the JSON). Use this exact structure:

{{
  "technical_match": <int 0-100>,
  "role_match": <int 0-100>,
  "seniority_match": <int 0-100>,
  "industry_match": <int 0-100>,
  "location_match": <int 0-100>,
  "key_matches": "<bullet-pointed list of specific strengths, referencing concrete CV experience and job requirements>",
  "gaps_notes": "<bullet-pointed list of gaps, concerns, or warnings — be specific about what's missing or misaligned>"
}}

Use bullet points starting with "• " in the key_matches and gaps_notes fields. Be specific — reference actual role names, companies, technologies, and achievements from the CV."""


def score_job_with_claude(client, cv_text, title, company, location, description):
    """Call the Claude API to score a single job. Returns parsed JSON result or None."""
    prompt = build_scoring_prompt(cv_text, title, company, location, description)

    for attempt in range(API_RETRY_ATTEMPTS):
        try:
            response = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}],
            )

            response_text = response.content[0].text.strip()

            # Strip markdown code fencing if present
            if response_text.startswith("```"):
                lines = response_text.split("\n")
                lines = [l for l in lines if not l.strip().startswith("```")]
                response_text = "\n".join(lines)

            result = json.loads(response_text)

            # Validate required fields
            required_fields = [
                "technical_match", "role_match", "seniority_match",
                "industry_match", "location_match", "key_matches", "gaps_notes"
            ]
            for field in required_fields:
                if field not in result:
                    raise ValueError(f"Missing field: {field}")

            # Clamp scores to 0-100
            for score_field in ["technical_match", "role_match", "seniority_match",
                                "industry_match", "location_match"]:
                result[score_field] = max(0, min(100, int(result[score_field])))

            return result

        except json.JSONDecodeError as e:
            print(f"\n    WARNING: JSON parse error (attempt {attempt + 1}/{API_RETRY_ATTEMPTS}): {e}")
            if attempt < API_RETRY_ATTEMPTS - 1:
                time.sleep(API_RETRY_DELAY)
        except Exception as e:
            print(f"\n    WARNING: API error (attempt {attempt + 1}/{API_RETRY_ATTEMPTS}): {e}")
            if attempt < API_RETRY_ATTEMPTS - 1:
                time.sleep(API_RETRY_DELAY)

    return None  # All retries exhausted


def get_recommendation(overall_score):
    """Map overall score to recommendation using configured thresholds."""
    if overall_score >= RECOMMENDATION_THRESHOLDS["apply_immediately"]:
        return "Apply Immediately"
    elif overall_score >= RECOMMENDATION_THRESHOLDS["strong_consider"]:
        return "Strong Consider"
    elif overall_score >= RECOMMENDATION_THRESHOLDS["consider"]:
        return "Consider"
    else:
        return "Skip"


# ============================================================================
# EXCEL OUTPUT
# ============================================================================

def write_results_to_excel(results, output_file):
    """Write scored results to a formatted Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Phase 2 Scored Results"

    header_font = Font(name="Arial", bold=True, size=10)
    header_fill = PatternFill("solid", fgColor="D5E8F0")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    body_font = Font(name="Arial", size=10)
    wrap_alignment = Alignment(vertical="top", wrap_text=True)
    centre_alignment = Alignment(horizontal="center", vertical="top")

    headers = [
        "Rank", "Company", "Title", "Location", "URL",
        "Technical\nMatch", "Role\nMatch", "Seniority\nMatch",
        "Industry\nMatch", "Location\nMatch", "Overall\nScore",
        "Recommendation", "Key Matches", "Gaps / Notes"
    ]

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    col_widths = {
        1: 6, 2: 22, 3: 40, 4: 25, 5: 45,
        6: 12, 7: 10, 8: 12, 9: 12, 10: 12,
        11: 11, 12: 18, 13: 60, 14: 60,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    for row_idx, result in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=result["rank"]).font = body_font
        ws.cell(row=row_idx, column=1).alignment = centre_alignment
        ws.cell(row=row_idx, column=1).border = thin_border

        ws.cell(row=row_idx, column=2, value=result["company"]).font = body_font
        ws.cell(row=row_idx, column=2).alignment = wrap_alignment
        ws.cell(row=row_idx, column=2).border = thin_border

        ws.cell(row=row_idx, column=3, value=result["title"]).font = body_font
        ws.cell(row=row_idx, column=3).alignment = wrap_alignment
        ws.cell(row=row_idx, column=3).border = thin_border

        ws.cell(row=row_idx, column=4, value=result["location"]).font = body_font
        ws.cell(row=row_idx, column=4).alignment = wrap_alignment
        ws.cell(row=row_idx, column=4).border = thin_border

        ws.cell(row=row_idx, column=5, value=result["url"]).font = body_font
        ws.cell(row=row_idx, column=5).alignment = wrap_alignment
        ws.cell(row=row_idx, column=5).border = thin_border

        for col, key in [(6, "technical_match"), (7, "role_match"), (8, "seniority_match"),
                         (9, "industry_match"), (10, "location_match")]:
            cell = ws.cell(row=row_idx, column=col, value=result[key])
            cell.font = body_font
            cell.alignment = centre_alignment
            cell.border = thin_border

        # Overall Score as AVERAGE formula
        ws.cell(row=row_idx, column=11,
                value=f"=AVERAGE(F{row_idx},G{row_idx},H{row_idx},I{row_idx},J{row_idx})")
        ws.cell(row=row_idx, column=11).font = Font(name="Arial", bold=True, size=10)
        ws.cell(row=row_idx, column=11).alignment = centre_alignment
        ws.cell(row=row_idx, column=11).border = thin_border

        rec_cell = ws.cell(row=row_idx, column=12, value=result["recommendation"])
        rec_cell.font = body_font
        rec_cell.alignment = centre_alignment
        rec_cell.border = thin_border
        rec_cell.fill = PatternFill("solid", fgColor=RECOMMENDATION_COLOURS.get(result["recommendation"], "FFFFFF"))

        ws.cell(row=row_idx, column=13, value=result["key_matches"]).font = body_font
        ws.cell(row=row_idx, column=13).alignment = wrap_alignment
        ws.cell(row=row_idx, column=13).border = thin_border

        ws.cell(row=row_idx, column=14, value=result["gaps_notes"]).font = body_font
        ws.cell(row=row_idx, column=14).alignment = wrap_alignment
        ws.cell(row=row_idx, column=14).border = thin_border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:N{len(results) + 1}"
    wb.save(output_file)


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    print("=" * 80)
    print("PHASE 2: LLM-POWERED CV-MATCHING JOB SCORER")
    print(f"Model: {CLAUDE_MODEL}")
    print("=" * 80)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    date_suffix = datetime.now().strftime('%Y%m%d')

    # Parse command line arguments
    input_file = None
    for i, arg in enumerate(sys.argv[1:], 1):
        if arg == "--date" and i < len(sys.argv) - 1:
            date_suffix = sys.argv[i + 1]
        elif arg == "--input" and i < len(sys.argv) - 1:
            input_file = sys.argv[i + 1]

    # Initialise Claude API client
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("ERROR: ANTHROPIC_API_KEY not set.")
        print("  Add ANTHROPIC_API_KEY=your_key to your .env file")
        print(f"  (.env location: {os.path.join(WORKING_DIR, '.env')})")
        print("  Or set the ANTHROPIC_API_KEY environment variable.")
        sys.exit(1)

    try:
        import anthropic
    except ImportError:
        print("ERROR: anthropic package not installed.")
        print("  Run: pip install anthropic")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    print(f"Claude API client initialised (model: {CLAUDE_MODEL})")

    # Find input file
    if input_file is None:
        input_file = find_pass2_file(date_suffix)

    if input_file is None:
        print("ERROR: Could not find pass2_top20_detailed file.")
        print(f"  Looked for: pass2_top20_detailed-{date_suffix}.xlsx")
        print(f"  In directory: {WORKING_DIR}")
        sys.exit(1)

    print(f"Input file: {input_file}")
    print(f"CV file: {CV_PATH}")
    print(f"Date suffix: {date_suffix}")

    # Read the Phase 1 output
    try:
        df = pd.read_excel(input_file, sheet_name='Top 20 Detailed')
    except Exception:
        df = pd.read_excel(input_file)

    if df.empty:
        print("WARNING: No jobs found in input file. Creating empty output.")

    print(f"\nJobs to score: {len(df)}")

    # Extract CV text
    cv_text = extract_cv_text()
    print(f"CV text extracted: {len(cv_text)} characters")

    # ========================================================================
    # SCORE EACH JOB VIA CLAUDE API
    # ========================================================================

    results = []
    failed_jobs = []
    print(f"\nScoring {len(df)} jobs via Claude API...\n")

    for idx, row in df.iterrows():
        company = str(row.get('company', ''))
        title = str(row.get('title', ''))
        location = str(row.get('location', ''))
        url = str(row.get('url', ''))
        description = str(row.get('full_description', ''))

        print(f"  [{idx + 1}/{len(df)}] {company} — {title[:55]}...", end=" ")

        result = score_job_with_claude(client, cv_text, title, company, location, description)

        if result is None:
            print("FAILED (all retries exhausted)")
            failed_jobs.append({"company": company, "title": title})
            continue

        overall = (
            result["technical_match"] +
            result["role_match"] +
            result["seniority_match"] +
            result["industry_match"] +
            result["location_match"]
        ) / 5.0

        recommendation = get_recommendation(overall)

        results.append({
            "company": company,
            "title": title,
            "location": location,
            "url": url,
            "technical_match": result["technical_match"],
            "role_match": result["role_match"],
            "seniority_match": result["seniority_match"],
            "industry_match": result["industry_match"],
            "location_match": result["location_match"],
            "overall_score": overall,
            "recommendation": recommendation,
            "key_matches": result["key_matches"],
            "gaps_notes": result["gaps_notes"],
        })

        print(f"Overall: {overall:.0f} — {recommendation}")

        # Brief pause between API calls to avoid rate limiting
        time.sleep(API_RATE_LIMIT_PAUSE)

    # Sort by overall score descending
    results.sort(key=lambda x: x["overall_score"], reverse=True)

    # Assign ranks
    for i, r in enumerate(results, 1):
        r["rank"] = i

    # ========================================================================
    # WRITE OUTPUT
    # ========================================================================

    os.makedirs(RESULTS_DIR, exist_ok=True)
    output_file = os.path.join(RESULTS_DIR, f"phase2_scored_results-{date_suffix}.xlsx")
    print(f"\nWriting results to: {output_file}")

    write_results_to_excel(results, output_file)
    print(f"Output saved: {output_file}")

    # ========================================================================
    # SUMMARY
    # ========================================================================

    print("\n" + "=" * 80)
    print("PHASE 2 SUMMARY")
    print("=" * 80)
    print(f"\nJobs scored: {len(results)}")
    if failed_jobs:
        print(f"Jobs failed (API errors): {len(failed_jobs)}")
        for fj in failed_jobs:
            print(f"  - {fj['company']}: {fj['title'][:60]}")

    rec_counts = {}
    for r in results:
        rec = r["recommendation"]
        rec_counts[rec] = rec_counts.get(rec, 0) + 1

    for rec in ["Apply Immediately", "Strong Consider", "Consider", "Skip"]:
        count = rec_counts.get(rec, 0)
        if count > 0:
            print(f"  {rec}: {count}")

    print(f"\nTOP 5 CANDIDATES:")
    for r in results[:5]:
        print(f"\n  #{r['rank']} [{r['overall_score']:.0f}] {r['company']}")
        print(f"      {r['title'][:70]}")
        print(f"      {r['recommendation']}")
        print(f"      Tech:{r['technical_match']} Role:{r['role_match']} "
              f"Senior:{r['seniority_match']} Industry:{r['industry_match']} "
              f"Loc:{r['location_match']}")

    print(f"\nOutput: {output_file}")
    print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()
