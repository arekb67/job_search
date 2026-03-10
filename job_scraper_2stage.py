#!/usr/bin/env python3
"""
COMPREHENSIVE 2-STAGE JOB SCRAPER - FULLY AUTONOMOUS
=====================================================

STAGE 1: Quick scan of target companies (titles only)
STAGE 2: Auto-fetch full job descriptions for top candidates

Then YOU paste top candidates to Claude for strategic analysis.

REQUIREMENTS:
    pip install requests beautifulsoup4 pandas openpyxl

SETUP:
    Configuration is loaded from config.yaml using config_loader module.

USAGE:
    python3 job_scraper_2stage.py

OUTPUT (saved to search_results/ subfolder, filenames include date suffix -yyyymmdd):
    - search_results/pass1_quick_results-DATE.xlsx    (all jobs, title-based scoring)
    - search_results/pass2_top20_detailed-DATE.xlsx   (top 20 with full descriptions + Failed Companies sheet)
    - search_results/failed_web_searches-DATE.xlsx    (consolidated failure log with error codes & diagnostics)
    - search_results/failed_companies-DATE.xlsx       (company career pages that could not be scraped)
    - search_results/rejected_jobs-DATE.xlsx          (jobs rejected by location/domain validation in Stage 2)
"""

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import time
import os
import sys
from datetime import datetime
from urllib.parse import urljoin, urlparse
import glob

# Load all settings from config.yaml
from config_loader import load_config
config = load_config()

# ============================================================================
# CONFIGURATION (loaded from config.yaml — edit that file, not this one)
# ============================================================================

# Scraper behaviour
MIN_SCORE_THRESHOLD = config['scraper_settings']['min_score_threshold']
DESCRIPTION_MAX_CHARS = config['scraper_settings']['description_max_chars']
REQUEST_TIMEOUT = config['scraper_settings']['request_timeout']
POLITE_DELAY = config['scraper_settings']['polite_delay']
MAX_RETRIES = config['scraper_settings']['max_retries']
RETRY_BACKOFF = config['scraper_settings']['retry_backoff']
STAGE2_TARGET_VALID = config['scraper_settings']['stage2_target_valid']
STAGE2_CANDIDATE_POOL = config['scraper_settings']['stage2_candidate_pool']
USER_AGENT = config['scraper_settings']['user_agent']

# Target job titles and their scores (dict: title -> {score, category})
TARGET_KEYWORDS = config['target_job_titles']

# Which job title categories to search at each company type
COMPANY_TYPE_CATEGORIES = config['company_type_categories']

# Job titles to automatically exclude
EXCLUDE_KEYWORDS = config['exclude_keywords']

# Preferred locations for initial filtering
PREFERRED_LOCATIONS = config['preferred_locations']

# Location verification keywords
LONDON_LOCATION_KEYWORDS = config['location_verification']['london_keywords']
REMOTE_LOCATION_KEYWORDS = config['location_verification']['remote_keywords']
NON_LONDON_CITIES = config['location_verification']['non_london_cities']

# Domain verification keywords
REQUIRED_DOMAIN_KEYWORDS = config['domain_verification']['required_keywords']
EXCLUDED_DOMAIN_KEYWORDS = config['domain_verification']['excluded_keywords']

# Salary detection patterns
USD_SALARY_PATTERN = config['salary_detection']['usd_pattern']
GBP_SALARY_PATTERN = config['salary_detection']['gbp_pattern']

# Target companies (loaded from config.yaml)
COMPANIES = config['target_companies']

# Output directory
output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'search_results')
os.makedirs(output_dir, exist_ok=True)

# ============================================================================
# HELPER: HTTP REQUEST WITH RETRY
# ============================================================================

def request_with_retry(method, url, **kwargs):
    """Make an HTTP request with automatic retry on transient failures.
    Returns (response, exception) - one will be None."""
    kwargs.setdefault('timeout', REQUEST_TIMEOUT)
    kwargs.setdefault('headers', {'User-Agent': USER_AGENT})

    last_exception = None
    for attempt in range(MAX_RETRIES):
        try:
            if method == 'GET':
                response = requests.get(url, **kwargs)
            else:
                response = requests.post(url, **kwargs)
            response.raise_for_status()
            return response, None
        except Exception as e:
            last_exception = e
            # Only retry on transient errors (timeout, 503, 429, connection errors)
            retryable = False
            if isinstance(e, (requests.exceptions.Timeout, requests.exceptions.ConnectionError)):
                retryable = True
            elif isinstance(e, requests.exceptions.HTTPError) and hasattr(e, 'response') and e.response is not None:
                if e.response.status_code in (429, 503, 502):
                    retryable = True

            if retryable and attempt < MAX_RETRIES - 1:
                wait = RETRY_BACKOFF * (attempt + 1)
                print(f"(retry in {wait}s)...", end=" ")
                time.sleep(wait)
            else:
                break

    return None, last_exception

# ============================================================================
# STAGE 1: QUICK SCORING
# ============================================================================

def quick_score(title, location="", company="", company_type=""):
    """Quick score based on title only (0-100).

    When company_type is provided, only title categories allowed for that
    company type are considered (e.g. hedge funds skip sales titles).
    """
    title_lower = title.lower().strip()
    location_lower = location.lower()

    # STRICT EXCLUSION: engineer or scientist
    if "engineer" in title_lower:
        if not ("engineering manager" in title_lower or "engineering director" in title_lower):
            return 0

    if "scientist" in title_lower:
        return 0

    # Exclude IC roles
    for exclude in EXCLUDE_KEYWORDS:
        if exclude in title_lower:
            return 0

    score = 0

    # Determine which categories are allowed for this company type
    allowed_categories = None
    if company_type and company_type in COMPANY_TYPE_CATEGORIES:
        allowed_categories = COMPANY_TYPE_CATEGORIES[company_type]

    # Check keywords
    for keyword, keyword_data in TARGET_KEYWORDS.items():
        if keyword in title_lower:
            kw_score = keyword_data['score']
            kw_category = keyword_data.get('category', 'infrastructure')

            # If we know the company type, skip categories not allowed
            if allowed_categories is not None and kw_category not in allowed_categories:
                continue

            score = max(score, kw_score)

    # Location bonus
    if any(loc in location_lower for loc in PREFERRED_LOCATIONS):
        score = min(100, score + 5)

    return score

# ============================================================================
# STAGE 1: SCRAPING FUNCTIONS
# ============================================================================

def classify_error(exception, response=None):
    """Classify a scraping error into a human-readable category with diagnostic details."""
    error_str = str(exception)
    error_type = type(exception).__name__

    # Build diagnostic info
    diag = {
        'error_type': error_type,
        'http_status': None,
        'status_reason': None,
        'server_header': None,
        'classification': 'Unknown',
        'detail': error_str[:200],
    }

    # Extract HTTP status from response if available
    if response is not None:
        diag['http_status'] = response.status_code
        diag['status_reason'] = response.reason
        diag['server_header'] = response.headers.get('Server', 'Not provided')

    # Classify the error
    if 'timeout' in error_str.lower() or 'timed out' in error_str.lower():
        diag['classification'] = 'Timeout'
    elif 'ConnectionError' in error_type or 'connection' in error_str.lower():
        diag['classification'] = 'Connection Failed'
    elif 'SSLError' in error_type or 'ssl' in error_str.lower() or 'certificate' in error_str.lower():
        diag['classification'] = 'SSL/TLS Error'
    elif 'TooManyRedirects' in error_type:
        diag['classification'] = 'Too Many Redirects'
    elif response is not None:
        code = response.status_code
        if code == 403:
            # Check for known anti-bot providers
            server = (response.headers.get('Server', '') + ' ' + response.text[:500]).lower()
            if 'cloudflare' in server:
                diag['classification'] = 'Blocked (Cloudflare)'
            elif 'akamai' in server:
                diag['classification'] = 'Blocked (Akamai)'
            elif 'captcha' in server or 'challenge' in server:
                diag['classification'] = 'Blocked (CAPTCHA/Challenge)'
            else:
                diag['classification'] = 'Blocked (403 Forbidden)'
        elif code == 401:
            diag['classification'] = 'Auth Required (401)'
        elif code == 404:
            diag['classification'] = 'Page Not Found (404)'
        elif code == 429:
            diag['classification'] = 'Rate Limited (429)'
        elif code == 503:
            diag['classification'] = 'Service Unavailable (503)'
        elif 400 <= code < 500:
            diag['classification'] = f'Client Error ({code})'
        elif 500 <= code < 600:
            diag['classification'] = f'Server Error ({code})'
    elif 'Max retries' in error_str:
        diag['classification'] = 'Max Retries Exceeded'
    elif 'Name or service not known' in error_str or 'getaddrinfo' in error_str:
        diag['classification'] = 'DNS Resolution Failed'

    return diag


def scrape_company_jobs(company_name, company_data):
    """Scrape job titles from company career page. Tries alt_urls if primary fails."""
    print(f"  {company_name}...", end=" ")

    # Build list of URLs to try: primary first, then any alternatives
    urls_to_try = [company_data['url']]
    if 'alt_urls' in company_data:
        urls_to_try.extend(company_data['alt_urls'])

    last_response = None
    last_exception = None

    for attempt_url in urls_to_try:
        jobs = []

        response, exc = request_with_retry('GET', attempt_url)

        if exc is not None:
            last_exception = exc
            last_response = None
            if attempt_url != urls_to_try[-1]:
                continue
            break

        soup = BeautifulSoup(response.text, 'html.parser')
        job_links = soup.find_all('a', href=True)

        seen_jobs = set()

        for link in job_links:
            title = link.get_text(strip=True)
            url = urljoin(attempt_url, link['href'])

            if len(title) < 10 or len(title) > 250:
                continue

            title_lower = title.lower()
            has_target = any(kw in title_lower for kw in TARGET_KEYWORDS.keys())

            if not has_target:
                continue

            score = quick_score(title, company_data.get('location', ''), company_name, company_data.get('type', ''))

            if score < MIN_SCORE_THRESHOLD:
                continue

            job_key = f"{title}_{url}"
            if job_key in seen_jobs:
                continue
            seen_jobs.add(job_key)

            jobs.append({
                'company': company_name,
                'title': title,
                'url': url,
                'location': company_data.get('location', ''),
                'type': company_data.get('type', ''),
                'source': 'Company Website'
            })

        if attempt_url != company_data['url']:
            print(f"OK {len(jobs)} jobs (via alt URL: {attempt_url[:50]})")
        else:
            print(f"OK {len(jobs)} jobs")
        time.sleep(POLITE_DELAY)
        return jobs, None  # Success, no error

    # All URLs failed - report the last error
    diag = classify_error(last_exception, last_response)
    print(f"FAIL {diag['classification']}")
    return [], {
        'company': company_name,
        'url': company_data['url'],
        'type': company_data.get('type', ''),
        'error_classification': diag['classification'],
        'http_status': diag['http_status'],
        'status_reason': diag['status_reason'],
        'server_header': diag['server_header'],
        'error_type': diag['error_type'],
        'error_detail': diag['detail'],
        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'source': 'Company Website',
    }


# ============================================================================
# STAGE 2: FETCH FULL JOB DESCRIPTIONS
# ============================================================================

def verify_job_location_and_domain(url):
    """
    Fetch job description and verify it's in London/Remote and related to trading/infrastructure.
    Returns: (description_text, is_valid, reason)
    """
    try:
        response, exc = request_with_retry('GET', url)
        if exc is not None:
            return f"Error fetching description: {str(exc)[:100]}", False, "Error fetching"

        soup = BeautifulSoup(response.text, 'html.parser')

        # Remove script and style elements
        for script in soup(["script", "style", "nav", "header", "footer"]):
            script.decompose()

        # Get full text
        full_text = soup.get_text()

        # Clean up whitespace
        lines = (line.strip() for line in full_text.splitlines())
        chunks = (phrase.strip() for line in lines for phrase in line.split("  "))
        text = '\n'.join(chunk for chunk in chunks if chunk)

        # Convert to lowercase for checking
        text_lower = text.lower()

        # ===================================================================
        # STRICT LOCATION VERIFICATION
        # ===================================================================
        #
        # IMPORTANT: We must detect the JOB-SPECIFIC location, not just any
        # mention of "London" anywhere on the page. Many global firms list
        # all their office cities in navigation/footer boilerplate, so a
        # Chicago job on Optiver's site will still mention "London" in the
        # site chrome. We solve this by:
        #   1. Looking for explicit location markers near the top of the page
        #      (e.g., "Location: Chicago", standalone city name near title)
        #   2. If a non-London city appears as the job's specific location,
        #      reject it -- even if "London" also appears elsewhere on the page
        # ===================================================================

        # --- Step 1: Extract the job-specific location from structured patterns ---
        top_section = text_lower[:2000]

        # Common patterns for job-specific location lines
        location_patterns = [
            r'location[:\s]+([a-z\s,]+)',        # "Location: Chicago"
            r'office[:\s]+([a-z\s,]+)',           # "Office: Chicago"
            r'(?:full[- ]time|part[- ]time|contract)\s*[/|–—-]\s*([a-z\s,]+)',  # "Full-time | Chicago"
            r'([a-z\s,]+)\s*[/|–—-]\s*(?:full[- ]time|part[- ]time|contract)',  # "Chicago | Full-time"
        ]

        detected_job_location = None

        # Check structured patterns in top section
        for pattern in location_patterns:
            match = re.search(pattern, top_section)
            if match:
                loc_text = match.group(1).strip()
                for city_name, city_variants in NON_LONDON_CITIES.items():
                    if any(variant in loc_text for variant in city_variants):
                        detected_job_location = city_name
                        break
            if detected_job_location:
                break

        # Also check for a standalone city name appearing prominently in the
        # top section (e.g., Optiver pages show just "Chicago" on its own line
        # right below the job title)
        if not detected_job_location:
            top_lines = [line.strip().lower() for line in text[:2000].splitlines() if line.strip()]
            for line in top_lines:
                # Short lines (under 40 chars) that ARE a city name = likely the location field
                if len(line) < 40:
                    for city_name, city_variants in NON_LONDON_CITIES.items():
                        if any(line == variant or line.startswith(variant + ",") or line.startswith(variant + " ")
                               for variant in city_variants):
                            detected_job_location = city_name
                            break
                if detected_job_location:
                    break

        # If we detected a specific non-London location for this job, reject it
        if detected_job_location:
            return text, False, f"Location: {detected_job_location.title()} (job-specific)"

        # --- Step 2: Standard location check (only if no specific non-London location found) ---
        has_london = any(keyword in text_lower for keyword in LONDON_LOCATION_KEYWORDS)
        has_remote = any(keyword in text_lower for keyword in REMOTE_LOCATION_KEYWORDS)

        if not (has_london or has_remote):
            # Check if any non-London city appears anywhere in the text
            for city_name, city_variants in NON_LONDON_CITIES.items():
                for variant in city_variants:
                    if variant in text_lower:
                        return text, False, f"Location: {city_name.title()}"

            # No clear location found - reject to be safe
            return text, False, "Location: Not London/Remote"

        # --- Step 3: Additional guard -- if London appears but ONLY in boilerplate ---
        # Some career pages (e.g. Optiver) mention "London" in site-wide footer
        # text but the actual job is in Austin/Amsterdam/etc.  The city may not
        # appear in our extracted text at all (stripped <header>, JS-rendered).
        # If London is NOT in the top 2000 chars (i.e. near the job content),
        # reject the job — do not assume it is London-based.
        if has_london and not has_remote:
            london_in_top = any(kw in top_section for kw in LONDON_LOCATION_KEYWORDS)
            if not london_in_top:
                # Check if a specific non-London city IS in the top section
                for city_name, city_variants in NON_LONDON_CITIES.items():
                    for variant in city_variants:
                        if variant in top_section:
                            return text, False, f"Location: {city_name.title()} (London only in boilerplate)"
                # No city found at all — London is only in boilerplate, reject
                return text, False, "Location: Uncertain (London only in site boilerplate, not in job content)"

        # ===================================================================
        # STRICT DOMAIN VERIFICATION (Trading Platform/Infrastructure)
        # ===================================================================

        has_relevant_domain = any(domain in text_lower for domain in REQUIRED_DOMAIN_KEYWORDS)

        if not has_relevant_domain:
            for excluded in EXCLUDED_DOMAIN_KEYWORDS:
                if excluded in text_lower:
                    return text, False, f"Domain: {excluded.title()}"

            return text, False, "Domain: Not trading platform/infrastructure"

        # ===================================================================
        # STEP 4: FULL-TEXT STANDALONE CITY SCAN
        # ===================================================================
        all_lines = [line.strip().lower() for line in text.splitlines() if line.strip()]
        for line in all_lines:
            if len(line) < 40:
                for city_name, city_variants in NON_LONDON_CITIES.items():
                    for variant in city_variants:
                        if line == variant or line.startswith(variant + ",") or line.startswith(variant + " "):
                            return text, False, f"Location: {city_name.title()} (standalone city line in full text)"

        # ===================================================================
        # STEP 5: USD SALARY DETECTION
        # ===================================================================
        if re.search(USD_SALARY_PATTERN, text):
            if not re.search(GBP_SALARY_PATTERN, text):
                return text, False, "Location: Likely US-based (USD salary, no GBP)"

        # ===================================================================
        # STEP 6: US-SPECIFIC BENEFITS DETECTION
        # ===================================================================
        # 401(k) is a US-only retirement plan; its presence strongly indicates
        # a US-based role.  Check for it when no GBP salary is mentioned.
        if re.search(r'401\s*\(?\s*k\s*\)?', text_lower):
            if not re.search(GBP_SALARY_PATTERN, text):
                return text, False, "Location: Likely US-based (401k benefits, no GBP)"

        # ===================================================================
        # PASSED ALL CHECKS
        # ===================================================================

        if len(text) > DESCRIPTION_MAX_CHARS:
            text = text[:DESCRIPTION_MAX_CHARS] + "... [truncated]"

        return text, True, "Valid: London/Remote + Trading/Infrastructure"

    except Exception as e:
        return f"Error fetching description: {str(e)[:100]}", False, "Error fetching"

# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    print("=" * 80)
    print("2-STAGE JOB SEARCH - FULLY AUTONOMOUS")
    print("=" * 80)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # Date suffix for output filenames (format: yyyymmdd)
    date_suffix = datetime.now().strftime('%Y%m%d')

    # Output directory for all result files
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'search_results')
    os.makedirs(output_dir, exist_ok=True)

    # Consolidated failure log
    all_failures = []

    # ========================================================================
    # STAGE 1: QUICK SCAN OF ALL COMPANIES
    # ========================================================================

    print("STAGE 1: QUICK SCAN (TITLES ONLY)")
    print("=" * 80)

    # Display all sites being searched
    tier1 = {k: v for k, v in COMPANIES.items() if v.get('type') == 'Hedge Fund / Prop Trading'}
    tier2 = {k: v for k, v in COMPANIES.items() if v.get('type') == 'FinTech'}
    tier3 = {k: v for k, v in COMPANIES.items() if v.get('type') == 'Recruitment Agency'}

    print(f"\nSITES BEING SEARCHED ({len(COMPANIES)} companies):")
    print(f"\n   TIER 1 -- Hedge Funds & Prop Trading ({len(tier1)} companies):")
    for name, data in tier1.items():
        print(f"      {name:<30s} {data['url'][:60]}")

    print(f"\n   TIER 2 -- FinTech & Trading Technology ({len(tier2)} companies):")
    for name, data in tier2.items():
        print(f"      {name:<30s} {data['url'][:60]}")

    if tier3:
        print(f"\n   TIER 3 -- Recruitment Agencies ({len(tier3)} companies):")
        for name, data in tier3.items():
            note_tag = f" ({data['note']})" if 'note' in data else ""
            print(f"      {name:<30s} {data['url'][:60]}{note_tag}")

    print()

    all_jobs = []
    failed_companies = []

    # Scrape company websites
    print(f"\nScraping {len(COMPANIES)} companies:")
    for company_name, company_data in COMPANIES.items():
        jobs, error = scrape_company_jobs(company_name, company_data)
        all_jobs.extend(jobs)
        if error:
            failed_companies.append(error)
            all_failures.append(error)

    # Score all jobs
    print(f"\nScoring & filtering...")
    for job in all_jobs:
        job['pass1_score'] = quick_score(job['title'], job.get('location', ''), job.get('company', ''), job.get('type', ''))

    # Filter to threshold
    filtered_jobs = [job for job in all_jobs if job['pass1_score'] >= MIN_SCORE_THRESHOLD]

    # Remove duplicates
    seen_urls = set()
    unique_jobs = []
    for job in filtered_jobs:
        if job['url'] not in seen_urls:
            unique_jobs.append(job)
            seen_urls.add(job['url'])

    # Sort by score
    unique_jobs.sort(key=lambda x: x['pass1_score'], reverse=True)

    print(f"   Total extracted: {len(all_jobs)}")
    print(f"   Scoring {MIN_SCORE_THRESHOLD}+: {len(filtered_jobs)}")
    print(f"   After dedup: {len(unique_jobs)}")

    # Save Stage 1 results
    if unique_jobs:
        df1 = pd.DataFrame(unique_jobs)
        pass1_filename = os.path.join(output_dir, f'pass1_quick_results-{date_suffix}.xlsx')
        df1.to_excel(pass1_filename, index=False, sheet_name='All Jobs')
        print(f"\nStage 1 saved: {pass1_filename}")

    # Save failed companies report
    if failed_companies:
        df_failed = pd.DataFrame(failed_companies)
        failed_companies_filename = os.path.join(output_dir, f'failed_companies-{date_suffix}.xlsx')
        df_failed.to_excel(failed_companies_filename, index=False, sheet_name='Failed Scrapes')
        print(f"Failed companies report saved: {failed_companies_filename} ({len(failed_companies)} companies)")
        print(f"    Recommendation: Manually check these {len(failed_companies)} companies' career pages")

    # ========================================================================
    # STAGE 2: FETCH FULL DESCRIPTIONS FOR TOP CANDIDATES
    # ========================================================================

    print(f"\n\nSTAGE 2: FETCH FULL DESCRIPTIONS + VALIDATION (TOP {STAGE2_TARGET_VALID})")
    print("=" * 80)

    # Use a larger pool to ensure we find enough valid ones after filtering
    candidate_pool = unique_jobs[:STAGE2_CANDIDATE_POOL]
    pool_size = len(candidate_pool)

    print(f"\nFetching and validating job descriptions for top {pool_size} candidates:")
    print("(Filtering for London/Remote + Trading Platform/Infrastructure domains)")

    validated_jobs = []
    rejected_jobs = []

    for i, job in enumerate(candidate_pool, 1):
        print(f"\n[{i}/{pool_size}] {job['company']} - {job['title'][:50]}...")
        print(f"        Fetching: {job['url'][:60]}...", end=" ")

        description, is_valid, reason = verify_job_location_and_domain(job['url'])
        job['full_description'] = description
        job['validation_status'] = reason

        if is_valid:
            validated_jobs.append(job)
            print(f"OK VALID ({len(description)} chars)")
        else:
            rejected_jobs.append(job)
            print(f"REJECTED ({reason})")
            if reason == "Error fetching":
                all_failures.append({
                    'company': job.get('company', 'Unknown'),
                    'url': job['url'],
                    'type': job.get('type', ''),
                    'error_classification': 'Stage 2 Fetch Error',
                    'http_status': None,
                    'status_reason': None,
                    'server_header': None,
                    'error_type': 'FetchError',
                    'error_detail': f"Job: {job['title'][:80]} | {description[:150]}",
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'source': job.get('source', 'Unknown'),
                })

        time.sleep(POLITE_DELAY)

        # Stop once we have enough valid jobs
        if len(validated_jobs) >= STAGE2_TARGET_VALID:
            print(f"\nFound {STAGE2_TARGET_VALID} valid jobs, stopping early")
            break

    top_results = validated_jobs[:STAGE2_TARGET_VALID]

    print(f"\n{'=' * 80}")
    print(f"VALIDATION SUMMARY:")
    print(f"{'=' * 80}")
    print(f"Checked: {len(validated_jobs) + len(rejected_jobs)} jobs")
    print(f"Passed validation: {len(validated_jobs)} jobs")
    print(f"Rejected: {len(rejected_jobs)} jobs")

    if rejected_jobs:
        print(f"\nRejection reasons:")
        rejection_counts = {}
        for job in rejected_jobs:
            reason = job['validation_status']
            rejection_counts[reason] = rejection_counts.get(reason, 0) + 1
        for reason, count in rejection_counts.items():
            print(f"  - {reason}: {count} jobs")

    # Save rejected jobs report
    if rejected_jobs:
        df_rejected = pd.DataFrame(rejected_jobs)
        rejected_columns = ['company', 'title', 'location', 'url', 'validation_status', 'pass1_score']
        df_rejected = df_rejected[rejected_columns]
        rejected_filename = os.path.join(output_dir, f'rejected_jobs-{date_suffix}.xlsx')
        df_rejected.to_excel(rejected_filename, index=False, sheet_name='Rejected Jobs')
        print(f"\nRejected jobs saved: {rejected_filename} ({len(rejected_jobs)} jobs)")
        print(f"   Review these if you want to see what was filtered out")

    # ========================================================================
    # SAVE CONSOLIDATED FAILURE LOG
    # ========================================================================

    if all_failures:
        df_all_failures = pd.DataFrame(all_failures)
        failure_columns = [
            'timestamp', 'source', 'company', 'url', 'type',
            'error_classification', 'http_status', 'status_reason',
            'server_header', 'error_type', 'error_detail'
        ]
        available_cols = [c for c in failure_columns if c in df_all_failures.columns]
        df_all_failures = df_all_failures[available_cols]
        failed_searches_filename = os.path.join(output_dir, f'failed_web_searches-{date_suffix}.xlsx')
        df_all_failures.to_excel(failed_searches_filename, index=False, sheet_name='All Failures')

        print(f"\n{'=' * 80}")
        print(f"FAILURE LOG SUMMARY:")
        print(f"{'=' * 80}")
        print(f"Total failures: {len(all_failures)}")

        class_counts = {}
        for f in all_failures:
            cls = f.get('error_classification', 'Unknown')
            class_counts[cls] = class_counts.get(cls, 0) + 1
        for cls, count in sorted(class_counts.items(), key=lambda x: -x[1]):
            print(f"  - {cls}: {count}")

        source_counts = {}
        for f in all_failures:
            src = f.get('source', 'Unknown')
            source_counts[src] = source_counts.get(src, 0) + 1
        print(f"\nFailures by source:")
        for src, count in sorted(source_counts.items(), key=lambda x: -x[1]):
            print(f"  - {src}: {count}")

        print(f"\nFull failure log saved: failed_web_searches-{date_suffix}.xlsx")

    # ========================================================================
    # SAVE STAGE 2 RESULTS (with Failed Companies reference sheet)
    # ========================================================================

    output_file = os.path.join(output_dir, f'pass2_top20_detailed-{date_suffix}.xlsx')

    if top_results:
        df2 = pd.DataFrame(top_results)
        columns_order = ['pass1_score', 'company', 'title', 'location', 'url', 'type', 'source', 'full_description']
        df2 = df2[columns_order]
    else:
        df2 = pd.DataFrame(columns=['pass1_score', 'company', 'title', 'location', 'url', 'type', 'source', 'full_description'])

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df2.to_excel(writer, index=False, sheet_name='Top 20 Detailed')

        if failed_companies:
            fc_summary = []
            for fc in failed_companies:
                fc_summary.append({
                    'Company': fc['company'],
                    'Career Page URL': fc['url'],
                    'Type': fc.get('type', ''),
                    'Error': fc.get('error_classification', fc.get('error', 'Unknown')),
                    'HTTP Status': fc.get('http_status', ''),
                    'Action': 'CHECK MANUALLY -- career page could not be scraped',
                })
            df_fc = pd.DataFrame(fc_summary)
            df_fc.to_excel(writer, index=False, sheet_name='Failed Companies')

    print(f"\nStage 2 saved: {output_file}")
    if failed_companies:
        print(f"   (includes 'Failed Companies' sheet with {len(failed_companies)} companies to check manually)")

    # ========================================================================
    # SUMMARY
    # ========================================================================

    print("\n\n" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    print(f"\nStage 1: {len(unique_jobs)} jobs found (saved to pass1_quick_results-{date_suffix}.xlsx)")
    print(f"Stage 2: {len(top_results)} valid jobs with full descriptions (saved to {output_file})")

    if rejected_jobs:
        print(f"Rejected: {len(rejected_jobs)} jobs failed location/domain validation (saved to rejected_jobs-{date_suffix}.xlsx)")

    if all_failures:
        print(f"Failures: {len(all_failures)} total web search failures (saved to failed_web_searches-{date_suffix}.xlsx)")

    if failed_companies:
        print(f"\n{len(failed_companies)} companies could not be accessed")
        print(f"   Companies blocked: {', '.join([f['company'] for f in failed_companies[:5]])}")
        if len(failed_companies) > 5:
            print(f"   ... and {len(failed_companies) - 5} more")
        print(f"   These are also listed in the 'Failed Companies' sheet of {output_file}")

    print(f"\nTOP 10 CANDIDATES:")
    for i, job in enumerate(unique_jobs[:10], 1):
        print(f"\n{i}. [{job['pass1_score']}] {job['company']}")
        print(f"   {job['title']}")
        print(f"   {job['location']}")

    print("\n\n" + "=" * 80)
    print("NEXT STEPS:")
    print("=" * 80)
    print(f"\n1. Open: {output_file}")
    print("2. Review the top 20 jobs with full descriptions")
    print(f"3. Check the 'Failed Companies' sheet -- {len(failed_companies)} companies need manual review")
    print("4. Select your top 10-15 for Claude analysis")

    if failed_companies:
        print(f"\n5. MANUALLY CHECK FAILED COMPANIES:")
        print(f"   Open the 'Failed Companies' sheet in {output_file} (or failed_companies-{date_suffix}.xlsx)")
        print(f"   Visit each company's career page directly")
        print(f"   These sites may have anti-bot protection or require JavaScript")
        print(f"   Look for PM/Platform/TPM roles and add to your list if found")

    next_step = 6 if failed_companies else 5
    print(f"\n{next_step}. Paste to Claude:")
    print("   'Analyse these job descriptions from pass2_top20_detailed.xlsx")
    print("    and provide:")
    print("    - Technical match score (0-100)")
    print("    - Role match score (0-100)")
    print("    - Seniority match score (0-100)")
    print("    - Key matches & gaps")
    print("    - Apply/Consider/Skip recommendation'")

    if all_failures:
        print(f"\n{next_step + 1}. REVIEW FAILURE LOG (optional):")
        print(f"   Open: failed_web_searches-{date_suffix}.xlsx")
        print(f"   Analyse error patterns to improve future scraping runs")

    print("\n" + "=" * 80)
    print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    main()
