"""
config_loader.py — Shared configuration loader for the job search pipeline.

Both job_scraper_2stage.py and phase2_scorer.py import this module to load
their settings from config.yaml.

Usage:
    from config_loader import load_config
    config = load_config()
"""

import os
import sys
import yaml


# All top-level sections that must exist in config.yaml
REQUIRED_SECTIONS = [
    "scraper_settings",
    "target_job_titles",
    "exclude_keywords",
    "preferred_locations",
    "location_verification",
    "domain_verification",
    "salary_detection",
    "target_companies",
    "company_type_categories",
    "scorer_settings",
    "candidate_profile",
    "industry_tiers",
    "location_scores",
    "recommendations",
]


def load_config(config_path=None):
    """
    Load and validate config.yaml.

    If config_path is not provided, looks for config.yaml in the same
    directory as the calling script.

    Returns:
        dict: The full configuration dictionary.

    Exits with code 1 if the config file is missing, has invalid YAML
    syntax, or is missing required sections.
    """
    if config_path is None:
        # Look for config.yaml next to the script that called us
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(script_dir, "config.yaml")

    # --- Check file exists ---
    if not os.path.isfile(config_path):
        print(f"\nERROR: Configuration file not found:")
        print(f"  {config_path}")
        print(f"\nPlease ensure config.yaml exists in the same folder as the scripts.")
        sys.exit(1)

    # --- Load and parse YAML ---
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
    except yaml.YAMLError as e:
        print(f"\nERROR: config.yaml has invalid syntax:")
        print(f"  {e}")
        print(f"\nPlease check the file for formatting errors (wrong indentation,")
        print(f"missing colons, unmatched quotes, or tabs instead of spaces).")
        sys.exit(1)

    if config is None:
        print(f"\nERROR: config.yaml is empty.")
        sys.exit(1)

    # --- Validate required sections exist ---
    missing = [s for s in REQUIRED_SECTIONS if s not in config]
    if missing:
        print(f"\nERROR: config.yaml is missing these required sections:")
        for s in missing:
            print(f"  - {s}")
        print(f"\nPlease add the missing sections to config.yaml.")
        sys.exit(1)

    return config
