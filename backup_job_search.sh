#!/bin/bash
# =============================================================================
# Job Search Pipeline — Manual Backup Script
# =============================================================================
# Creates a timestamped backup of all files needed to run the job search pipeline.
#
# Usage:
#   ./backup_job_search.sh
#
# You will be prompted to enter a version number (e.g. v3.3).
# The backup folder will be created as:  backup_v3.3_20260303/
# =============================================================================

set -eo pipefail

# ── Location ─────────────────────────────────────────────────────────────────
# The directory where the job search pipeline lives.
# Change this if you ever move the project to a different folder.
PROJECT_DIR="/Users/arek/Documents/Claude_Code/job_search"

# ── Prompt for version number ────────────────────────────────────────────────
echo ""
echo "=== Job Search Pipeline — Backup ==="
echo ""
read -rp "Enter version number (e.g. v3.3): " VERSION

# Validate that something was entered
if [ -z "$VERSION" ]; then
    echo "ERROR: No version number entered. Exiting."
    exit 1
fi

# Add "v" prefix if the user forgot it (e.g. "3.3" becomes "v3.3")
if [[ ! "$VERSION" == v* ]]; then
    VERSION="v${VERSION}"
fi

# ── Build backup folder name ─────────────────────────────────────────────────
TODAY=$(date +%Y%m%d)
BACKUP_DIR="${PROJECT_DIR}/backup_${VERSION}_${TODAY}"

# Check if backup folder already exists
if [ -d "$BACKUP_DIR" ]; then
    echo ""
    echo "WARNING: Backup folder already exists:"
    echo "  $BACKUP_DIR"
    read -rp "Overwrite? (y/n): " CONFIRM
    if [[ ! "$CONFIRM" =~ ^[Yy]$ ]]; then
        echo "Backup cancelled."
        exit 0
    fi
fi

# Create the backup folder
mkdir -p "$BACKUP_DIR"
echo ""
echo "Backing up to: $BACKUP_DIR"
echo ""

# ── Files to back up ────────────────────────────────────────────────────────
# Python scripts
COPIED=0
FAILED=0

copy_file() {
    local src="$1"
    if [ -f "$src" ]; then
        cp "$src" "$BACKUP_DIR/"
        echo "  ✓ $(basename "$src")"
        COPIED=$((COPIED + 1))
    else
        echo "  ✗ $(basename "$src") — NOT FOUND (skipped)"
        FAILED=$((FAILED + 1))
    fi
}

echo "Python scripts:"
copy_file "${PROJECT_DIR}/job_scraper_2stage.py"
copy_file "${PROJECT_DIR}/phase2_scorer.py"

echo ""
echo "Shell scripts:"
copy_file "${PROJECT_DIR}/run_daily_job_search.sh"
copy_file "${PROJECT_DIR}/run_job_search_now.sh"
copy_file "${PROJECT_DIR}/start_scheduler.sh"
copy_file "${PROJECT_DIR}/stop_scheduler.sh"

echo ""
echo "Scheduler config:"
copy_file "${PROJECT_DIR}/com.arek.daily-job-search.plist"

echo ""
echo "Environment file:"
copy_file "${PROJECT_DIR}/.env"

echo ""
echo "User guide:"
copy_file "${PROJECT_DIR}/Job_Scraper_User_Guide.docx"

# Target companies spreadsheet — there may be dated versions, copy all matches
echo ""
echo "Target companies spreadsheet(s):"
FOUND_COMPANIES=0
for f in "${PROJECT_DIR}"/Job_Search_Target_Companies*.xlsx; do
    if [ -f "$f" ]; then
        cp "$f" "$BACKUP_DIR/"
        echo "  ✓ $(basename "$f")"
        COPIED=$((COPIED + 1))
        FOUND_COMPANIES=$((FOUND_COMPANIES + 1))
    fi
done
if [ "$FOUND_COMPANIES" -eq 0 ]; then
    echo "  ✗ No target companies spreadsheet found (skipped)"
    FAILED=$((FAILED + 1))
fi

# CV files — copy all matching CV files
echo ""
echo "CV file(s):"
FOUND_CVS=0
for f in "${PROJECT_DIR}"/CV\ -\ Arek\ Baranowski*.docx; do
    if [ -f "$f" ]; then
        cp "$f" "$BACKUP_DIR/"
        echo "  ✓ $(basename "$f")"
        COPIED=$((COPIED + 1))
        FOUND_CVS=$((FOUND_CVS + 1))
    fi
done
if [ "$FOUND_CVS" -eq 0 ]; then
    echo "  ✗ No CV files found (skipped)"
    FAILED=$((FAILED + 1))
fi

# ── Summary ──────────────────────────────────────────────────────────────────
echo ""
echo "─────────────────────────────────────"
echo "Backup complete: ${COPIED} files copied"
if [ "$FAILED" -gt 0 ]; then
    echo "WARNING: ${FAILED} file(s) not found — see above"
fi
echo "Location: ${BACKUP_DIR}"
echo "─────────────────────────────────────"
echo ""
