#!/bin/bash
# =============================================================================
# DAILY JOB SEARCH — Phase 1 (Scraper) + Phase 2 (CV Scorer)
# =============================================================================
#
# Runs the full two-phase job search pipeline:
#   Phase 1: Python scraper visits 59 company career pages + Jooble API
#   Phase 2: Autonomous CV-matching scorer analyses results against CV
#
# All output files are written to the search_results/ subfolder with date suffixes.
#
# Usage:
#   ./run_daily_job_search.sh              # run both phases
#   ./run_daily_job_search.sh --phase1     # run Phase 1 only
#   ./run_daily_job_search.sh --phase2     # run Phase 2 only
#
# =============================================================================

set -eo pipefail

SCRIPT_DIR="/Users/arek/Documents/GitHub repos/job_search"
VENV_DIR="$SCRIPT_DIR/venv"
LOG_DIR="$SCRIPT_DIR/logs"
DATE=$(date +%Y%m%d)
LOG_FILE="$LOG_DIR/job_search_${DATE}.log"

# Create logs directory if it doesn't exist
mkdir -p "$LOG_DIR"

# Activate virtual environment
source "$VENV_DIR/bin/activate"

cd "$SCRIPT_DIR"

echo "============================================================" | tee -a "$LOG_FILE"
echo "DAILY JOB SEARCH — $(date '+%Y-%m-%d %H:%M:%S')" | tee -a "$LOG_FILE"
echo "============================================================" | tee -a "$LOG_FILE"

RUN_PHASE1=true
RUN_PHASE2=true

if [ "$1" = "--phase1" ]; then
    RUN_PHASE2=false
elif [ "$1" = "--phase2" ]; then
    RUN_PHASE1=false
fi

# Phase 1: Job Scraper
if [ "$RUN_PHASE1" = true ]; then
    echo "" | tee -a "$LOG_FILE"
    echo "PHASE 1: Running job scraper..." | tee -a "$LOG_FILE"
    echo "------------------------------------------------------------" | tee -a "$LOG_FILE"

    if python3 "$SCRIPT_DIR/job_scraper_2stage.py" 2>&1 | tee -a "$LOG_FILE"; then
        echo "" | tee -a "$LOG_FILE"
        echo "Phase 1 completed successfully." | tee -a "$LOG_FILE"
    else
        echo "" | tee -a "$LOG_FILE"
        echo "ERROR: Phase 1 failed. Check log: $LOG_FILE" | tee -a "$LOG_FILE"
        exit 1
    fi
fi

# Phase 2: CV Matching Scorer
if [ "$RUN_PHASE2" = true ]; then
    echo "" | tee -a "$LOG_FILE"
    echo "PHASE 2: Running CV matching scorer..." | tee -a "$LOG_FILE"
    echo "------------------------------------------------------------" | tee -a "$LOG_FILE"

    if python3 "$SCRIPT_DIR/phase2_scorer.py" --date "$DATE" 2>&1 | tee -a "$LOG_FILE"; then
        echo "" | tee -a "$LOG_FILE"
        echo "Phase 2 completed successfully." | tee -a "$LOG_FILE"
    else
        echo "" | tee -a "$LOG_FILE"
        echo "ERROR: Phase 2 failed. Check log: $LOG_FILE" | tee -a "$LOG_FILE"
        exit 1
    fi
fi

echo "" | tee -a "$LOG_FILE"
echo "============================================================" | tee -a "$LOG_FILE"
echo "ALL DONE — $(date '+%Y-%m-%d %H:%M:%S')" | tee -a "$LOG_FILE"
echo "============================================================" | tee -a "$LOG_FILE"
echo "" | tee -a "$LOG_FILE"
echo "Output files:" | tee -a "$LOG_FILE"
ls -la "$SCRIPT_DIR/search_results"/*-${DATE}.xlsx 2>/dev/null | tee -a "$LOG_FILE"
echo "" | tee -a "$LOG_FILE"
echo "Log file: $LOG_FILE" | tee -a "$LOG_FILE"

deactivate
