#!/bin/bash
# =============================================================================
# RUN JOB SEARCH NOW — Manually triggers the full pipeline (Phase 1 + Phase 2)
#
# Use this to run the job search on demand, outside the normal schedule.
# Runs Phase 1 (scraping ~10 min) then Phase 2 (Claude API scoring ~2-3 min).
# All output files are written to this directory with today's date suffix.
# Logs are saved to the logs/ subdirectory.
#
# Usage:
#   ./run_job_search_now.sh              # run both phases
#   ./run_job_search_now.sh --phase1     # run Phase 1 only (scraping)
#   ./run_job_search_now.sh --phase2     # run Phase 2 only (Claude scoring)
# =============================================================================

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/run_daily_job_search.sh" "$@"
