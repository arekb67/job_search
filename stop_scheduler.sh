#!/bin/bash
# =============================================================================
# STOP SCHEDULER — Deactivates the daily job search schedule
#
# Unloads the launchd plist so the pipeline no longer runs automatically.
# You can restart it at any time with ./start_scheduler.sh
# =============================================================================

PLIST_NAME="com.arek.daily-job-search.plist"
DEST="$HOME/Library/LaunchAgents/$PLIST_NAME"

if [ ! -f "$DEST" ]; then
    echo "Schedule is not installed (no plist found at $DEST)."
    exit 0
fi

launchctl unload "$DEST"
echo "Schedule stopped successfully."
echo ""
echo "The daily job search will no longer run automatically."
echo "To restart: ./start_scheduler.sh"
echo "To run manually: ./run_job_search_now.sh"
