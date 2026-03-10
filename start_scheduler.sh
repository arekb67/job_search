#!/bin/bash
# =============================================================================
# START SCHEDULER — Activates the daily job search schedule (Mon-Fri 10am)
#
# Copies the launchd plist to ~/Library/LaunchAgents and loads it.
# The job search pipeline will then run automatically every weekday at 10:00 AM.
# Run this once to set up the schedule. It persists across reboots.
# =============================================================================

PLIST_NAME="com.arek.daily-job-search.plist"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE="$SCRIPT_DIR/$PLIST_NAME"
DEST="$HOME/Library/LaunchAgents/$PLIST_NAME"

# Check the plist file exists
if [ ! -f "$SOURCE" ]; then
    echo "ERROR: $SOURCE not found."
    exit 1
fi

# Create LaunchAgents directory if it doesn't exist
mkdir -p "$HOME/Library/LaunchAgents"

# Copy plist to LaunchAgents
cp "$SOURCE" "$DEST"
echo "Copied $PLIST_NAME to ~/Library/LaunchAgents/"

# Unload first in case it's already loaded (ignore errors)
launchctl unload "$DEST" 2>/dev/null

# Load the schedule
launchctl load "$DEST"
echo "Schedule loaded successfully."
echo ""
echo "The job search pipeline will now run automatically:"
echo "  When: Monday-Friday at 10:00 AM"
echo "  What: Phase 1 (scraping) + Phase 2 (Claude API scoring)"
echo "  Output: /Users/arek/Documents/GitHub repos/job_search/"
echo ""
echo "To verify: launchctl list | grep daily-job-search"
echo "To stop:   ./stop_scheduler.sh"
