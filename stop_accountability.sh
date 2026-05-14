#!/bin/bash
set -euo pipefail
PLIST="$HOME/Library/LaunchAgents/com.gunjan.accountability.plist"
launchctl unload "$PLIST" 2>/dev/null || true
echo "Stopped accountability tracker."
