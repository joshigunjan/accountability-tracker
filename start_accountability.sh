#!/bin/bash
set -euo pipefail

INTERVAL_MINUTES="${1:-30}"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
UV_BIN="$(command -v uv || true)"

if [ -z "$UV_BIN" ]; then
  echo "uv not found. Install it with: brew install uv"
  exit 1
fi

mkdir -p "$HOME/.accountability_tracker" "$HOME/Library/LaunchAgents"
RUNNER="$HOME/.accountability_tracker/run_accountability.sh"
PLIST="$HOME/Library/LaunchAgents/com.gunjan.accountability.plist"

cat > "$RUNNER" <<EOF
#!/bin/bash
cd "$REPO_DIR"
"$UV_BIN" run python accountability_prompt.py --scheduled --interval-minutes "$INTERVAL_MINUTES"
EOF
chmod +x "$RUNNER"

# Stop old agents from earlier SAP/autotime versions if present.
launchctl unload "$HOME/Library/LaunchAgents/com.gunjan.autotime.plist" 2>/dev/null || true
launchctl unload "$HOME/Library/LaunchAgents/com.gunjan.accountability.plist" 2>/dev/null || true

cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>Label</key>
    <string>com.gunjan.accountability</string>

    <key>ProgramArguments</key>
    <array>
      <string>/bin/bash</string>
      <string>$RUNNER</string>
    </array>

    <!-- Check every minute. The script itself decides whether to show a popup. -->
    <key>StartInterval</key>
    <integer>60</integer>

    <key>RunAtLoad</key>
    <true/>

    <key>StandardOutPath</key>
    <string>$HOME/.accountability_tracker/launchd.out.log</string>

    <key>StandardErrorPath</key>
    <string>$HOME/.accountability_tracker/launchd.err.log</string>
  </dict>
</plist>
EOF

plutil -lint "$PLIST"
launchctl load "$PLIST"

echo "Started accountability tracker. Visible prompt interval: ${INTERVAL_MINUTES} minutes."
echo "Data folder: $HOME/Documents/autotime_sap"
echo "Overtime workbook: $HOME/Documents/autotime_sap/overtime.xlsx"
