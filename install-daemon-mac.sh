#!/bin/bash
# Install Liquor-Bot as a macOS LaunchAgent daemon
# Usage: ./install-daemon-mac.sh
#
# The bot runs headless with auto-restart. First run `python bot_script.py`
# manually to complete 2FA login and save auth_state.json.

set -e

BOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PLIST_NAME="com.liquorbot.daemon"
PLIST_PATH="$HOME/Library/LaunchAgents/${PLIST_NAME}.plist"
LOG_DIR="$HOME/Library/Logs/liquor-bot"
VENV_PYTHON="${BOT_DIR}/venv/bin/python3"

# Verify venv exists
if [ ! -f "$VENV_PYTHON" ]; then
    echo "Error: Python venv not found at ${VENV_PYTHON}"
    echo "Run: cd ${BOT_DIR} && python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt && playwright install chromium"
    exit 1
fi

# Verify auth state exists
if [ ! -f "${BOT_DIR}/auth_state.json" ]; then
    echo "Warning: auth_state.json not found."
    echo "Run the bot manually first to complete 2FA login:"
    echo "  cd ${BOT_DIR} && source venv/bin/activate && python bot_script.py"
    echo ""
    read -p "Continue anyway? (y/N) " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

# Create log directory
mkdir -p "$LOG_DIR"

# Unload existing if present
if launchctl list | grep -q "$PLIST_NAME" 2>/dev/null; then
    echo "Unloading existing daemon..."
    launchctl unload "$PLIST_PATH" 2>/dev/null || true
fi

# Write plist
cat > "$PLIST_PATH" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${PLIST_NAME}</string>
    <key>ProgramArguments</key>
    <array>
        <string>${VENV_PYTHON}</string>
        <string>${BOT_DIR}/bot_script.py</string>
        <string>--daemon</string>
    </array>
    <key>WorkingDirectory</key>
    <string>${BOT_DIR}</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>${LOG_DIR}/bot.log</string>
    <key>StandardErrorPath</key>
    <string>${LOG_DIR}/bot-error.log</string>
    <key>EnvironmentVariables</key>
    <dict>
        <key>PATH</key>
        <string>/usr/local/bin:/usr/bin:/bin:${BOT_DIR}/venv/bin</string>
        <key>HOME</key>
        <string>${HOME}</string>
    </dict>
    <key>ThrottleInterval</key>
    <integer>30</integer>
</dict>
</plist>
PLIST

echo "LaunchAgent plist written to: $PLIST_PATH"

# Load
launchctl load "$PLIST_PATH"
echo ""
echo "Liquor-Bot daemon installed and started!"
echo ""
echo "Commands:"
echo "  Status:  launchctl list | grep liquorbot"
echo "  Logs:    tail -f ${LOG_DIR}/bot.log"
echo "  Errors:  tail -f ${LOG_DIR}/bot-error.log"
echo "  Stop:    launchctl unload ${PLIST_PATH}"
echo "  Start:   launchctl load ${PLIST_PATH}"
echo "  Remove:  launchctl unload ${PLIST_PATH} && rm ${PLIST_PATH}"
