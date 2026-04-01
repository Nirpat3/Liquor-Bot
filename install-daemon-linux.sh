#!/bin/bash
# Install Liquor-Bot as a Linux systemd service
# Usage: ./install-daemon-linux.sh
#
# The bot runs headless with auto-restart. First run `python bot_script.py`
# manually to complete 2FA login and save auth_state.json.

set -e

BOT_DIR="$(cd "$(dirname "$0")" && pwd)"
SERVICE_NAME="liquor-bot"
SERVICE_PATH="/etc/systemd/system/${SERVICE_NAME}.service"
VENV_PYTHON="${BOT_DIR}/venv/bin/python3"
CURRENT_USER="$(whoami)"

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

# Need sudo for systemd
if [ "$EUID" -ne 0 ]; then
    echo "Systemd service installation requires sudo."
    echo "Re-running with sudo..."
    exec sudo "$0" "$@"
fi

# Write service file
cat > "$SERVICE_PATH" << SERVICE
[Unit]
Description=Mississippi DOR Liquor Order Bot
After=network.target

[Service]
Type=simple
User=${CURRENT_USER}
WorkingDirectory=${BOT_DIR}
ExecStart=${VENV_PYTHON} ${BOT_DIR}/bot_script.py --daemon
Restart=always
RestartSec=30
Environment=PATH=/usr/local/bin:/usr/bin:/bin:${BOT_DIR}/venv/bin
Environment=HOME=/home/${CURRENT_USER}
Environment=DISPLAY=:99

# Auto-restart limits
StartLimitIntervalSec=300
StartLimitBurst=5

[Install]
WantedBy=multi-user.target
SERVICE

echo "Service file written to: $SERVICE_PATH"

# Reload and enable
systemctl daemon-reload
systemctl enable "$SERVICE_NAME"
systemctl start "$SERVICE_NAME"

echo ""
echo "Liquor-Bot daemon installed and started!"
echo ""
echo "Commands:"
echo "  Status:  systemctl status ${SERVICE_NAME}"
echo "  Logs:    journalctl -u ${SERVICE_NAME} -f"
echo "  Stop:    sudo systemctl stop ${SERVICE_NAME}"
echo "  Start:   sudo systemctl start ${SERVICE_NAME}"
echo "  Disable: sudo systemctl disable ${SERVICE_NAME}"
echo "  Remove:  sudo systemctl stop ${SERVICE_NAME} && sudo rm ${SERVICE_PATH} && sudo systemctl daemon-reload"
