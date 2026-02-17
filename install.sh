#!/bin/bash
#
# Mississippi DOR Order Bot — macOS Installer
#
# This script:
#   1. Installs Homebrew (if missing)
#   2. Installs Python 3 via Homebrew (if missing)
#   3. Clones/updates the bot repo
#   4. Creates a virtual environment and installs all dependencies
#   5. Installs Playwright Chromium browser
#   6. Creates a "Liquor Bot.app" on the Desktop
#
# Usage (one-liner):
#   curl -fsSL https://raw.githubusercontent.com/krishp0130/Liquor-Bot/main/install.sh -o /tmp/install.sh && bash /tmp/install.sh
#
# Or download first:
#   curl -fsSL https://raw.githubusercontent.com/krishp0130/Liquor-Bot/main/install.sh -o install.sh
#   bash install.sh
#

set -e

# ── Refuse to run as root ──
if [[ "$EUID" -eq 0 ]]; then
    echo ""
    echo "ERROR: Do not run this installer with 'sudo'."
    echo ""
    echo "Homebrew (and this installer) must run as your normal user."
    echo ""
    echo "Instead, run:"
    echo "  curl -fsSL https://raw.githubusercontent.com/krishp0130/Liquor-Bot/main/install.sh -o install.sh && bash install.sh"
    echo ""
    exit 1
fi

# ── Colors ──
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
BOLD='\033[1m'
NC='\033[0m'

print_step() { echo -e "\n${BLUE}${BOLD}[$1/$TOTAL_STEPS]${NC} ${BOLD}$2${NC}"; }
print_ok()   { echo -e "  ${GREEN}✓${NC} $1"; }
print_warn() { echo -e "  ${YELLOW}⚠${NC} $1"; }
print_err()  { echo -e "  ${RED}✗${NC} $1"; }

TOTAL_STEPS=6
REPO_URL="https://github.com/krishp0130/Liquor-Bot.git"
INSTALL_DIR="$HOME/LiquorBot"
APP_NAME="Liquor Bot"
DESKTOP_APP="$HOME/Desktop/${APP_NAME}.app"

echo -e "${BOLD}"
echo "╔══════════════════════════════════════════════════╗"
echo "║   Mississippi DOR Order Bot — macOS Installer    ║"
echo "╚══════════════════════════════════════════════════╝"
echo -e "${NC}"

# ── Step 1: Homebrew ──
print_step 1 "Checking for Homebrew..."

if command -v brew &>/dev/null; then
    print_ok "Homebrew is already installed"
else
    print_warn "Homebrew not found. Installing..."
    NONINTERACTIVE=1 /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

    # Add Homebrew to PATH for Apple Silicon and Intel Macs
    if [[ -f /opt/homebrew/bin/brew ]]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
        # Persist for future shells
        if ! grep -q 'homebrew' "$HOME/.zprofile" 2>/dev/null; then
            echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> "$HOME/.zprofile"
        fi
    elif [[ -f /usr/local/bin/brew ]]; then
        eval "$(/usr/local/bin/brew shellenv)"
    fi
    print_ok "Homebrew installed"
fi

# ── Step 2: Python 3 ──
print_step 2 "Checking for Python 3..."

if command -v python3 &>/dev/null; then
    PY_VERSION=$(python3 --version 2>&1)
    print_ok "Python 3 is already installed ($PY_VERSION)"
else
    print_warn "Python 3 not found. Installing via Homebrew..."
    brew install python@3.13
    print_ok "Python 3 installed"
fi

# Verify pip
if ! python3 -m pip --version &>/dev/null; then
    print_warn "pip not found. Installing..."
    python3 -m ensurepip --upgrade
fi
print_ok "pip is available"

# ── Step 3: Clone/Update Repo ──
print_step 3 "Setting up bot files in ${INSTALL_DIR}..."

if [[ -d "$INSTALL_DIR/.git" ]]; then
    print_ok "Bot directory already exists. Pulling latest changes..."
    cd "$INSTALL_DIR"
    git pull origin main || print_warn "Could not pull latest (offline or conflict). Using existing files."
else
    if [[ -d "$INSTALL_DIR" ]]; then
        print_warn "Directory exists but is not a git repo. Backing up and re-cloning..."
        mv "$INSTALL_DIR" "${INSTALL_DIR}_backup_$(date +%s)"
    fi
    git clone "$REPO_URL" "$INSTALL_DIR"
    print_ok "Repository cloned"
    cd "$INSTALL_DIR"
fi

# Fix port: macOS AirPlay Receiver uses 5000, so switch to 5050
if grep -q "port=5000" "$INSTALL_DIR/web_gui.py" 2>/dev/null; then
    sed -i '' 's/port=5000/port=5050/g' "$INSTALL_DIR/web_gui.py"
    sed -i '' 's/localhost:5000/localhost:5050/g' "$INSTALL_DIR/web_gui.py"
    print_ok "Fixed Flask port (5000 → 5050, avoids macOS AirPlay conflict)"
fi
if grep -q "port=5000" "$INSTALL_DIR/run_bot.py" 2>/dev/null; then
    sed -i '' 's/port=5000/port=5050/g' "$INSTALL_DIR/run_bot.py"
    sed -i '' 's/5000/5050/g' "$INSTALL_DIR/run_bot.py"
fi

# ── Step 4: Virtual Environment & Dependencies ──
print_step 4 "Creating virtual environment and installing dependencies..."

if [[ ! -d "$INSTALL_DIR/venv" ]]; then
    python3 -m venv "$INSTALL_DIR/venv"
    print_ok "Virtual environment created"
else
    print_ok "Virtual environment already exists"
fi

source "$INSTALL_DIR/venv/bin/activate"

pip install --upgrade pip --quiet
pip install flask playwright python-dotenv --quiet
print_ok "Python dependencies installed"

# ── Step 5: Playwright Chromium ──
print_step 5 "Installing Playwright Chromium browser..."

playwright install chromium
print_ok "Chromium browser installed"

# ── Step 6: Create Desktop Application ──
print_step 6 "Creating ${APP_NAME}.app on Desktop..."

# Remove old .app if it exists
if [[ -d "$DESKTOP_APP" ]]; then
    rm -rf "$DESKTOP_APP"
fi

# Create .app bundle structure
mkdir -p "$DESKTOP_APP/Contents/MacOS"
mkdir -p "$DESKTOP_APP/Contents/Resources"

# Info.plist
cat > "$DESKTOP_APP/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>Liquor Bot</string>
    <key>CFBundleDisplayName</key>
    <string>Liquor Bot</string>
    <key>CFBundleIdentifier</key>
    <string>com.liquorbot.app</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>CFBundleExecutable</key>
    <string>launch</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>11.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
PLIST

# Launch script
cat > "$DESKTOP_APP/Contents/MacOS/launch" << LAUNCHER
#!/bin/bash
#
# Liquor Bot launcher — starts the web GUI and opens the browser
#

INSTALL_DIR="$INSTALL_DIR"
LOG_FILE="\$INSTALL_DIR/bot_launch.log"

# Activate virtual environment
source "\$INSTALL_DIR/venv/bin/activate"
cd "\$INSTALL_DIR"

# Load environment
export PATH="/opt/homebrew/bin:/usr/local/bin:\$PATH"

# Check if already running
if lsof -i :5050 -sTCP:LISTEN &>/dev/null; then
    # Already running, just open the browser
    open "http://127.0.0.1:5050"
    exit 0
fi

# Start the web GUI server in the background
python "\$INSTALL_DIR/web_gui.py" > "\$LOG_FILE" 2>&1 &
SERVER_PID=\$!

# Wait for server to be ready (up to 10 seconds)
for i in {1..20}; do
    if curl -s "http://127.0.0.1:5050" > /dev/null 2>&1; then
        break
    fi
    sleep 0.5
done

# Open browser
open "http://127.0.0.1:5050"

# Keep the process alive so the .app stays "open"
# When the user quits the app, the server will be killed
wait \$SERVER_PID
LAUNCHER

chmod +x "$DESKTOP_APP/Contents/MacOS/launch"

# Generate an app icon (simple colored square with text)
# Use built-in macOS tools to create an .icns from a PNG
ICON_DIR=$(mktemp -d)
export ICON_DIR

python3 << 'ICONSCRIPT'
import struct, zlib, os

def create_png(width, height, filepath):
    """Create a simple PNG icon with a gradient background and bottle shape."""

    def make_pixel(x, y, w, h):
        r = int(30 + (x / w) * 40)
        g = int(30 + (y / h) * 20)
        b = int(80 + (x / w) * 60 + (y / h) * 40)

        cx, cy = w // 2, h // 2
        body_w = w // 4
        body_h = int(h / 2.5)
        body_top = cy - int(body_h * 0.3)
        body_bot = cy + int(body_h * 0.7)
        body_left = cx - body_w // 2
        body_right = cx + body_w // 2

        neck_w = body_w // 3
        neck_top = body_top - int(body_h * 0.35)
        neck_left = cx - neck_w // 2
        neck_right = cx + neck_w // 2

        in_body = body_left <= x <= body_right and body_top <= y <= body_bot
        in_neck = neck_left <= x <= neck_right and neck_top <= y <= body_top

        if y > body_bot - body_w // 4:
            dist_from_center = abs(x - cx)
            curve_radius = body_w // 2
            y_offset = y - (body_bot - body_w // 4)
            if dist_from_center ** 2 + y_offset ** 2 > curve_radius ** 2:
                in_body = False

        if in_body or in_neck:
            r = 220
            g = min(255, 170 + int((y - neck_top) / max(1, body_bot - neck_top) * 30))
            b = 50
            if in_body and abs(y - cy) < body_h * 0.12:
                r, g, b = 255, 255, 240

        return bytes([min(255, max(0, r)), min(255, max(0, g)), min(255, max(0, b)), 255])

    raw_data = b''
    for y in range(height):
        raw_data += b'\x00'
        for x in range(width):
            raw_data += make_pixel(x, y, width, height)

    def chunk(chunk_type, data):
        c = chunk_type + data
        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)

    ihdr = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    png = b'\x89PNG\r\n\x1a\n'
    png += chunk(b'IHDR', ihdr)
    png += chunk(b'IDAT', zlib.compress(raw_data))
    png += chunk(b'IEND', b'')

    with open(filepath, 'wb') as f:
        f.write(png)

icon_dir = os.environ['ICON_DIR']
for size in [16, 32, 64, 128, 256, 512]:
    create_png(size, size, os.path.join(icon_dir, f'icon_{size}.png'))
print("Icons generated")
ICONSCRIPT

# Create .iconset and convert to .icns
ICONSET_DIR="$ICON_DIR/AppIcon.iconset"
mkdir -p "$ICONSET_DIR"
cp "$ICON_DIR/icon_16.png"  "$ICONSET_DIR/icon_16x16.png"
cp "$ICON_DIR/icon_32.png"  "$ICONSET_DIR/icon_16x16@2x.png"
cp "$ICON_DIR/icon_32.png"  "$ICONSET_DIR/icon_32x32.png"
cp "$ICON_DIR/icon_64.png"  "$ICONSET_DIR/icon_32x32@2x.png"
cp "$ICON_DIR/icon_128.png" "$ICONSET_DIR/icon_128x128.png"
cp "$ICON_DIR/icon_256.png" "$ICONSET_DIR/icon_128x128@2x.png"
cp "$ICON_DIR/icon_256.png" "$ICONSET_DIR/icon_256x256.png"
cp "$ICON_DIR/icon_512.png" "$ICONSET_DIR/icon_256x256@2x.png"
cp "$ICON_DIR/icon_512.png" "$ICONSET_DIR/icon_512x512.png"

iconutil -c icns "$ICONSET_DIR" -o "$DESKTOP_APP/Contents/Resources/AppIcon.icns" 2>/dev/null || \
    print_warn "Could not generate .icns icon (cosmetic only, app still works)"

rm -rf "$ICON_DIR"

# Create .env template if it doesn't exist
if [[ ! -f "$INSTALL_DIR/.env" ]]; then
    cat > "$INSTALL_DIR/.env" << 'ENVTEMPLATE'
SITE_USERNAME=
SITE_PASSWORD=
SITE_URL=https://tap.dor.ms.gov/
HEADLESS=False
ENVTEMPLATE
    print_ok "Created .env template (enter credentials via the app)"
fi

# Create orders.csv if it doesn't exist
if [[ ! -f "$INSTALL_DIR/orders.csv" ]]; then
    echo "item_number,quantity,order_filled" > "$INSTALL_DIR/orders.csv"
    print_ok "Created empty orders.csv"
fi

# Also create an uninstall script
cat > "$INSTALL_DIR/uninstall.sh" << 'UNINSTALL'
#!/bin/bash
echo "Uninstalling Liquor Bot..."
# Kill any running server
pkill -f "web_gui.py" 2>/dev/null
# Remove desktop app
rm -rf "$HOME/Desktop/Liquor Bot.app"
# Remove install directory (preserves orders.csv backup)
if [[ -f "$HOME/LiquorBot/orders.csv" ]]; then
    cp "$HOME/LiquorBot/orders.csv" "$HOME/Desktop/orders_backup.csv"
    echo "Backed up orders.csv to Desktop"
fi
rm -rf "$HOME/LiquorBot"
echo "Liquor Bot has been uninstalled."
UNINSTALL
chmod +x "$INSTALL_DIR/uninstall.sh"

# Strip quarantine attribute so macOS Gatekeeper doesn't block the app
xattr -cr "$DESKTOP_APP" 2>/dev/null || true
print_ok "Cleared macOS quarantine flags"

print_ok "Desktop application created at: $DESKTOP_APP"

# ── Done ──
echo ""
echo -e "${GREEN}${BOLD}╔══════════════════════════════════════════════════╗${NC}"
echo -e "${GREEN}${BOLD}║          Installation Complete!                   ║${NC}"
echo -e "${GREEN}${BOLD}╚══════════════════════════════════════════════════╝${NC}"
echo ""
echo -e "  ${BOLD}App location:${NC}    ~/Desktop/${APP_NAME}.app"
echo -e "  ${BOLD}Bot files:${NC}       ${INSTALL_DIR}/"
echo -e "  ${BOLD}To uninstall:${NC}    ${INSTALL_DIR}/uninstall.sh"
echo ""
echo -e "  ${BOLD}Getting started:${NC}"
echo -e "    1. Double-click '${APP_NAME}' on your Desktop"
echo -e "    2. Enter your credentials in the Settings tab"
echo -e "    3. Add items in the Orders tab"
echo -e "    4. Click 'Start Bot' in the Control tab"
echo ""
echo -e "  ${YELLOW}Note:${NC} On first launch, macOS may ask you to allow"
echo -e "  the app. Right-click → Open if it won't open normally."
echo ""

# Offer to launch now
read -p "Launch Liquor Bot now? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    open "$DESKTOP_APP"
fi
