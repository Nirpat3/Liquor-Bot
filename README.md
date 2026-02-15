# Mississippi DOR Liquor Order Bot

Automates ordering from the Mississippi Department of Revenue liquor ordering site. Add items to your list and the bot continuously checks availability and places orders when items are in stock.

## Features

- **Web interface** – Add items, manage orders, start/stop the bot
- **Desktop GUI** (optional) – Tkinter-based app when available
- **Saved login** – Skips login after first successful auth (handles session expiry)
- **2FA/OTP support** – Complete phone OTP manually when prompted; bot waits and continues
- **Order reporting** – Logs when orders succeed or fail

## Quick Start

### 1. Install

```bash
git clone https://github.com/yourusername/Liquor-Bot.git
cd Liquor-Bot
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure

Copy the example config and add your credentials:

```bash
cp .env.example .env
# Edit .env with your username and password
```

Or use the Settings tab in the web interface after starting the app.

### 3. Run

```bash
python run_bot.py
```

Opens the web interface at http://localhost:5000 (or Tkinter GUI if available).

### 4. Use

1. Enter login credentials in **Settings** and save
2. Add items (item number + quantity) in **Orders**
3. Click **Start Bot** in the Control tab
4. If prompted for OTP, complete it in the browser
5. The bot runs continuously and reports order success/failure in the Logs tab

## CSV Format

`orders.csv` columns: `item_number`, `quantity`, `order_filled`

- Add items via the UI or edit the CSV directly
- `order_filled` is set to `yes` when the item has been ordered

## Requirements

- Python 3.8+
- Playwright (Chromium)
- Flask (for web UI)
