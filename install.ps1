#
# Mississippi DOR Order Bot — Windows Installer (PowerShell)
#
# This script:
#   1. Installs Python 3 (if missing)
#   2. Clones/updates the bot repo
#   3. Creates a virtual environment and installs all dependencies
#   4. Installs Playwright Chromium browser
#   5. Creates a Desktop shortcut
#
# Usage (one-liner in PowerShell):
#   irm https://raw.githubusercontent.com/krishp0130/Liquor-Bot/windows/install.ps1 | iex
#
# Or download first:
#   Invoke-WebRequest -Uri https://raw.githubusercontent.com/krishp0130/Liquor-Bot/windows/install.ps1 -OutFile install.ps1
#   .\install.ps1
#

$ErrorActionPreference = "Stop"

$RepoUrl    = "https://github.com/krishp0130/Liquor-Bot.git"
$InstallDir = "$env:USERPROFILE\LiquorBot"
$AppName    = "Liquor Bot"
$TotalSteps = 6

function Write-Step ($step, $msg) { Write-Host "`n[$step/$TotalSteps] $msg" -ForegroundColor Cyan }
function Write-Ok   ($msg)        { Write-Host "  ✓ $msg" -ForegroundColor Green }
function Write-Warn ($msg)        { Write-Host "  ⚠ $msg" -ForegroundColor Yellow }
function Write-Err  ($msg)        { Write-Host "  ✗ $msg" -ForegroundColor Red }

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor White
Write-Host "║   Mississippi DOR Order Bot — Windows Installer   ║" -ForegroundColor White
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor White
Write-Host ""

# ── Step 1: Python 3 ──
Write-Step 1 "Checking for Python 3..."

$pythonCmd = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $ver = & $cmd --version 2>&1
        if ($ver -match "Python 3") {
            $pythonCmd = $cmd
            Write-Ok "Python 3 found ($ver) via '$cmd'"
            break
        }
    } catch {}
}

if (-not $pythonCmd) {
    Write-Warn "Python 3 not found. Installing via winget..."
    try {
        winget install Python.Python.3.13 --accept-package-agreements --accept-source-agreements
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        $pythonCmd = "python"
        Write-Ok "Python 3 installed"
    } catch {
        Write-Err "Could not install Python automatically."
        Write-Host "  Please install Python 3 from https://www.python.org/downloads/" -ForegroundColor Yellow
        Write-Host "  Make sure to check 'Add Python to PATH' during installation." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Verify pip
try {
    & $pythonCmd -m pip --version | Out-Null
    Write-Ok "pip is available"
} catch {
    Write-Warn "pip not found. Installing..."
    & $pythonCmd -m ensurepip --upgrade
}

# ── Step 2: Git ──
Write-Step 2 "Checking for Git..."

$gitAvailable = $false
try {
    git --version | Out-Null
    $gitAvailable = $true
    Write-Ok "Git is already installed"
} catch {
    Write-Warn "Git not found. Installing via winget..."
    try {
        winget install Git.Git --accept-package-agreements --accept-source-agreements
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        $gitAvailable = $true
        Write-Ok "Git installed"
    } catch {
        Write-Err "Could not install Git automatically."
        Write-Host "  Please install Git from https://git-scm.com/download/win" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# ── Step 3: Clone/Update Repo ──
Write-Step 3 "Setting up bot files in $InstallDir..."

if (Test-Path "$InstallDir\.git") {
    Write-Ok "Bot directory already exists. Pulling latest changes..."
    Push-Location $InstallDir
    try {
        git pull origin windows
        Write-Ok "Updated to latest version"
    } catch {
        Write-Warn "Could not pull latest (offline or conflict). Using existing files."
    }
    Pop-Location
} else {
    if (Test-Path $InstallDir) {
        Write-Warn "Directory exists but is not a git repo. Backing up..."
        $backup = "${InstallDir}_backup_$(Get-Date -Format 'yyyyMMddHHmmss')"
        Rename-Item $InstallDir $backup
    }
    git clone -b windows $RepoUrl $InstallDir
    Write-Ok "Repository cloned"
}

# ── Step 4: Virtual Environment & Dependencies ──
Write-Step 4 "Creating virtual environment and installing dependencies..."

$venvPath = "$InstallDir\venv"
$venvPython = "$venvPath\Scripts\python.exe"
$venvPip = "$venvPath\Scripts\pip.exe"
$venvPlaywright = "$venvPath\Scripts\playwright.exe"

if (-not (Test-Path $venvPython)) {
    & $pythonCmd -m venv $venvPath
    Write-Ok "Virtual environment created"
} else {
    Write-Ok "Virtual environment already exists"
}

& $venvPip install --upgrade pip --quiet
& $venvPip install -r "$InstallDir\requirements.txt" --quiet
Write-Ok "Python dependencies installed"

# ── Step 5: Playwright Chromium ──
Write-Step 5 "Installing Playwright Chromium browser..."

# Clear old browser cache to ensure exact version match
$pwCache = "$env:LOCALAPPDATA\ms-playwright"
if (Test-Path $pwCache) {
    Remove-Item -Recurse -Force $pwCache
    Write-Ok "Cleared old browser cache"
}

& $venvPlaywright install chromium
Write-Ok "Chromium browser installed"

# ── Step 6: Create Desktop Shortcut & Data Folders ──
Write-Step 6 "Creating desktop shortcut and data folders..."

# Create data folders
New-Item -ItemType Directory -Force -Path "$InstallDir\Order Data\FutureSPA" | Out-Null
New-Item -ItemType Directory -Force -Path "$InstallDir\Order Data\CurrentPrices" | Out-Null
Write-Ok "Order Data folders ready"

# Create .env template if it doesn't exist
if (-not (Test-Path "$InstallDir\.env")) {
    @"
SITE_USERNAME=
SITE_PASSWORD=
SITE_URL=https://tap.dor.ms.gov/
HEADLESS=False
"@ | Set-Content "$InstallDir\.env"
    Write-Ok "Created .env template (enter credentials via the app)"
}

# Create orders.csv if it doesn't exist
if (-not (Test-Path "$InstallDir\orders.csv")) {
    "item_number,quantity,order_filled" | Set-Content "$InstallDir\orders.csv"
    Write-Ok "Created empty orders.csv"
}

# Create specialorder.csv if it doesn't exist
if (-not (Test-Path "$InstallDir\specialorder.csv")) {
    "item_number,quantity,name,order_number,order_date" | Set-Content "$InstallDir\specialorder.csv"
    Write-Ok "Created empty specialorder.csv"
}

# Create launcher batch file
$launcherPath = "$InstallDir\launch.bat"
@"
@echo off
cd /d "$InstallDir"
call venv\Scripts\activate.bat
start "" http://127.0.0.1:5050
python web_gui.py
"@ | Set-Content $launcherPath
Write-Ok "Launcher script created"

# Create Desktop shortcut
$desktopPath = [System.Environment]::GetFolderPath("Desktop")
$shortcutPath = "$desktopPath\$AppName.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcherPath
$shortcut.WorkingDirectory = $InstallDir
$shortcut.Description = "Mississippi DOR Order Bot"
$shortcut.IconLocation = "shell32.dll,21"
$shortcut.Save()
Write-Ok "Desktop shortcut created: $shortcutPath"

# Create uninstall script
@"
@echo off
echo Uninstalling Liquor Bot...
taskkill /f /im python.exe /fi "WINDOWTITLE eq web_gui*" 2>nul
if exist "$env:USERPROFILE\Desktop\Liquor Bot.lnk" del "$env:USERPROFILE\Desktop\Liquor Bot.lnk"
if exist "$InstallDir\orders.csv" copy "$InstallDir\orders.csv" "$env:USERPROFILE\Desktop\orders_backup.csv"
echo Backed up orders.csv to Desktop
rmdir /s /q "$InstallDir"
echo Liquor Bot has been uninstalled.
pause
"@ | Set-Content "$InstallDir\uninstall.bat"
Write-Ok "Uninstall script created"

# ── Done ──
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║          Installation Complete!                   ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "  App location:    ~\Desktop\$AppName.lnk" -ForegroundColor White
Write-Host "  Bot files:       $InstallDir\" -ForegroundColor White
Write-Host "  To uninstall:    $InstallDir\uninstall.bat" -ForegroundColor White
Write-Host ""
Write-Host "  Getting started:" -ForegroundColor White
Write-Host "    1. Double-click '$AppName' on your Desktop" -ForegroundColor White
Write-Host "    2. Enter your credentials in the Settings tab" -ForegroundColor White
Write-Host "    3. Add items in the Orders tab" -ForegroundColor White
Write-Host "    4. Click 'Start Bot' in the Control tab" -ForegroundColor White
Write-Host ""

$launch = Read-Host "Launch Liquor Bot now? (y/n)"
if ($launch -eq "y" -or $launch -eq "Y") {
    Start-Process $launcherPath
}
