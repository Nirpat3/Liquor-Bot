#!/usr/bin/env python3
"""
Mississippi DOR Order Bot - GUI Launcher
Easy-to-use interface for automated ordering
"""

import sys
import subprocess
import os
from pathlib import Path

def check_requirements():
    """Check if required packages are installed"""
    required = ['playwright', 'dotenv', 'flask']
    missing = []

    for package in required:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)

    return missing

def install_requirements():
    """Install missing requirements"""
    print("Installing required packages...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

    print("Installing browser...")
    subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])

def main():
    print("Mississippi DOR Order Bot")
    print("-" * 30)

    missing = check_requirements()
    if missing:
        print(f"Missing packages: {', '.join(missing)}")
        install_requirements()

    try:
        # Try Tkinter GUI
        from gui_bot import main as run_gui
        run_gui()
    except ImportError:
        print("Tkinter not available. Starting web interface...")
        from web_gui import app
        import webbrowser
        webbrowser.open('http://127.0.0.1:5050')
        app.run(debug=False, port=5050, host='0.0.0.0')

if __name__ == "__main__":
    main()