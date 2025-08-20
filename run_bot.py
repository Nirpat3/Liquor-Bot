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
    required = ['playwright', 'python-dotenv']
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
    subprocess.check_call([sys.executable, "-m", "pip", "install", 
                          "playwright", "python-dotenv"])
    
    print("Installing browser...")
    subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
    
def main():
    print("Mississippi DOR Order Bot")
    print("-" * 30)
    
    # check requirements
    missing = check_requirements()
    
    if missing:
        print(f"Missing packages: {', '.join(missing)}")
        response = input("Install missing packages? (y/n): ")
        
        if response.lower() == 'y':
            install_requirements()
        else:
            print("Cannot run without required packages.")
            return
    
    # run the GUI
    print("Starting GUI...")
    
    # import and run
    from gui_bot import main as run_gui
    run_gui()

if __name__ == "__main__":
    main()