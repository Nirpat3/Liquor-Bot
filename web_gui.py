# Save as web_gui.py
from flask import Flask, render_template_string, request, jsonify, send_file
import threading
import asyncio
import csv
import fcntl
import json
from pathlib import Path
import logging
import os
import re
from datetime import datetime
from dotenv import load_dotenv
import sys
from odf.opendocument import load as load_ods
from odf.table import Table, TableRow, TableCell
from odf.text import P

app = Flask(__name__)
bot_thread = None
bot_running = False
stop_event = threading.Event()

CSV_FIELDNAMES = ['item_number', 'quantity', 'name', 'size', 'units', 'order_filled']


def _read_csv_locked():
    """Read orders.csv with shared file lock."""
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as f:
            fcntl.flock(f.fileno(), fcntl.LOCK_SH)
            try:
                orders = list(csv.DictReader(f))
            finally:
                fcntl.flock(f.fileno(), fcntl.LOCK_UN)
    return orders


def _write_csv_locked(orders):
    """Write orders.csv with exclusive file lock."""
    with open('orders.csv', 'w', newline='') as f:
        fcntl.flock(f.fileno(), fcntl.LOCK_EX)
        try:
            writer = csv.DictWriter(f, fieldnames=CSV_FIELDNAMES, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(orders)
        finally:
            fcntl.flock(f.fileno(), fcntl.LOCK_UN)

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# HTML template
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MS DOR Order Bot</title>
    <style>
        :root {
            --bg-primary: #0f1117;
            --bg-secondary: #1a1d27;
            --bg-card: #21242f;
            --bg-input: #2a2d3a;
            --bg-hover: #2e3140;
            --border: #333750;
            --text-primary: #e4e6ef;
            --text-secondary: #8b8fa3;
            --text-muted: #5d6178;
            --accent: #6366f1;
            --accent-hover: #818cf8;
            --accent-glow: rgba(99,102,241,0.15);
            --success: #22c55e;
            --success-bg: rgba(34,197,94,0.12);
            --danger: #ef4444;
            --danger-bg: rgba(239,68,68,0.12);
            --warning: #f59e0b;
            --warning-bg: rgba(245,158,11,0.12);
            --info: #3b82f6;
            --radius: 10px;
            --radius-sm: 6px;
            --shadow: 0 4px 24px rgba(0,0,0,0.3);
            --transition: 0.2s ease;
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
        }
        .app-header {
            background: var(--bg-secondary);
            border-bottom: 1px solid var(--border);
            padding: 16px 24px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            position: sticky;
            top: 0;
            z-index: 100;
        }
        .app-title {
            font-size: 20px;
            font-weight: 700;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 10px;
        }
        .app-title .logo {
            width: 32px;
            height: 32px;
            background: linear-gradient(135deg, var(--accent), #a78bfa);
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
        }
        .status-pill {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 13px;
            font-weight: 600;
        }
        .status-pill.running { background: var(--success-bg); color: var(--success); }
        .status-pill.stopped { background: var(--danger-bg); color: var(--danger); }
        .status-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            display: inline-block;
        }
        .status-dot.running { background: var(--success); box-shadow: 0 0 8px var(--success); animation: pulse 2s infinite; }
        .status-dot.stopped { background: var(--danger); }
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.4; }
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }
        /* Tabs */
        .tabs {
            display: flex;
            gap: 4px;
            background: var(--bg-secondary);
            border-radius: var(--radius);
            padding: 4px;
            margin-bottom: 20px;
            overflow-x: auto;
            border: 1px solid var(--border);
        }
        .tab {
            padding: 10px 18px;
            cursor: pointer;
            border-radius: var(--radius-sm);
            font-size: 13px;
            font-weight: 500;
            color: var(--text-secondary);
            transition: var(--transition);
            white-space: nowrap;
            user-select: none;
        }
        .tab:hover { color: var(--text-primary); background: var(--bg-hover); }
        .tab.active {
            background: var(--accent);
            color: white;
            box-shadow: 0 2px 8px rgba(99,102,241,0.3);
        }
        .tab-content { display: none; animation: fadeIn 0.3s ease; }
        .tab-content.active { display: block; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }
        /* Cards */
        .card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 20px;
            margin-bottom: 16px;
        }
        .card-title {
            font-size: 15px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 16px;
        }
        /* Forms */
        input[type="text"], input[type="password"], input[type="number"] {
            background: var(--bg-input);
            border: 1px solid var(--border);
            color: var(--text-primary);
            padding: 10px 14px;
            border-radius: var(--radius-sm);
            font-size: 14px;
            outline: none;
            transition: var(--transition);
        }
        input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px var(--accent-glow); }
        input::placeholder { color: var(--text-muted); }
        .input-row {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
        }
        .input-row label {
            min-width: 100px;
            font-size: 14px;
            color: var(--text-secondary);
        }
        .input-row input { flex: 1; max-width: 360px; }
        /* Checkbox */
        .checkbox-row {
            display: flex;
            align-items: center;
            gap: 8px;
            margin: 12px 0;
            font-size: 14px;
            color: var(--text-secondary);
            cursor: pointer;
        }
        .checkbox-row input[type="checkbox"] {
            accent-color: var(--accent);
            width: 16px;
            height: 16px;
        }
        /* Buttons */
        .btn {
            padding: 9px 18px;
            border: none;
            border-radius: var(--radius-sm);
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        .btn-primary { background: var(--accent); color: white; }
        .btn-primary:hover { background: var(--accent-hover); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(99,102,241,0.3); }
        .btn-success { background: var(--success); color: white; }
        .btn-success:hover { background: #16a34a; }
        .btn-danger { background: var(--danger); color: white; }
        .btn-danger:hover { background: #dc2626; }
        .btn-ghost { background: var(--bg-input); color: var(--text-secondary); border: 1px solid var(--border); }
        .btn-ghost:hover { background: var(--bg-hover); color: var(--text-primary); }
        .btn-sm { padding: 5px 12px; font-size: 12px; }
        .btn-lg { padding: 14px 32px; font-size: 16px; }
        .btn-group { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px; }
        /* Tables */
        table { width: 100%; border-collapse: separate; border-spacing: 0; }
        thead th {
            background: var(--bg-secondary);
            color: var(--text-secondary);
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            padding: 10px 14px;
            text-align: left;
            border-bottom: 1px solid var(--border);
            position: sticky;
            top: 0;
            cursor: pointer;
            user-select: none;
        }
        thead th:hover { color: var(--accent); }
        tbody td {
            padding: 10px 14px;
            font-size: 14px;
            border-bottom: 1px solid rgba(51,55,80,0.5);
            transition: var(--transition);
        }
        tbody tr { transition: var(--transition); }
        tbody tr:hover { background: var(--bg-hover); }
        .badge {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            padding: 3px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }
        .badge-success { background: var(--success-bg); color: var(--success); }
        .badge-warning { background: var(--warning-bg); color: var(--warning); }
        .badge-danger { background: var(--danger-bg); color: var(--danger); }
        /* Stats grid */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 12px;
            margin-bottom: 20px;
        }
        .stat-card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px;
            text-align: center;
        }
        .stat-value {
            font-size: 32px;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .stat-label {
            font-size: 12px;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        /* Progress bar */
        .progress-bar {
            height: 8px;
            background: var(--bg-input);
            border-radius: 4px;
            overflow: hidden;
            margin-top: 12px;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, var(--accent), var(--success));
            border-radius: 4px;
            transition: width 0.5s ease;
        }
        /* Control panel */
        .control-center {
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 40px 20px;
        }
        .power-btn {
            width: 120px;
            height: 120px;
            border-radius: 50%;
            border: 3px solid var(--border);
            background: var(--bg-secondary);
            color: var(--text-secondary);
            font-size: 40px;
            cursor: pointer;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 20px;
        }
        .power-btn:hover { border-color: var(--accent); color: var(--accent); transform: scale(1.05); }
        .power-btn.running {
            border-color: var(--success);
            color: var(--success);
            box-shadow: 0 0 30px rgba(34,197,94,0.2);
            animation: pulseGlow 2s infinite;
        }
        @keyframes pulseGlow {
            0%, 100% { box-shadow: 0 0 20px rgba(34,197,94,0.15); }
            50% { box-shadow: 0 0 40px rgba(34,197,94,0.3); }
        }
        /* Logs */
        .log-container {
            background: #0d1017;
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px;
            height: 450px;
            overflow-y: auto;
            font-family: 'SF Mono', 'Fira Code', 'Cascadia Code', monospace;
            font-size: 12px;
            line-height: 1.6;
        }
        .log-entry {
            padding: 2px 0;
            color: var(--text-secondary);
            border-bottom: 1px solid rgba(51,55,80,0.2);
        }
        .log-entry:hover { color: var(--text-primary); }
        /* Toast notifications */
        .toast-container {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }
        .toast {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 12px 20px;
            min-width: 280px;
            box-shadow: var(--shadow);
            display: flex;
            align-items: center;
            gap: 10px;
            animation: slideIn 0.3s ease;
            font-size: 14px;
        }
        .toast.success { border-left: 3px solid var(--success); }
        .toast.error { border-left: 3px solid var(--danger); }
        .toast.info { border-left: 3px solid var(--info); }
        .toast .close-toast {
            margin-left: auto;
            cursor: pointer;
            color: var(--text-muted);
            font-size: 16px;
        }
        .toast .close-toast:hover { color: var(--text-primary); }
        @keyframes slideIn { from { transform: translateX(100%); opacity: 0; } to { transform: none; opacity: 1; } }
        @keyframes slideOut { from { transform: none; opacity: 1; } to { transform: translateX(100%); opacity: 0; } }
        /* File list */
        .file-list {
            background: var(--bg-input);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 12px;
            max-height: 200px;
            overflow-y: auto;
            margin-bottom: 16px;
        }
        .file-list label {
            display: block;
            padding: 4px 0;
            font-size: 13px;
            color: var(--text-secondary);
            cursor: pointer;
        }
        .file-list label:hover { color: var(--text-primary); }
        /* Order data row highlights */
        .row-sale { background: rgba(34,197,94,0.08) !important; }
        .row-sale:hover { background: rgba(34,197,94,0.15) !important; }
        .row-special { background: rgba(239,68,68,0.08) !important; }
        .row-special:hover { background: rgba(239,68,68,0.15) !important; }
        .low-stock {
            background: var(--warning);
            color: #000;
            font-weight: 700;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 12px;
        }
        /* Instructions */
        .instructions {
            background: var(--bg-input);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 16px 20px;
            margin-top: 16px;
        }
        .instructions ol {
            padding-left: 20px;
            color: var(--text-secondary);
            font-size: 14px;
            line-height: 2;
        }
        /* Responsive */
        @media (max-width: 768px) {
            .container { padding: 12px; }
            .tabs { flex-wrap: nowrap; overflow-x: auto; }
            .input-row { flex-direction: column; align-items: stretch; }
            .input-row label { min-width: auto; }
            .input-row input { max-width: 100%; }
            .stats-grid { grid-template-columns: 1fr 1fr; }
            .btn-group { flex-direction: column; }
            .app-header { flex-direction: column; gap: 10px; }
            table { font-size: 12px; }
            thead th, tbody td { padding: 8px; }
        }
        /* Light mode */
        [data-theme="light"] {
            --bg-primary: #f5f6fa;
            --bg-secondary: #ffffff;
            --bg-card: #ffffff;
            --bg-input: #f0f1f5;
            --bg-hover: #e8e9f0;
            --border: #d4d6e0;
            --text-primary: #1a1d2e;
            --text-secondary: #5c6078;
            --text-muted: #9498b0;
            --accent: #6366f1;
            --accent-hover: #4f46e5;
            --accent-glow: rgba(99,102,241,0.12);
            --success-bg: rgba(34,197,94,0.1);
            --danger-bg: rgba(239,68,68,0.1);
            --warning-bg: rgba(245,158,11,0.1);
            --shadow: 0 2px 12px rgba(0,0,0,0.08);
        }
        [data-theme="light"] .log-container { background: #f8f9fc; }
        [data-theme="light"] .log-entry { color: #3c3f52; }
        [data-theme="light"] .log-entry:hover { color: #1a1d2e; }
        [data-theme="light"] .toast { background: #fff; }
        /* Theme toggle button */
        .theme-toggle {
            background: var(--bg-input);
            border: 1px solid var(--border);
            color: var(--text-secondary);
            width: 36px;
            height: 36px;
            border-radius: 50%;
            cursor: pointer;
            font-size: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: var(--transition);
        }
        .theme-toggle:hover { background: var(--bg-hover); color: var(--text-primary); }
        /* Scrollbar */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: var(--text-muted); }
        /* Hidden file input */
        .hidden { display: none; }
    </style>
</head>
<body>
    <div class="toast-container" id="toast-container"></div>

    <header class="app-header">
        <div class="app-title">
            <div class="logo">B</div>
            MS DOR Order Bot
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
            <div id="header-status" class="status-pill stopped">
                <span class="status-dot stopped" id="header-dot"></span>
                <span id="header-status-text">Stopped</span>
            </div>
            <button class="theme-toggle" id="theme-toggle" onclick="toggleTheme()" title="Toggle light/dark mode">&#9790;</button>
        </div>
    </header>

    <div class="container">
        <div class="tabs">
            <div class="tab active" onclick="showTab('control', this)">Control</div>
            <div class="tab" onclick="showTab('orders', this)">Orders</div>
            <div class="tab" onclick="showTab('settings', this)">Settings</div>
            <div class="tab" onclick="showTab('logs', this)">Logs</div>
            <div class="tab" onclick="showTab('orderdata', this)">Order Data</div>
            <div class="tab" onclick="showTab('specialorders', this)">Special Orders</div>
        </div>

        <!-- Control Tab -->
        <div id="control" class="tab-content active">
            <div class="stats-grid" id="stats-grid">
                <div class="stat-card">
                    <div class="stat-value" id="stat-total" style="color:var(--accent)">0</div>
                    <div class="stat-label">Total Items</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="stat-completed" style="color:var(--success)">0</div>
                    <div class="stat-label">Completed</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="stat-remaining" style="color:var(--warning)">0</div>
                    <div class="stat-label">Remaining</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="stat-percent" style="color:var(--text-primary)">0%</div>
                    <div class="stat-label">Progress</div>
                </div>
            </div>
            <div class="card">
                <div class="progress-bar">
                    <div class="progress-fill" id="progress-fill" style="width: 0%"></div>
                </div>
            </div>
            <div class="card">
                <div class="control-center">
                    <button class="power-btn" id="power-btn" onclick="toggleBot()" title="Start/Stop Bot">&#9654;</button>
                    <div id="control-status" style="font-size:18px; font-weight:600; color:var(--text-secondary)">Press to Start</div>
                </div>
            </div>
        </div>

        <!-- Orders Tab -->
        <div id="orders" class="tab-content">
            <div class="card">
                <div class="card-title">Add Item</div>
                <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:flex-end;">
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Item #</label>
                        <input type="number" id="item_number" placeholder="12345" style="width:120px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Name</label>
                        <input type="text" id="item_name" placeholder="Product name" style="width:200px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Size</label>
                        <input type="text" id="item_size" placeholder="750ml" style="width:100px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Qty</label>
                        <input type="number" id="quantity" placeholder="10" style="width:80px;">
                    </div>
                    <button class="btn btn-success" onclick="addItem()">+ Add</button>
                </div>
            </div>
            <div class="card">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                    <div class="card-title" style="margin-bottom:0">Order Queue</div>
                    <div class="btn-group" style="margin-bottom:0">
                        <button class="btn btn-ghost btn-sm" onclick="loadOrders()">Refresh</button>
                        <button class="btn btn-ghost btn-sm" onclick="sortOrders()">Sort #</button>
                        <button class="btn btn-ghost btn-sm" onclick="downloadCSV()">Export</button>
                        <button class="btn btn-ghost btn-sm" onclick="document.getElementById('file-upload').click()">Import</button>
                        <button class="btn btn-danger btn-sm" onclick="clearCompleted()">Clear Done</button>
                        <input type="file" id="file-upload" class="hidden" accept=".csv" onchange="uploadCSV(event)">
                    </div>
                </div>
                <div style="overflow-x:auto;">
                    <table id="orders-table">
                        <thead>
                            <tr>
                                <th>Item #</th>
                                <th>Name</th>
                                <th>Size</th>
                                <th>Units</th>
                                <th>Qty</th>
                                <th>Status</th>
                                <th style="width:80px">Action</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Settings Tab -->
        <div id="settings" class="tab-content">
            <div class="card">
                <div class="card-title">Login Credentials</div>
                <div class="input-row">
                    <label>Username</label>
                    <input type="text" id="username" placeholder="Enter username">
                </div>
                <div class="input-row">
                    <label>Password</label>
                    <input type="password" id="password" placeholder="Enter password">
                </div>
                <div class="input-row">
                    <label>Site URL</label>
                    <input type="text" id="url" value="https://tap.dor.ms.gov/">
                </div>
                <label class="checkbox-row">
                    <input type="checkbox" id="headless">
                    Run in background (headless mode)
                </label>
                <button class="btn btn-primary" onclick="saveSettings()" style="margin-top:8px;">Save Settings</button>
            </div>
            <div class="instructions">
                <div class="card-title">Getting Started</div>
                <ol>
                    <li>Enter your DOR login credentials above and save</li>
                    <li>Add items to the order queue in the <strong>Orders</strong> tab</li>
                    <li>Go to <strong>Control</strong> and press the power button to start</li>
                    <li>For first-time use, complete 2FA manually in the browser window</li>
                    <li>The bot runs continuously until you stop it</li>
                </ol>
            </div>
        </div>

        <!-- Logs Tab -->
        <div id="logs" class="tab-content">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                <div class="card-title" style="margin-bottom:0">System Logs</div>
                <div class="btn-group" style="margin-bottom:0">
                    <button class="btn btn-ghost btn-sm" onclick="loadLogs()">Refresh</button>
                    <button class="btn btn-danger btn-sm" onclick="clearLogs()">Clear</button>
                </div>
            </div>
            <div class="log-container" id="logs-content"></div>
        </div>

        <!-- Order Data Tab -->
        <div id="orderdata" class="tab-content">
            <div class="card">
                <div class="card-title">Order Data Lookup</div>
                <p style="color:var(--text-muted); font-size:13px; margin-bottom:12px;">Select order files to search across. Shows all items with source file and order date.</p>
                <div class="btn-group">
                    <button class="btn btn-ghost btn-sm" onclick="loadOrderFiles()">Refresh Files</button>
                    <button class="btn btn-primary btn-sm" onclick="searchOrderData()">Search Selected</button>
                    <label class="checkbox-row" style="margin:0;">
                        <input type="checkbox" id="select-all-files" onchange="toggleAllFiles(this)"> Select All
                    </label>
                </div>
                <div class="file-list" id="order-files-list"></div>
                <div style="display:flex; flex-wrap:wrap; gap:16px; margin-bottom:12px;">
                    <label class="checkbox-row" style="margin:0;">
                        <input type="checkbox" id="include-special-orders" checked> Special Orders
                    </label>
                    <label class="checkbox-row" style="margin:0;">
                        <input type="checkbox" id="include-current-prices" checked> Current Prices
                    </label>
                    <label class="checkbox-row" style="margin:0;">
                        <input type="checkbox" id="include-sales-data" checked> Sales Data
                    </label>
                </div>
                <div id="order-data-status" style="margin-bottom:10px; font-weight:600; color:var(--text-secondary);"></div>
                <input type="text" id="item-search-filter" placeholder="Item #s (comma-separated) or Name/Category..." oninput="filterOrderResults()" onkeydown="if(event.key==='Enter')searchOrderData()" style="width:100%; max-width:500px; margin-bottom:12px;">
            </div>
            <div class="card" style="overflow-x:auto;">
                <table id="order-data-table" style="display:none;">
                    <thead>
                        <tr>
                            <th onclick="sortOrderResults('item_num')">Item #</th>
                            <th onclick="sortOrderResults('size')">Size</th>
                            <th onclick="sortOrderResults('units')">Units</th>
                            <th onclick="sortOrderResults('available')">Avail</th>
                            <th>Qty Req</th>
                            <th onclick="sortOrderResults('units_sold')">Sold</th>
                            <th onclick="sortOrderResults('qty_on_hand')">On Hand</th>
                            <th onclick="sortOrderResults('name')">Name</th>
                            <th onclick="sortOrderResults('source_file')">Source</th>
                            <th onclick="sortOrderResults('spa_date')">Sale Date</th>
                            <th>SPA Price</th>
                            <th onclick="sortOrderResults('case_cost')">Case Cost</th>
                            <th>Discount</th>
                            <th style="width:70px">Action</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>

        <!-- Special Orders Tab -->
        <div id="specialorders" class="tab-content">
            <div class="card">
                <div class="card-title">Add Special Order</div>
                <div style="display:flex; flex-wrap:wrap; gap:8px; align-items:flex-end;">
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Item #</label>
                        <input type="number" id="so-item-number" placeholder="12345" style="width:120px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Qty</label>
                        <input type="number" id="so-quantity" placeholder="10" style="width:80px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Name</label>
                        <input type="text" id="so-name" placeholder="Product name" style="width:200px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Order #</label>
                        <input type="text" id="so-order-number" placeholder="SO-001" style="width:120px;">
                    </div>
                    <div>
                        <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px;">Date</label>
                        <input type="text" id="so-order-date" placeholder="MM/DD/YY" style="width:120px;">
                    </div>
                    <button class="btn btn-success" onclick="addSpecialOrder()">+ Add</button>
                    <button class="btn btn-ghost" onclick="loadSpecialOrders()">Refresh</button>
                </div>
            </div>
            <div class="card" style="overflow-x:auto;">
                <table id="special-orders-table">
                    <thead>
                        <tr>
                            <th>Item #</th>
                            <th>Qty</th>
                            <th>Name</th>
                            <th>Order #</th>
                            <th>Date</th>
                            <th style="width:80px">Action</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        let logsInterval, ordersInterval, logEventSource;

        // ── Toast notifications ──
        function toast(message, type = 'info') {
            const container = document.getElementById('toast-container');
            const t = document.createElement('div');
            t.className = `toast ${type}`;
            t.innerHTML = `<span>${message}</span><span class="close-toast" onclick="this.parentElement.remove()">&times;</span>`;
            container.appendChild(t);
            setTimeout(() => { t.style.animation = 'slideOut 0.3s ease forwards'; setTimeout(() => t.remove(), 300); }, 4000);
        }

        // ── Tabs ──
        function showTab(tabName, el) {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            if (el) el.classList.add('active');
            document.getElementById(tabName).classList.add('active');

            clearInterval(logsInterval);
            clearInterval(ordersInterval);
            if (logEventSource) { logEventSource.close(); logEventSource = null; }

            if (tabName === 'logs') { loadLogs(); startLogStream(); }
            if (tabName === 'orders') { loadOrders(); ordersInterval = setInterval(loadOrders, 5000); }
            if (tabName === 'orderdata') loadOrderFiles();
            if (tabName === 'specialorders') loadSpecialOrders();
        }

        // ── SSE log streaming ──
        function startLogStream() {
            if (logEventSource) logEventSource.close();
            logEventSource = new EventSource('/stream_logs');
            logEventSource.onmessage = (e) => {
                const logsDiv = document.getElementById('logs-content');
                const entry = document.createElement('div');
                entry.className = 'log-entry';
                entry.textContent = e.data;
                logsDiv.appendChild(entry);
                logsDiv.scrollTop = logsDiv.scrollHeight;
                // cap at 200 entries
                while (logsDiv.children.length > 200) logsDiv.removeChild(logsDiv.firstChild);
            };
            logEventSource.onerror = () => {
                logEventSource.close();
                logEventSource = null;
                // fall back to polling
                logsInterval = setInterval(loadLogs, 2000);
            };
        }

        // ── Settings ──
        function saveSettings() {
            fetch('/save_settings', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    username: document.getElementById('username').value,
                    password: document.getElementById('password').value,
                    url: document.getElementById('url').value,
                    headless: document.getElementById('headless').checked
                })
            }).then(() => { toast('Settings saved successfully', 'success'); loadSettings(); });
        }

        function loadSettings() {
            fetch('/get_settings').then(r => r.json()).then(data => {
                document.getElementById('username').value = data.username || '';
                document.getElementById('password').value = data.password || '';
                document.getElementById('url').value = data.url || 'https://tap.dor.ms.gov/';
                document.getElementById('headless').checked = data.headless || false;
            });
        }

        // ── Orders ──
        function addItem() {
            const item_number = document.getElementById('item_number').value;
            const name = document.getElementById('item_name').value.trim();
            const size = document.getElementById('item_size').value.trim();
            const quantity = document.getElementById('quantity').value;
            if (!item_number || !quantity) { toast('Enter at least item number and quantity', 'error'); return; }
            fetch('/add_item', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ item_number, name, size, quantity })
            }).then(() => {
                ['item_number','item_name','item_size','quantity'].forEach(id => document.getElementById(id).value = '');
                loadOrders();
                toast('Item added to queue', 'success');
            });
        }

        function deleteItem(index) {
            fetch('/delete_item', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({index})
            }).then(() => { loadOrders(); toast('Item removed', 'info'); });
        }

        function clearCompleted() {
            fetch('/clear_completed', {method: 'POST'}).then(() => { loadOrders(); toast('Completed items cleared', 'success'); });
        }

        function loadOrders() {
            fetch('/get_orders').then(r => r.json()).then(data => {
                const tbody = document.querySelector('#orders-table tbody');
                tbody.innerHTML = data.map((item, i) => `
                    <tr>
                        <td style="font-weight:600">${item.item_number}</td>
                        <td>${item.name || '<span style="color:var(--text-muted)">-</span>'}</td>
                        <td>${item.size || '-'}</td>
                        <td>${item.units || '-'}</td>
                        <td>${item.quantity}</td>
                        <td>${item.order_filled === 'yes'
                            ? '<span class="badge badge-success"><span class="status-dot" style="background:var(--success)"></span> Done</span>'
                            : '<span class="badge badge-warning"><span class="status-dot" style="background:var(--warning)"></span> Pending</span>'}</td>
                        <td><button class="btn btn-danger btn-sm" onclick="deleteItem(${i})">Del</button></td>
                    </tr>
                `).join('');
                loadStats();
            });
        }

        function sortOrders() {
            fetch('/sort_orders', {method: 'POST'}).then(() => { loadOrders(); toast('Sorted by item number', 'info'); });
        }

        function downloadCSV() { window.location.href = '/download_csv'; }

        function uploadCSV(event) {
            const file = event.target.files[0];
            const fd = new FormData();
            fd.append('file', file);
            fetch('/upload_csv', { method: 'POST', body: fd }).then(() => { toast('CSV uploaded', 'success'); loadOrders(); });
            event.target.value = '';
        }

        // ── Stats ──
        function loadStats() {
            fetch('/get_stats').then(r => r.json()).then(data => {
                document.getElementById('stat-total').textContent = data.total;
                document.getElementById('stat-completed').textContent = data.completed;
                document.getElementById('stat-remaining').textContent = data.remaining;
                const pct = data.total > 0 ? Math.round((data.completed / data.total) * 100) : 0;
                document.getElementById('stat-percent').textContent = pct + '%';
                document.getElementById('progress-fill').style.width = pct + '%';
            });
        }

        // ── Bot control ──
        function toggleBot() {
            fetch('/toggle_bot', {method: 'POST'}).then(r => r.json()).then(data => {
                updateStatus(data);
                toast(data.running ? 'Bot started' : 'Bot stopped', data.running ? 'success' : 'info');
            });
        }

        function updateStatus(data) {
            const powerBtn = document.getElementById('power-btn');
            const controlStatus = document.getElementById('control-status');
            const headerPill = document.getElementById('header-status');
            const headerDot = document.getElementById('header-dot');
            const headerText = document.getElementById('header-status-text');

            if (data.running) {
                powerBtn.className = 'power-btn running';
                powerBtn.innerHTML = '&#9724;';
                controlStatus.textContent = 'Bot is running...';
                controlStatus.style.color = 'var(--success)';
                headerPill.className = 'status-pill running';
                headerDot.className = 'status-dot running';
                headerText.textContent = 'Running';
            } else {
                powerBtn.className = 'power-btn';
                powerBtn.innerHTML = '&#9654;';
                controlStatus.textContent = 'Press to Start';
                controlStatus.style.color = 'var(--text-secondary)';
                headerPill.className = 'status-pill stopped';
                headerDot.className = 'status-dot stopped';
                headerText.textContent = 'Stopped';
            }
            loadStats();
        }

        // ── Logs ──
        function loadLogs() {
            fetch('/get_logs').then(r => r.json()).then(data => {
                const logsDiv = document.getElementById('logs-content');
                logsDiv.innerHTML = data.logs.map(log => `<div class="log-entry">${log}</div>`).join('');
                logsDiv.scrollTop = logsDiv.scrollHeight;
            });
        }

        function clearLogs() {
            fetch('/clear_logs', {method: 'POST'}).then(() => { loadLogs(); toast('Logs cleared', 'info'); });
        }

        // ── Order Data ──
        let orderDataResults = [];
        let orderDataSortKey = 'order_date';
        let orderDataSortAsc = false;

        function loadOrderFiles() {
            fetch('/get_order_files').then(r => r.json()).then(data => {
                const container = document.getElementById('order-files-list');
                if (data.files.length === 0) {
                    container.innerHTML = '<p style="color:var(--text-muted);">No .ods files found in Order Data folder.</p>';
                    return;
                }
                container.innerHTML = data.files.map(f => `
                    <label><input type="checkbox" class="order-file-cb" value="${f.filename}" checked>
                    ${f.display_name} <span style="color:var(--text-muted); font-size:11px;">(${f.filename})</span></label>
                `).join('');
                document.getElementById('select-all-files').checked = true;
            });
        }

        function toggleAllFiles(cb) {
            document.querySelectorAll('.order-file-cb').forEach(c => c.checked = cb.checked);
        }

        function searchOrderData() {
            const selected = Array.from(document.querySelectorAll('.order-file-cb:checked')).map(c => c.value);
            const includeSpecial = document.getElementById('include-special-orders').checked;
            const includeCurrentPrices = document.getElementById('include-current-prices').checked;
            const includeSalesData = document.getElementById('include-sales-data').checked;
            if (selected.length === 0 && !includeSpecial) { toast('Select at least one file', 'error'); return; }
            document.getElementById('order-data-status').textContent = 'Searching...';
            fetch('/search_order_data', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({files: selected, include_special: includeSpecial, include_current_prices: includeCurrentPrices, include_sales_data: includeSalesData})
            }).then(r => r.json()).then(data => {
                orderDataResults = data.items;
                const unique = new Set(data.items.map(i => i.item_num)).size;
                let src = selected.length + ' file(s)';
                if (includeSpecial) src += ' + special orders';
                document.getElementById('order-data-status').textContent = `Found ${data.items.length} rows (${unique} unique items) across ${src}.`;
                renderOrderResults();
            }).catch(err => { document.getElementById('order-data-status').textContent = 'Error: ' + err; });
        }

        function renderOrderResults() {
            const table = document.getElementById('order-data-table');
            const raw = document.getElementById('item-search-filter').value.trim();
            let items = orderDataResults;
            if (raw) {
                const parts = raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                const numericParts = parts.filter(p => /^[0-9]+$/.test(p));
                if (numericParts.length === parts.length && numericParts.length > 0) {
                    const numSet = new Set(numericParts);
                    items = items.filter(i => numSet.has(i.item_num));
                } else {
                    items = items.filter(i =>
                        parts.some(p => i.item_num.toLowerCase().includes(p) || i.name.toLowerCase().includes(p) || (i.category||'').toLowerCase().includes(p))
                    );
                }
            }
            const tbody = table.querySelector('tbody');
            tbody.innerHTML = items.map(item => {
                const isSO = item.source_file.startsWith('SO#');
                const hasSale = item.spa_date && item.spa_date !== '';
                let rowClass = '';
                if (hasSale) rowClass = 'row-sale';
                else if (isSO) rowClass = 'row-special';
                const sold = parseFloat(item.units_sold) || 0;
                const onHand = parseFloat(item.qty_on_hand) || 0;
                const lowStock = sold > 0 && onHand > 0 && (sold * 3 / 4) > onHand;
                return `
                <tr class="${rowClass}">
                    <td style="font-weight:600">${item.item_num}</td>
                    <td>${item.size || ''}</td>
                    <td>${item.units || ''}</td>
                    <td>${item.available || ''}</td>
                    <td>${item.qty_requested || ''}</td>
                    <td>${item.units_sold || ''}</td>
                    <td>${lowStock ? '<span class="low-stock">' + (item.qty_on_hand||'') + '</span>' : (item.qty_on_hand || '')}</td>
                    <td>${item.name}</td>
                    <td><span style="font-size:12px; color:var(--text-muted)">${item.source_file}</span></td>
                    <td>${item.spa_date || ''}</td>
                    <td>${item.spa_price || ''}</td>
                    <td>${item.case_cost || ''}</td>
                    <td>${item.spa_discount || ''}</td>
                    <td><button class="btn btn-primary btn-sm" onclick="addToBot('${item.item_num}', '${(item.name||'').replace(/'/g,"\\'")}', '${(item.size||'').replace(/'/g,"\\'")}', '${item.units||''}')">+ Bot</button></td>
                </tr>`;
            }).join('');
            table.style.display = items.length > 0 ? 'table' : 'none';
        }

        function filterOrderResults() { renderOrderResults(); }

        function sortOrderResults(key) {
            if (orderDataSortKey === key) orderDataSortAsc = !orderDataSortAsc;
            else { orderDataSortKey = key; orderDataSortAsc = true; }
            orderDataResults.sort((a, b) => {
                const va = (key === 'spa_date' ? a.spa_sort_date : a[key]) || '';
                const vb = (key === 'spa_date' ? b.spa_sort_date : b[key]) || '';
                if (va < vb) return orderDataSortAsc ? -1 : 1;
                if (va > vb) return orderDataSortAsc ? 1 : -1;
                return 0;
            });
            renderOrderResults();
        }

        function addToBot(itemNum, name, size, units) {
            const qty = prompt('Enter quantity for item ' + itemNum + ':');
            if (qty === null || qty.trim() === '') return;
            if (isNaN(qty) || parseInt(qty) <= 0) { toast('Enter a valid quantity', 'error'); return; }
            fetch('/add_item', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({item_number: itemNum, name: name||'', size: size||'', units: units||'', quantity: qty.trim()})
            }).then(() => toast(`Item ${itemNum} (qty: ${qty.trim()}) added to queue`, 'success'));
        }

        // ── Special Orders ──
        function loadSpecialOrders() {
            fetch('/get_special_orders').then(r => r.json()).then(data => {
                document.querySelector('#special-orders-table tbody').innerHTML = data.map((item, i) => `
                    <tr>
                        <td style="font-weight:600">${item.item_number}</td>
                        <td>${item.quantity}</td>
                        <td>${item.name}</td>
                        <td>${item.order_number}</td>
                        <td>${item.order_date}</td>
                        <td><button class="btn btn-danger btn-sm" onclick="deleteSpecialOrder(${i})">Del</button></td>
                    </tr>
                `).join('');
            });
        }

        function addSpecialOrder() {
            const item_number = document.getElementById('so-item-number').value;
            if (!item_number) { toast('Enter at least an item number', 'error'); return; }
            fetch('/add_special_order', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    item_number,
                    quantity: document.getElementById('so-quantity').value,
                    name: document.getElementById('so-name').value,
                    order_number: document.getElementById('so-order-number').value,
                    order_date: document.getElementById('so-order-date').value
                })
            }).then(() => {
                ['so-item-number','so-quantity','so-name','so-order-number','so-order-date'].forEach(id => document.getElementById(id).value = '');
                loadSpecialOrders();
                toast('Special order added', 'success');
            });
        }

        function deleteSpecialOrder(index) {
            fetch('/delete_special_order', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({index})
            }).then(() => { loadSpecialOrders(); toast('Special order removed', 'info'); });
        }

        // ── Periodic status check ──
        setInterval(() => {
            fetch('/get_status').then(r => r.json()).then(data => updateStatus(data));
        }, 3000);

        // ── Theme toggle ──
        function toggleTheme() {
            const html = document.documentElement;
            const current = html.getAttribute('data-theme') || 'dark';
            const next = current === 'dark' ? 'light' : 'dark';
            html.setAttribute('data-theme', next);
            localStorage.setItem('theme', next);
            document.getElementById('theme-toggle').innerHTML = next === 'dark' ? '&#9790;' : '&#9728;';
        }
        (function initTheme() {
            const saved = localStorage.getItem('theme') || 'dark';
            document.documentElement.setAttribute('data-theme', saved);
            document.getElementById('theme-toggle').innerHTML = saved === 'dark' ? '&#9790;' : '&#9728;';
        })();

        // ── Init ──
        loadSettings();
        loadOrders();
    </script>
</body>
</html>
'''

# Store logs in memory
logs = []

class LogCapture(logging.Handler):
    def emit(self, record):
        global logs
        log_entry = self.format(record)
        logs.append(f"[{record.asctime}] {log_entry}")
        if len(logs) > 100:  # Keep only last 100 logs
            logs.pop(0)

# Set up logging
log_handler = LogCapture()
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S'))
logging.getLogger().addHandler(log_handler)

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/save_settings', methods=['POST'])
def save_settings():
    data = request.json
    with open('.env', 'w') as f:
        f.write(f"SITE_USERNAME={data['username']}\n")
        f.write(f"SITE_PASSWORD={data['password']}\n")
        f.write(f"SITE_URL={data['url']}\n")
        f.write(f"HEADLESS={data.get('headless', False)}\n")
    load_dotenv(override=True)
    return jsonify({'success': True})

@app.route('/get_settings')
def get_settings():
    load_dotenv()
    return jsonify({
        'username': os.getenv('SITE_USERNAME', ''),
        'password': os.getenv('SITE_PASSWORD', ''),
        'url': os.getenv('SITE_URL', 'https://tap.dor.ms.gov/'),
        'headless': os.getenv('HEADLESS', 'False').lower() == 'true'
    })

@app.route('/get_orders')
def get_orders():
    return jsonify(_read_csv_locked())

@app.route('/add_item', methods=['POST'])
def add_item():
    data = request.json
    orders = _read_csv_locked()
    orders.append({
        'item_number': data['item_number'],
        'quantity': data['quantity'],
        'name': data.get('name', ''),
        'size': data.get('size', ''),
        'units': data.get('units', ''),
        'order_filled': ''
    })
    _write_csv_locked(orders)
    return jsonify({'success': True})

@app.route('/delete_item', methods=['POST'])
def delete_item():
    data = request.json
    index = data['index']
    orders = _read_csv_locked()
    if 0 <= index < len(orders):
        orders.pop(index)
    _write_csv_locked(orders)
    return jsonify({'success': True})

@app.route('/clear_completed', methods=['POST'])
def clear_completed():
    orders = _read_csv_locked()
    orders = [row for row in orders if row.get('order_filled', '').lower() != 'yes']
    _write_csv_locked(orders)
    return jsonify({'success': True})

@app.route('/sort_orders', methods=['POST'])
def sort_orders():
    orders = _read_csv_locked()
    orders.sort(key=lambda x: int(x.get('item_number', 0) or 0))
    _write_csv_locked(orders)
    return jsonify({'success': True})

@app.route('/get_stats')
def get_stats():
    orders = _read_csv_locked()
    total = len(orders)
    completed = sum(1 for o in orders if o.get('order_filled', '').lower() == 'yes')
    return jsonify({
        'total': total,
        'completed': completed,
        'remaining': total - completed
    })

@app.route('/toggle_bot', methods=['POST'])
def toggle_bot():
    global bot_running, bot_thread
    
    if not bot_running:
        # Validate settings first
        load_dotenv()
        if not os.getenv('SITE_USERNAME') or not os.getenv('SITE_PASSWORD'):
            return jsonify({'error': 'Please configure settings first'}), 400
        
        bot_running = True
        stop_event.clear()
        bot_thread = threading.Thread(target=run_bot_thread, daemon=True)
        bot_thread.start()
        logging.info("Bot started")
    else:
        bot_running = False
        stop_event.set()
        logging.info("Bot stopped")
    
    return jsonify({'running': bot_running})

@app.route('/get_status')
def get_status():
    return jsonify({'running': bot_running})

@app.route('/get_logs')
def get_logs():
    global logs
    return jsonify({'logs': logs[-50:]})

@app.route('/stream_logs')
def stream_logs():
    """SSE endpoint for real-time log streaming."""
    from flask import Response
    import time as _time

    def generate():
        last_count = len(logs)
        while True:
            current_count = len(logs)
            if current_count > last_count:
                for entry in logs[last_count:current_count]:
                    yield f"data: {entry}\n\n"
                last_count = current_count
            _time.sleep(0.5)

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})

@app.route('/clear_logs', methods=['POST'])
def clear_logs():
    global logs
    logs = []
    return jsonify({'success': True})

@app.route('/download_csv')
def download_csv():
    if Path('orders.csv').exists():
        return send_file('orders.csv', as_attachment=True)
    return "No CSV file found", 404

@app.route('/upload_csv', methods=['POST'])
def upload_csv():
    file = request.files['file']
    if file and file.filename.endswith('.csv'):
        file.save('orders.csv')
        return jsonify({'success': True})
    return jsonify({'error': 'Invalid file'}), 400

ORDER_DATA_DIR = Path('Order Data')

def parse_date_from_filename(filename):
    """Extract date from filenames like 'R301542-2-16-26.ods' (M-DD-YY)."""
    match = re.search(r'(\d{1,2})-(\d{1,2})-(\d{2,4})', filename)
    if match:
        try:
            m, d, y = match.group(1), match.group(2), match.group(3)
            if len(y) == 2:
                y = '20' + y
            return datetime(int(y), int(m), int(d))
        except (ValueError, TypeError):
            pass
    return None

def read_ods_file(filepath):
    """Read an ODS file. Returns (date, list_of_item_dicts)."""
    doc = load_ods(str(filepath))
    dt = parse_date_from_filename(filepath.name)
    items = []
    for sheet in doc.body.getElementsByType(Table):
        rows = sheet.getElementsByType(TableRow)
        if not rows:
            continue
        for row in rows[1:]:
            cells = row.getElementsByType(TableCell)
            vals = []
            for cell in cells:
                repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
                ps = cell.getElementsByType(P)
                text = ''.join(p.firstChild.data if p.firstChild else '' for p in ps)
                vals.extend([text] * min(repeat, 20))
            if len(vals) >= 9 and vals[0].strip():
                items.append({
                    'item_num': vals[0].strip(),
                    'name': vals[1].strip(),
                    'category': vals[2].strip(),
                    'pkg_type': vals[3].strip(),
                    'units': vals[4].strip(),
                    'price': vals[5].strip(),
                    'qty_requested': vals[6].strip(),
                    'qty_reserved': vals[7].strip(),
                    'sub_total': vals[8].strip(),
                })
    return dt, items

@app.route('/get_order_files')
def get_order_files():
    if not ORDER_DATA_DIR.exists():
        return jsonify({'files': []})
    files = []
    for f in sorted(ORDER_DATA_DIR.iterdir()):
        if f.suffix.lower() == '.ods':
            dt = parse_date_from_filename(f.name)
            display = dt.strftime('%b %d, %Y') if dt else f.stem
            files.append({
                'filename': f.name,
                'display_name': display,
                'sort_key': dt.isoformat() if dt else '',
            })
    files.sort(key=lambda x: x['sort_key'], reverse=True)
    return jsonify({'files': files})

SPECIAL_ORDER_CSV = Path('specialorder.csv')
SPECIAL_ORDER_FIELDS = ['item_number', 'quantity', 'name', 'order_number', 'order_date']

def _read_special_orders():
    orders = []
    if SPECIAL_ORDER_CSV.exists():
        with open(SPECIAL_ORDER_CSV, 'r', newline='') as f:
            reader = csv.DictReader(f)
            orders = list(reader)
    return orders

def _write_special_orders(orders):
    with open(SPECIAL_ORDER_CSV, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=SPECIAL_ORDER_FIELDS)
        writer.writeheader()
        writer.writerows(orders)

@app.route('/get_special_orders')
def get_special_orders():
    return jsonify(_read_special_orders())

@app.route('/add_special_order', methods=['POST'])
def add_special_order():
    data = request.json
    orders = _read_special_orders()
    orders.append({
        'item_number': data.get('item_number', ''),
        'quantity': data.get('quantity', ''),
        'name': data.get('name', ''),
        'order_number': data.get('order_number', ''),
        'order_date': data.get('order_date', ''),
    })
    _write_special_orders(orders)
    return jsonify({'success': True})

@app.route('/delete_special_order', methods=['POST'])
def delete_special_order():
    index = request.json['index']
    orders = _read_special_orders()
    if 0 <= index < len(orders):
        orders.pop(index)
    _write_special_orders(orders)
    return jsonify({'success': True})

FUTURE_SPA_DIR = ORDER_DATA_DIR / 'FutureSPA'
CURRENT_PRICES_DIR = ORDER_DATA_DIR / 'CurrentPrices'
SALES_DATA_DIR = ORDER_DATA_DIR / 'SalesData'

def _load_current_prices():
    """Load CurrentPrices data into a dict keyed by item code."""
    prices = {}
    if not CURRENT_PRICES_DIR.exists():
        return prices
    for f in CURRENT_PRICES_DIR.iterdir():
        if f.suffix.lower() != '.ods':
            continue
        try:
            doc = load_ods(str(f))
        except Exception:
            continue
        for sheet in doc.body.getElementsByType(Table):
            rows = sheet.getElementsByType(TableRow)
            if not rows:
                continue
            for row in rows[1:]:
                cells = row.getElementsByType(TableCell)
                vals = []
                for cell in cells:
                    repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
                    ps = cell.getElementsByType(P)
                    text = ''.join(p.firstChild.data if p.firstChild else '' for p in ps)
                    vals.extend([text] * min(repeat, 20))
                if len(vals) >= 4 and vals[0].strip():
                    item_code = vals[0].strip()
                    prices[item_code] = {
                        'name': vals[1].strip() if len(vals) > 1 else '',
                        'size': vals[2].strip() if len(vals) > 2 else '',
                        'available': vals[3].strip() if len(vals) > 3 else '',
                        'units': vals[5].strip() if len(vals) > 5 else '',
                        'case_cost': vals[6].strip() if len(vals) > 6 else '',
                    }
    return prices

def _load_sales_data():
    """Load SalesData from .xls files (HTML tables saved as .xls) keyed by item number."""
    from html.parser import HTMLParser

    class TableParser(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows = []
            self.current_row = None
            self.current_cell = ''
            self.in_cell = False
        def handle_starttag(self, tag, attrs):
            if tag == 'tr':
                self.current_row = []
            elif tag in ('td', 'th'):
                self.in_cell = True
                self.current_cell = ''
        def handle_endtag(self, tag):
            if tag in ('td', 'th') and self.in_cell:
                self.in_cell = False
                if self.current_row is not None:
                    self.current_row.append(self.current_cell.strip())
            elif tag == 'tr' and self.current_row is not None:
                self.rows.append(self.current_row)
                self.current_row = None
        def handle_data(self, data):
            if self.in_cell:
                self.current_cell += data

    sales = {}
    if not SALES_DATA_DIR.exists():
        return sales
    for f in SALES_DATA_DIR.iterdir():
        if f.suffix.lower() not in ('.xls', '.html', '.htm'):
            continue
        try:
            with open(f, 'r', errors='ignore') as fh:
                content = fh.read()
        except Exception:
            continue
        parser = TableParser()
        parser.feed(content)
        header_idx = None
        for i, row in enumerate(parser.rows):
            if any('Item #' in c or 'Item#' in c for c in row):
                header_idx = i
                break
        if header_idx is None:
            continue
        for row in parser.rows[header_idx + 1:]:
            if len(row) < 10:
                continue
            item_num = row[1].strip()
            if not item_num or not any(c.isdigit() for c in item_num):
                continue
            name = row[3].strip() if len(row) > 3 else ''
            units_sold = row[4].strip() if len(row) > 4 else ''
            qty_oh = row[9].strip() if len(row) > 9 else ''
            sales[item_num] = {
                'name': name,
                'units_sold': units_sold,
                'qty_on_hand': qty_oh,
            }
    return sales

def _load_future_spa():
    """Load FutureSPA data into a dict keyed by item code."""
    spa = {}
    if not FUTURE_SPA_DIR.exists():
        return spa
    for f in FUTURE_SPA_DIR.iterdir():
        if f.suffix.lower() != '.ods':
            continue
        try:
            doc = load_ods(str(f))
        except Exception:
            continue
        for sheet in doc.body.getElementsByType(Table):
            rows = sheet.getElementsByType(TableRow)
            if not rows:
                continue
            for row in rows[1:]:
                cells = row.getElementsByType(TableCell)
                vals = []
                for cell in cells:
                    repeat = int(cell.getAttribute('numbercolumnsrepeated') or 1)
                    ps = cell.getElementsByType(P)
                    text = ''.join(p.firstChild.data if p.firstChild else '' for p in ps)
                    vals.extend([text] * min(repeat, 20))
                if len(vals) >= 9 and vals[1].strip():
                    item_code = vals[1].strip()
                    spa[item_code] = {
                        'name': vals[2].strip(),
                        'spa_date': vals[5].strip(),
                        'spa_price': vals[7].strip(),
                        'spa_discount': vals[8].strip(),
                    }
    return spa

@app.route('/search_order_data', methods=['POST'])
def search_order_data():
    selected = request.json.get('files', [])
    include_special = request.json.get('include_special', False)
    include_current_prices = request.json.get('include_current_prices', False)
    include_sales_data = request.json.get('include_sales_data', False)

    spa_lookup = _load_future_spa()
    prices_lookup = _load_current_prices() if include_current_prices else {}
    sales_lookup = _load_sales_data() if include_sales_data else {}

    all_rows = []
    for fname in selected:
        fpath = ORDER_DATA_DIR / fname
        if not fpath.exists():
            continue
        try:
            dt, rows = read_ods_file(fpath)
        except Exception:
            continue
        date_str = dt.strftime('%b %d, %Y') if dt else fname
        sort_date = dt.isoformat() if dt else ''
        for row in rows:
            row['sort_date'] = sort_date
            row['source_file'] = fname
            spa = spa_lookup.get(row['item_num'], {})
            row['spa_date'] = spa.get('spa_date', '')
            row['spa_price'] = spa.get('spa_price', '')
            row['spa_discount'] = spa.get('spa_discount', '')
            row['spa_sort_date'] = spa.get('spa_date', '')
            cp = prices_lookup.get(row['item_num'], {})
            row['size'] = cp.get('size', '')
            row['units'] = cp.get('units', '')
            row['available'] = cp.get('available', '')
            row['case_cost'] = cp.get('case_cost', '')
            sd = sales_lookup.get(row['item_num'], {})
            row['units_sold'] = sd.get('units_sold', '')
            row['qty_on_hand'] = sd.get('qty_on_hand', '')
            all_rows.append(row)

    if include_special and SPECIAL_ORDER_CSV.exists():
        for srow in _read_special_orders():
            item_num = srow.get('item_number', '').strip()
            if not item_num:
                continue
            spa = spa_lookup.get(item_num, {})
            cp = prices_lookup.get(item_num, {})
            sd = sales_lookup.get(item_num, {})
            all_rows.append({
                'item_num': item_num,
                'name': srow.get('name', '').strip(),
                'qty_requested': '',
                'sort_date': '',
                'source_file': 'SO#' + srow.get('order_number', '').strip(),
                'spa_date': spa.get('spa_date', ''),
                'spa_price': spa.get('spa_price', ''),
                'spa_discount': spa.get('spa_discount', ''),
                'spa_sort_date': spa.get('spa_date', ''),
                'size': cp.get('size', ''),
                'units': cp.get('units', ''),
                'available': cp.get('available', ''),
                'case_cost': cp.get('case_cost', ''),
                'units_sold': sd.get('units_sold', ''),
                'qty_on_hand': sd.get('qty_on_hand', ''),
            })

    seen_items = set(row['item_num'] for row in all_rows)
    for item_code, spa in spa_lookup.items():
        if item_code not in seen_items:
            cp = prices_lookup.get(item_code, {})
            sd = sales_lookup.get(item_code, {})
            all_rows.append({
                'item_num': item_code,
                'name': spa.get('name', ''),
                'qty_requested': '',
                'sort_date': '',
                'source_file': 'SPA Only',
                'spa_date': spa.get('spa_date', ''),
                'spa_price': spa.get('spa_price', ''),
                'spa_discount': spa.get('spa_discount', ''),
                'spa_sort_date': spa.get('spa_date', ''),
                'size': cp.get('size', ''),
                'units': cp.get('units', ''),
                'available': cp.get('available', ''),
                'case_cost': cp.get('case_cost', ''),
                'units_sold': sd.get('units_sold', ''),
                'qty_on_hand': sd.get('qty_on_hand', ''),
            })
            seen_items.add(item_code)

    all_supplementary = set()
    if include_current_prices:
        all_supplementary.update(prices_lookup.keys())
    if include_sales_data:
        all_supplementary.update(sales_lookup.keys())
    for item_code in all_supplementary:
        if item_code not in seen_items:
            cp = prices_lookup.get(item_code, {})
            sd = sales_lookup.get(item_code, {})
            spa = spa_lookup.get(item_code, {})
            name = cp.get('name', '') or sd.get('name', '') or spa.get('name', '')
            all_rows.append({
                'item_num': item_code,
                'name': name,
                'qty_requested': '',
                'sort_date': '',
                'source_file': '',
                'spa_date': spa.get('spa_date', ''),
                'spa_price': spa.get('spa_price', ''),
                'spa_discount': spa.get('spa_discount', ''),
                'spa_sort_date': spa.get('spa_date', ''),
                'size': cp.get('size', ''),
                'units': cp.get('units', ''),
                'available': cp.get('available', ''),
                'case_cost': cp.get('case_cost', ''),
                'units_sold': sd.get('units_sold', ''),
                'qty_on_hand': sd.get('qty_on_hand', ''),
            })
            seen_items.add(item_code)

    all_rows.sort(key=lambda x: x.get('sort_date', ''), reverse=True)
    return jsonify({'items': all_rows})

def run_bot_thread():
    """Run the bot in a separate thread, reusing bot_script.py logic."""
    global bot_running
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    try:
        from bot_script import WebAutomationBot, read_csv_file, update_csv_file

        async def run_bot():
            global bot_running
            load_dotenv()
            headless = os.getenv('HEADLESS', 'False').lower() == 'true'
            bot = WebAutomationBot(headless=headless)

            if not Path('orders.csv').exists():
                _write_csv_locked([])
                logging.info("Created empty orders.csv")

            try:
                await bot.setup(use_saved_auth=True)
                logging.info("Bot ready. If prompted for OTP, complete it in the browser.")

                consecutive_errors = 0
                on_item_entry = False
                while not stop_event.is_set():
                    try:
                        items = read_csv_file('orders.csv')
                        unfilled = [i for i in items if i.get('order_filled', '').lower() != 'yes']

                        if not unfilled:
                            logging.info("All items completed! Checking for new items in 5 seconds...")
                            await asyncio.sleep(5)
                            on_item_entry = False
                            continue

                        if not on_item_entry or 'itemEntry' not in bot.page.url:
                            await bot.start_order()
                            on_item_entry = True

                        logging.info(f"Checking {len(unfilled)} items for availability...")
                        items_found, total_qty_added = await bot.check_and_process_items(items)
                        consecutive_errors = 0

                        if items_found and total_qty_added >= 10:
                            item_numbers = [str(i['item_number']) for i in items_found]
                            logging.info(f"Found {len(items_found)} items, {total_qty_added} total qty: {', '.join(item_numbers)}")
                            await bot.submit_order()
                            update_csv_file('orders.csv', items)
                            on_item_entry = False
                            if not [i for i in items if i.get('order_filled', '').lower() != 'yes']:
                                logging.info("All items filled!")
                        elif items_found and total_qty_added < 10:
                            logging.warning(f"Need min 10 qty total (have {total_qty_added}). Reverting.")
                            for item in items_found:
                                item['order_filled'] = ''
                        else:
                            logging.info("No items available. Re-checking...")
                            await asyncio.sleep(1)

                    except Exception as e:
                        consecutive_errors += 1
                        logging.error(f"Error (attempt {consecutive_errors}): {e}")

                        if consecutive_errors < 3:
                            try:
                                await bot.start_order()
                                on_item_entry = True
                                logging.info("Recovered, resuming...")
                                continue
                            except Exception:
                                pass

                        logging.info("Re-initializing bot (full login)...")
                        consecutive_errors = 0
                        on_item_entry = False
                        try:
                            await bot.cleanup()
                        except Exception:
                            pass
                        bot = WebAutomationBot(headless=headless)
                        try:
                            await bot.setup(use_saved_auth=True)
                            logging.info("Bot recovered successfully")
                        except Exception as setup_err:
                            logging.error(f"Recovery failed: {setup_err}")
                            await asyncio.sleep(5)

            except Exception as e:
                logging.error(f"Bot startup error: {e}")
            finally:
                try:
                    await bot.cleanup()
                except Exception:
                    pass

        loop.run_until_complete(run_bot())
    except Exception as e:
        logging.error(f"Thread error: {e}")
    finally:
        loop.close()
        bot_running = False

if __name__ == '__main__':
    print("\n" + "="*50)
    print("Mississippi DOR Order Bot - Web Interface")
    print("="*50)
    print("\nStarting web server...")
    print("\n🌐 Open your browser and go to: http://localhost:5050")
    print("\nPress Ctrl+C to stop the server\n")
    
    import webbrowser
    webbrowser.open('http://localhost:5050')
    
    app.run(debug=False, port=5050, host='0.0.0.0')