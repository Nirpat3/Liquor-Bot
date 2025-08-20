# Save as web_gui.py
from flask import Flask, render_template, request, jsonify, send_file
import threading
import asyncio
import csv
import json
from pathlib import Path
import logging
import os
from dotenv import load_dotenv

app = Flask(__name__)
bot_thread = None
bot_running = False
stop_event = threading.Event()

# HTML template
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Mississippi DOR Order Bot</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .container { max-width: 1200px; margin: 0 auto; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .tab { padding: 10px 20px; background: #f0f0f0; cursor: pointer; border-radius: 5px; }
        .tab.active { background: #007bff; color: white; }
        .tab-content { display: none; }
        .tab-content.active { display: block; }
        input, button { padding: 8px; margin: 5px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .status { font-size: 18px; font-weight: bold; margin: 20px 0; }
        .running { color: green; }
        .stopped { color: red; }
        #logs { background: #f5f5f5; padding: 10px; height: 300px; overflow-y: scroll; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Mississippi DOR Order Bot</h1>
        
        <div class="tabs">
            <div class="tab active" onclick="showTab('settings')">Settings</div>
            <div class="tab" onclick="showTab('orders')">Orders</div>
            <div class="tab" onclick="showTab('control')">Control</div>
            <div class="tab" onclick="showTab('logs')">Logs</div>
        </div>
        
        <div id="settings" class="tab-content active">
            <h2>Login Settings</h2>
            <div>
                <label>Username: <input type="text" id="username" placeholder="Enter username"></label><br>
                <label>Password: <input type="password" id="password" placeholder="Enter password"></label><br>
                <label>URL: <input type="text" id="url" value="https://tap.dor.ms.gov/"></label><br>
                <button onclick="saveSettings()">Save Settings</button>
            </div>
        </div>
        
        <div id="orders" class="tab-content">
            <h2>Order Items</h2>
            <div>
                <input type="number" id="item_number" placeholder="Item Number">
                <input type="number" id="quantity" placeholder="Quantity">
                <button onclick="addItem()">Add Item</button>
                <button onclick="clearCompleted()">Clear Completed</button>
            </div>
            <table id="orders-table">
                <thead>
                    <tr>
                        <th>Item Number</th>
                        <th>Quantity</th>
                        <th>Status</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
        
        <div id="control" class="tab-content">
            <h2>Bot Control</h2>
            <button id="bot-toggle" onclick="toggleBot()">Start Bot</button>
            <div id="status" class="status stopped">Status: Stopped</div>
            <div id="stats">Items Processed: 0 | Items Remaining: 0</div>
        </div>
        
        <div id="logs" class="tab-content">
            <h2>Logs</h2>
            <div id="logs"></div>
            <button onclick="clearLogs()">Clear Logs</button>
        </div>
    </div>
    
    <script>
        function showTab(tabName) {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            event.target.classList.add('active');
            document.getElementById(tabName).classList.add('active');
        }
        
        function saveSettings() {
            fetch('/save_settings', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    username: document.getElementById('username').value,
                    password: document.getElementById('password').value,
                    url: document.getElementById('url').value
                })
            }).then(() => alert('Settings saved!'));
        }
        
        function addItem() {
            fetch('/add_item', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    item_number: document.getElementById('item_number').value,
                    quantity: document.getElementById('quantity').value
                })
            }).then(() => loadOrders());
        }
        
        function toggleBot() {
            fetch('/toggle_bot', {method: 'POST'})
                .then(r => r.json())
                .then(data => updateStatus(data));
        }
        
        function updateStatus(data) {
            document.getElementById('status').className = 'status ' + (data.running ? 'running' : 'stopped');
            document.getElementById('status').textContent = 'Status: ' + (data.running ? 'Running' : 'Stopped');
            document.getElementById('bot-toggle').textContent = data.running ? 'Stop Bot' : 'Start Bot';
        }
        
        function loadOrders() {
            fetch('/get_orders')
                .then(r => r.json())
                .then(data => {
                    const tbody = document.querySelector('#orders-table tbody');
                    tbody.innerHTML = data.map(item => `
                        <tr>
                            <td>${item.item_number}</td>
                            <td>${item.quantity}</td>
                            <td>${item.order_filled || 'Pending'}</td>
                            <td><button onclick="deleteItem(${item.item_number})">Delete</button></td>
                        </tr>
                    `).join('');
                });
        }
        
        // Load orders on page load
        loadOrders();
        setInterval(loadOrders, 5000); // Refresh every 5 seconds
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return HTML_TEMPLATE

@app.route('/save_settings', methods=['POST'])
def save_settings():
    data = request.json
    with open('.env', 'w') as f:
        f.write(f"SITE_USERNAME={data['username']}\n")
        f.write(f"SITE_PASSWORD={data['password']}\n")
        f.write(f"SITE_URL={data['url']}\n")
    return jsonify({'success': True})

@app.route('/toggle_bot', methods=['POST'])
def toggle_bot():
    global bot_running, bot_thread
    
    if not bot_running:
        bot_running = True
        stop_event.clear()
        bot_thread = threading.Thread(target=run_bot_thread, daemon=True)
        bot_thread.start()
    else:
        bot_running = False
        stop_event.set()
    
    return jsonify({'running': bot_running})

def run_bot_thread():
    # Your bot logic here
    pass

if __name__ == '__main__':
    print("Opening web browser at http://localhost:5000")
    import webbrowser
    webbrowser.open('http://localhost:5000')
    app.run(debug=False, port=5000)