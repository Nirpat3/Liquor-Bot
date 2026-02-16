# Save as web_gui.py
from flask import Flask, render_template_string, request, jsonify, send_file
import threading
import asyncio
import csv
import json
from pathlib import Path
import logging
import os
from dotenv import load_dotenv
import sys

app = Flask(__name__)
bot_thread = None
bot_running = False
stop_event = threading.Event()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# HTML template
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Mississippi DOR Order Bot</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f0f0f0; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #333; border-bottom: 3px solid #007bff; padding-bottom: 10px; }
        .tabs { display: flex; gap: 10px; margin-bottom: 20px; }
        .tab { padding: 10px 20px; background: #f0f0f0; cursor: pointer; border-radius: 5px; transition: all 0.3s; }
        .tab:hover { background: #e0e0e0; }
        .tab.active { background: #007bff; color: white; }
        .tab-content { display: none; padding: 20px; background: #f9f9f9; border-radius: 5px; }
        .tab-content.active { display: block; }
        input, select { padding: 8px; margin: 5px; border: 1px solid #ddd; border-radius: 4px; }
        button { padding: 10px 20px; margin: 5px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; }
        button:hover { background: #0056b3; }
        button.danger { background: #dc3545; }
        button.danger:hover { background: #c82333; }
        button.success { background: #28a745; }
        button.success:hover { background: #218838; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th { background: #007bff; color: white; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        tr:nth-child(even) { background: #f9f9f9; }
        .status { font-size: 20px; font-weight: bold; margin: 20px 0; padding: 10px; border-radius: 5px; }
        .running { background: #d4edda; color: #155724; }
        .stopped { background: #f8d7da; color: #721c24; }
        #logs-content { background: #2c3e50; color: #ecf0f1; padding: 15px; height: 400px; overflow-y: scroll; font-family: monospace; font-size: 12px; border-radius: 5px; }
        .log-entry { margin: 2px 0; }
        .stats-box { background: #e9ecef; padding: 15px; border-radius: 5px; margin: 20px 0; }
        .input-group { margin: 10px 0; }
        .input-group label { display: inline-block; width: 120px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🛒 Mississippi DOR Order Bot</h1>
        
        <div class="tabs">
            <div class="tab active" onclick="showTab('settings')">⚙️ Settings</div>
            <div class="tab" onclick="showTab('orders')">📋 Orders</div>
            <div class="tab" onclick="showTab('control')">🎮 Control</div>
            <div class="tab" onclick="showTab('logs')">📜 Logs</div>
        </div>
        
        <div id="settings" class="tab-content active">
            <h2>Login Settings</h2>
            <div class="input-group">
                <label>Username:</label>
                <input type="text" id="username" placeholder="Enter username" style="width: 300px;">
            </div>
            <div class="input-group">
                <label>Password:</label>
                <input type="password" id="password" placeholder="Enter password" style="width: 300px;">
            </div>
            <div class="input-group">
                <label>Site URL:</label>
                <input type="text" id="url" value="https://tap.dor.ms.gov/" style="width: 300px;">
            </div>
            <div class="input-group">
                <label>Headless Mode:</label>
                <input type="checkbox" id="headless"> Run in background (no browser window)
            </div>
            <button class="success" onclick="saveSettings()">💾 Save Settings</button>
            
            <div class="stats-box" style="margin-top: 30px;">
                <h3>📖 Instructions</h3>
                <ol>
                    <li>Enter your login credentials above and save</li>
                    <li>Add items to order in the 'Orders' tab</li>
                    <li>Go to 'Control' tab and click 'Start Bot'</li>
                    <li>For first-time use, complete 2FA manually in the browser</li>
                    <li>The bot will continuously process orders until stopped</li>
                </ol>
            </div>
        </div>
        
        <div id="orders" class="tab-content">
            <h2>Order Items</h2>
            <div style="margin-bottom: 20px;">
                <input type="number" id="item_number" placeholder="Item Number" style="width: 150px;">
                <input type="number" id="quantity" placeholder="Quantity" style="width: 150px;">
                <button class="success" onclick="addItem()">➕ Add Item</button>
                <button onclick="loadOrders()">🔄 Refresh</button>
                <button class="danger" onclick="clearCompleted()">🗑️ Clear Completed</button>
                <button onclick="downloadCSV()">📥 Download CSV</button>
                <button onclick="document.getElementById('file-upload').click()">📤 Upload CSV</button>
                <input type="file" id="file-upload" style="display: none;" accept=".csv" onchange="uploadCSV(event)">
            </div>
            <table id="orders-table">
                <thead>
                    <tr>
                        <th>Item Number</th>
                        <th>Quantity</th>
                        <th>Status</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
        
        <div id="control" class="tab-content">
            <h2>Bot Control</h2>
            <button id="bot-toggle" class="success" onclick="toggleBot()" style="font-size: 18px; padding: 15px 30px;">▶️ Start Bot</button>
            <div id="status" class="status stopped">⭕ Status: Stopped</div>
            <div class="stats-box">
                <h3>📊 Statistics</h3>
                <div id="stats">Loading...</div>
            </div>
        </div>
        
        <div id="logs" class="tab-content">
            <h2>System Logs</h2>
            <button class="danger" onclick="clearLogs()">🗑️ Clear Logs</button>
            <button onclick="loadLogs()">🔄 Refresh Logs</button>
            <div id="logs-content"></div>
        </div>
    </div>
    
    <script>
        let logsInterval;
        let ordersInterval;
        
        function showTab(tabName) {
            document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            event.target.classList.add('active');
            document.getElementById(tabName).classList.add('active');
            
            if (tabName === 'logs') {
                loadLogs();
                logsInterval = setInterval(loadLogs, 2000);
            } else {
                clearInterval(logsInterval);
            }
            
            if (tabName === 'orders') {
                loadOrders();
                ordersInterval = setInterval(loadOrders, 5000);
            } else {
                clearInterval(ordersInterval);
            }
        }
        
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
            }).then(() => {
                alert('✅ Settings saved successfully!');
                loadSettings();
            });
        }
        
        function loadSettings() {
            fetch('/get_settings')
                .then(r => r.json())
                .then(data => {
                    document.getElementById('username').value = data.username || '';
                    document.getElementById('password').value = data.password || '';
                    document.getElementById('url').value = data.url || 'https://tap.dor.ms.gov/';
                    document.getElementById('headless').checked = data.headless || false;
                });
        }
        
        function addItem() {
            const item_number = document.getElementById('item_number').value;
            const quantity = document.getElementById('quantity').value;
            
            if (!item_number || !quantity) {
                alert('Please enter both item number and quantity');
                return;
            }
            
            fetch('/add_item', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({
                    item_number: item_number,
                    quantity: quantity
                })
            }).then(() => {
                document.getElementById('item_number').value = '';
                document.getElementById('quantity').value = '';
                loadOrders();
            });
        }
        
        function deleteItem(index) {
            if (confirm('Delete this item?')) {
                fetch('/delete_item', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({index: index})
                }).then(() => loadOrders());
            }
        }
        
        function clearCompleted() {
            if (confirm('Clear all completed items?')) {
                fetch('/clear_completed', {method: 'POST'})
                    .then(() => loadOrders());
            }
        }
        
        function toggleBot() {
            fetch('/toggle_bot', {method: 'POST'})
                .then(r => r.json())
                .then(data => updateStatus(data));
        }
        
        function updateStatus(data) {
            const statusDiv = document.getElementById('status');
            const toggleBtn = document.getElementById('bot-toggle');
            
            if (data.running) {
                statusDiv.className = 'status running';
                statusDiv.innerHTML = '✅ Status: Running';
                toggleBtn.innerHTML = '⏹️ Stop Bot';
                toggleBtn.className = 'danger';
            } else {
                statusDiv.className = 'status stopped';
                statusDiv.innerHTML = '⭕ Status: Stopped';
                toggleBtn.innerHTML = '▶️ Start Bot';
                toggleBtn.className = 'success';
            }
            
            loadStats();
        }
        
        function loadOrders() {
            fetch('/get_orders')
                .then(r => r.json())
                .then(data => {
                    const tbody = document.querySelector('#orders-table tbody');
                    tbody.innerHTML = data.map((item, index) => `
                        <tr>
                            <td>${item.item_number}</td>
                            <td>${item.quantity}</td>
                            <td>${item.order_filled === 'yes' ? '✅ Completed' : '⏳ Pending'}</td>
                            <td>
                                <button class="danger" onclick="deleteItem(${index})">Delete</button>
                            </td>
                        </tr>
                    `).join('');
                    loadStats();
                });
        }
        
        function loadStats() {
            fetch('/get_stats')
                .then(r => r.json())
                .then(data => {
                    document.getElementById('stats').innerHTML = `
                        <p>📦 Total Items: ${data.total}</p>
                        <p>✅ Completed: ${data.completed}</p>
                        <p>⏳ Remaining: ${data.remaining}</p>
                    `;
                });
        }
        
        function loadLogs() {
            fetch('/get_logs')
                .then(r => r.json())
                .then(data => {
                    const logsDiv = document.getElementById('logs-content');
                    logsDiv.innerHTML = data.logs.map(log => 
                        `<div class="log-entry">${log}</div>`
                    ).join('');
                    logsDiv.scrollTop = logsDiv.scrollHeight;
                });
        }
        
        function clearLogs() {
            if (confirm('Clear all logs?')) {
                fetch('/clear_logs', {method: 'POST'})
                    .then(() => loadLogs());
            }
        }
        
        function downloadCSV() {
            window.location.href = '/download_csv';
        }
        
        function uploadCSV(event) {
            const file = event.target.files[0];
            const formData = new FormData();
            formData.append('file', file);
            
            fetch('/upload_csv', {
                method: 'POST',
                body: formData
            }).then(() => {
                alert('CSV uploaded successfully!');
                loadOrders();
            });
        }
        
        // Check bot status periodically
        setInterval(() => {
            fetch('/get_status')
                .then(r => r.json())
                .then(data => updateStatus(data));
        }, 3000);
        
        // Load initial data
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
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            orders = list(reader)
    return jsonify(orders)

@app.route('/add_item', methods=['POST'])
def add_item():
    data = request.json
    
    # Read existing orders
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            orders = list(reader)
    
    # Add new item
    orders.append({
        'item_number': data['item_number'],
        'quantity': data['quantity'],
        'order_filled': ''
    })
    
    # Write back to CSV
    with open('orders.csv', 'w', newline='') as file:
        fieldnames = ['item_number', 'quantity', 'order_filled']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(orders)
    
    return jsonify({'success': True})

@app.route('/delete_item', methods=['POST'])
def delete_item():
    data = request.json
    index = data['index']
    
    # Read orders
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            orders = list(reader)
    
    # Delete item
    if 0 <= index < len(orders):
        orders.pop(index)
    
    # Write back
    with open('orders.csv', 'w', newline='') as file:
        fieldnames = ['item_number', 'quantity', 'order_filled']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(orders)
    
    return jsonify({'success': True})

@app.route('/clear_completed', methods=['POST'])
def clear_completed():
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            orders = [row for row in reader if row.get('order_filled', '').lower() != 'yes']
    
    with open('orders.csv', 'w', newline='') as file:
        fieldnames = ['item_number', 'quantity', 'order_filled']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(orders)
    
    return jsonify({'success': True})

@app.route('/get_stats')
def get_stats():
    orders = []
    if Path('orders.csv').exists():
        with open('orders.csv', 'r', newline='') as file:
            reader = csv.DictReader(file)
            orders = list(reader)
    
    total = len(orders)
    completed = sum(1 for o in orders if o.get('order_filled', '').lower() == 'yes')
    remaining = total - completed
    
    return jsonify({
        'total': total,
        'completed': completed,
        'remaining': remaining
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
    return jsonify({'logs': logs[-50:]})  # Return last 50 logs

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

def run_bot_thread():
    """Run the bot in a separate thread with error recovery matching bot_script.py main loop"""
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
            
            # Ensure orders.csv exists
            if not Path('orders.csv').exists():
                with open('orders.csv', 'w', newline='') as f:
                    f.write('item_number,quantity,order_filled\n')
                logging.info("Created empty orders.csv")
            
            try:
                await bot.setup(use_saved_auth=True)
                logging.info("Bot ready. If prompted for OTP, complete it in the browser.")
                
                while not stop_event.is_set():
                    try:
                        items = read_csv_file('orders.csv')
                        unfilled_items = [item for item in items 
                                        if item.get('order_filled', '').lower() != 'yes']
                        
                        if not unfilled_items:
                            logging.info("All items completed! Checking for new items in 5 seconds...")
                            await asyncio.sleep(5)
                            continue
                        
                        logging.info(f"Checking {len(unfilled_items)} items for availability...")
                        
                        # Start a new order if not already on the order page
                        current_url = bot.page.url
                        if 'itemEntry' not in current_url:
                            await bot.start_order()
                        
                        items_found, total_qty_added = await bot.check_and_process_items(items)
                        
                        if items_found and total_qty_added >= 10:
                            item_numbers = [str(item['item_number']) for item in items_found]
                            logging.info(f"Found {len(items_found)} items, {total_qty_added} total qty: {', '.join(item_numbers)}")
                            await bot.submit_order()
                            update_csv_file('orders.csv', items)
                            logging.info("Immediately checking for remaining items...")
                        elif items_found and total_qty_added < 10:
                            logging.warning(f"Need min 10 qty total (have {total_qty_added}). Reverting - will retry.")
                            for item in items_found:
                                item['order_filled'] = ''
                        else:
                            logging.info("No items available. Checking again in 1 second...")
                            await asyncio.sleep(1)
                    
                    except Exception as e:
                        logging.error(f"Error during bot operation: {e}", exc_info=True)
                        
                        # Re-initialize the bot to recover from crash
                        logging.info("Attempting to recover by re-initializing the bot...")
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
                logging.error(f"Bot startup error: {e}", exc_info=True)
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