import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import asyncio
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv
import logging
from typing import Optional
from pathlib import Path
import csv
import time
import sys
from datetime import datetime
# import bot class
from bot_script import WebAutomationBot, read_csv_file, update_csv_file

# set up logging to capture to both file and GUI
class GuiLogHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        # thread-safe GUI update
        self.text_widget.after(0, append)

class BotGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mississippi DOR Order Bot")
        self.root.geometry("900x700")
        
        # variables
        self.bot_running = False
        self.bot_thread = None
        self.bot = None
        self.stop_event = threading.Event()
        
        # create main frames
        self.create_widgets()
        
        # load environment variables if they exist
        self.load_env_settings()
        
        # check for CSV file
        self.check_csv_file()
        
    def create_widgets(self):
        # create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # settings tab
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text='Settings')
        self.create_settings_tab(settings_frame)
        
        # csv editor tab
        csv_frame = ttk.Frame(notebook)
        notebook.add(csv_frame, text='Orders (CSV)')
        self.create_csv_tab(csv_frame)
        
        # bot control tab
        control_frame = ttk.Frame(notebook)
        notebook.add(control_frame, text='Bot Control')
        self.create_control_tab(control_frame)
        
        # log tab
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text='Logs')
        self.create_log_tab(log_frame)
        
    def create_settings_tab(self, parent):
        # frame for settings
        settings_frame = ttk.LabelFrame(parent, text="Login Settings", padding=10)
        settings_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # username
        ttk.Label(settings_frame, text="Username:").grid(row=0, column=0, sticky='w', pady=5)
        self.username_entry = ttk.Entry(settings_frame, width=40)
        self.username_entry.grid(row=0, column=1, pady=5, padx=5)
        
        # password
        ttk.Label(settings_frame, text="Password:").grid(row=1, column=0, sticky='w', pady=5)
        self.password_entry = ttk.Entry(settings_frame, width=40, show='*')
        self.password_entry.grid(row=1, column=1, pady=5, padx=5)
        
        # site url
        ttk.Label(settings_frame, text="Site URL:").grid(row=2, column=0, sticky='w', pady=5)
        self.url_entry = ttk.Entry(settings_frame, width=40)
        self.url_entry.grid(row=2, column=1, pady=5, padx=5)
        self.url_entry.insert(0, "https://tap.dor.ms.gov/")
        
        # headless mode checkbox
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings_frame, text="Run in background (headless mode)", 
                       variable=self.headless_var).grid(row=3, column=0, columnspan=2, pady=10)
        
        # use saved auth checkbox
        self.use_saved_auth_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings_frame, text="Use saved authentication (skip login if possible)", 
                       variable=self.use_saved_auth_var).grid(row=4, column=0, columnspan=2, pady=5)
        
        # save settings button
        ttk.Button(settings_frame, text="Save Settings", 
                  command=self.save_settings).grid(row=5, column=0, columnspan=2, pady=20)
        
        # instructions
        instructions = ttk.LabelFrame(parent, text="Instructions", padding=10)
        instructions.pack(fill='both', padx=10, pady=10)
        
        inst_text = """1. Enter your login credentials and save settings
2. Add items to order in the 'Orders (CSV)' tab
3. Go to 'Bot Control' tab and click 'Start Bot'
4. For first-time use, complete 2FA manually when prompted
5. The bot will continuously process orders until stopped"""
        
        ttk.Label(instructions, text=inst_text, justify='left').pack()
        
    def create_csv_tab(self, parent):
        # frame for csv controls
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(control_frame, text="Load CSV", command=self.load_csv).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Save CSV", command=self.save_csv).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Add Row", command=self.add_csv_row).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Delete Row", command=self.delete_csv_row).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear All 'Yes' Marks", command=self.clear_yes_marks).pack(side='left', padx=5)
        
        # frame for csv table
        table_frame = ttk.Frame(parent)
        table_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # create treeview for csv
        columns = ('Item Number', 'Quantity', 'Order Filled')
        self.csv_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # define headings
        for col in columns:
            self.csv_tree.heading(col, text=col)
            self.csv_tree.column(col, width=150)
        
        # scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.csv_tree.yview)
        scrollbar.pack(side='right', fill='y')
        self.csv_tree.configure(yscrollcommand=scrollbar.set)
        
        self.csv_tree.pack(side='left', fill='both', expand=True)
        
        # entry frame for adding items
        entry_frame = ttk.LabelFrame(parent, text="Add/Edit Item", padding=10)
        entry_frame.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(entry_frame, text="Item Number:").grid(row=0, column=0, padx=5)
        self.item_entry = ttk.Entry(entry_frame, width=20)
        self.item_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(entry_frame, text="Quantity:").grid(row=0, column=2, padx=5)
        self.quantity_entry = ttk.Entry(entry_frame, width=20)
        self.quantity_entry.grid(row=0, column=3, padx=5)
        
        ttk.Button(entry_frame, text="Add Item", command=self.add_item).grid(row=0, column=4, padx=20)
        
    def create_control_tab(self, parent):
        # bot control frame
        control_frame = ttk.LabelFrame(parent, text="Bot Control", padding=20)
        control_frame.pack(fill='x', padx=10, pady=10)
        
        # start/stop button
        self.start_button = ttk.Button(control_frame, text="Start Bot", 
                                      command=self.toggle_bot, state='normal')
        self.start_button.pack(pady=10)
        
        # status label
        self.status_label = ttk.Label(control_frame, text="Status: Stopped", 
                                     font=('Arial', 12, 'bold'))
        self.status_label.pack(pady=10)
        
        # statistics frame
        stats_frame = ttk.LabelFrame(parent, text="Statistics", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=10)
        
        self.stats_label = ttk.Label(stats_frame, text="Items Processed: 0\nItems Remaining: 0\nOrders Submitted: 0")
        self.stats_label.pack()
        
        # progress bar
        self.progress = ttk.Progressbar(parent, mode='indeterminate')
        self.progress.pack(fill='x', padx=10, pady=10)
        
    def create_log_tab(self, parent):
        # log text area
        self.log_text = scrolledtext.ScrolledText(parent, state='disabled', 
                                                  wrap='word', height=20)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)
        
        # clear log button
        ttk.Button(parent, text="Clear Logs", 
                  command=self.clear_logs).pack(pady=5)
        
        # set up logging handler
        log_handler = GuiLogHandler(self.log_text)
        log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                                  datefmt='%H:%M:%S'))
        logging.getLogger().addHandler(log_handler)
        logging.getLogger().setLevel(logging.INFO)
        
    def load_env_settings(self):
        # load from .env file if exists
        if Path('.env').exists():
            load_dotenv()
            self.username_entry.insert(0, os.getenv('SITE_USERNAME', ''))
            self.password_entry.insert(0, os.getenv('SITE_PASSWORD', ''))
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, os.getenv('SITE_URL', 'https://tap.dor.ms.gov/'))
            
    def save_settings(self):
        # save settings to .env file
        with open('.env', 'w') as f:
            f.write(f"SITE_USERNAME={self.username_entry.get()}\n")
            f.write(f"SITE_PASSWORD={self.password_entry.get()}\n")
            f.write(f"SITE_URL={self.url_entry.get()}\n")
        
        # reload environment variables
        load_dotenv(override=True)
        messagebox.showinfo("Success", "Settings saved successfully!")
        
    def check_csv_file(self):
        # check if orders.csv exists and load it
        if Path('orders.csv').exists():
            self.load_csv()
        
    def load_csv(self):
        # clear existing items
        for item in self.csv_tree.get_children():
            self.csv_tree.delete(item)
        
        # load from csv
        if Path('orders.csv').exists():
            with open('orders.csv', 'r', newline='') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    self.csv_tree.insert('', 'end', values=(
                        row.get('item_number', ''),
                        row.get('quantity', ''),
                        row.get('order_filled', '')
                    ))
        
        self.update_stats()
        
    def save_csv(self):
        # save treeview data to csv
        with open('orders.csv', 'w', newline='') as file:
            fieldnames = ['item_number', 'quantity', 'order_filled']
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            
            for item in self.csv_tree.get_children():
                values = self.csv_tree.item(item)['values']
                writer.writerow({
                    'item_number': values[0],
                    'quantity': values[1],
                    'order_filled': values[2]
                })
        
        messagebox.showinfo("Success", "CSV saved successfully!")
        
    def add_item(self):
        # add item from entry fields
        item_num = self.item_entry.get()
        quantity = self.quantity_entry.get()
        
        if item_num and quantity:
            self.csv_tree.insert('', 'end', values=(item_num, quantity, ''))
            self.item_entry.delete(0, tk.END)
            self.quantity_entry.delete(0, tk.END)
            self.save_csv()
            self.update_stats()
        
    def add_csv_row(self):
        # add empty row
        self.csv_tree.insert('', 'end', values=('', '', ''))
        
    def delete_csv_row(self):
        # delete selected row
        selected = self.csv_tree.selection()
        if selected:
            self.csv_tree.delete(selected)
            self.save_csv()
            self.update_stats()
            
    def clear_yes_marks(self):
        # clear all 'yes' marks in order_filled column
        for item in self.csv_tree.get_children():
            values = list(self.csv_tree.item(item)['values'])
            values[2] = ''  # clear order_filled
            self.csv_tree.item(item, values=values)
        self.save_csv()
        self.update_stats()
        messagebox.showinfo("Success", "All 'yes' marks cleared!")
        
    def update_stats(self):
        # update statistics
        total = len(self.csv_tree.get_children())
        filled = sum(1 for item in self.csv_tree.get_children() 
                    if str(self.csv_tree.item(item)['values'][2]).lower() == 'yes')
        remaining = total - filled
        
        self.stats_label.config(text=f"Items Total: {total}\nItems Processed: {filled}\nItems Remaining: {remaining}")
        
    def toggle_bot(self):
        if not self.bot_running:
            self.start_bot()
        else:
            self.stop_bot()
            
    def start_bot(self):
        # validate settings
        if not self.username_entry.get() or not self.password_entry.get():
            messagebox.showerror("Error", "Please enter username and password in Settings tab!")
            return
        
        # save current csv
        self.save_csv()
        
        # update UI
        self.bot_running = True
        self.start_button.config(text="Stop Bot")
        self.status_label.config(text="Status: Running", foreground='green')
        self.progress.start()
        
        # clear stop event
        self.stop_event.clear()
        
        # start bot in separate thread
        self.bot_thread = threading.Thread(target=self.run_bot_thread, daemon=True)
        self.bot_thread.start()
        
    def stop_bot(self):
        # set stop event
        self.stop_event.set()
        
        # update UI
        self.bot_running = False
        self.start_button.config(text="Start Bot")
        self.status_label.config(text="Status: Stopped", foreground='red')
        self.progress.stop()
        
        logging.info("Bot stopped by user")
        
    def run_bot_thread(self):
        # create new event loop for this thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        
        try:
            loop.run_until_complete(self.run_bot())
        except Exception as e:
            logging.error(f"Bot error: {e}")
        finally:
            loop.close()
            
            # update UI on main thread
            self.root.after(0, self.on_bot_stopped)
            
    def on_bot_stopped(self):
        # called when bot stops
        self.bot_running = False
        self.start_button.config(text="Start Bot")
        self.status_label.config(text="Status: Stopped", foreground='red')
        self.progress.stop()
        self.load_csv()  # reload csv to show updates
        self.update_stats()
        
    async def run_bot(self):
        # import bot class
        from bot_script import WebAutomationBot, read_csv_file, update_csv_file
        
        headless = self.headless_var.get()
        bot = WebAutomationBot(headless=headless)
        
        try:
            await bot.setup(use_saved_auth=self.use_saved_auth_var.get())
            logging.info("Bot ready. If prompted for OTP, complete it in the browser.")
            
            while not self.stop_event.is_set():
                try:
                    # read items from csv
                    items = read_csv_file('orders.csv')
                    
                    # filter unfilled items
                    unfilled_items = [item for item in items 
                                    if item.get('order_filled', '').lower() != 'yes']
                    
                    if not unfilled_items:
                        logging.info("All items completed! Checking for new items in 5 seconds...")
                        await asyncio.sleep(5)
                        self.root.after(0, self.load_csv)
                        self.root.after(0, self.update_stats)
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
                    
                    # update UI
                    self.root.after(0, self.load_csv)
                    self.root.after(0, self.update_stats)
                
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
            
    def clear_logs(self):
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')

def main():
    root = tk.Tk()
    app = BotGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()