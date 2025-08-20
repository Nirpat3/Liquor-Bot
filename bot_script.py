import asyncio
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv
import logging
from typing import Optional
from pathlib import Path
import csv
import time
# set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# loading env variables
load_dotenv()

class WebAutomationBot:
    def __init__(self, headless: bool = False):
        self.headless = headless
        self.browser = None
        self.context = None
        self.page = None
        self.playwright = None
    
    async def setup(self, use_saved_auth: bool = True):
        # starts playwright
        self.playwright = await async_playwright().start()
        
        # launch broswer with GUI
        self.browser = await self.playwright.chromium.launch(
            headless = self.headless,
            slow_mo = 0 # slows actions by 1 second for now
        )
        
        # check if auth state file exists and use it
        auth_file = "auth_state.json"
        if use_saved_auth and Path(auth_file).exists():
            # creates browser context with saved auth
            logger.info("Loading saved authentication state...")
            self.context = await self.browser.new_context(
                storage_state=auth_file  # loads cookies/localStorage
            )
        else:
            # creates new browser context without saved auth
            logger.info("Creating new browser context...")
            self.context = await self.browser.new_context()
        
        # set page timeout to never timeout
        self.context.set_default_timeout(0)
        
        self.page = await self.context.new_page()
        
        # navigates to new website
        site_url = os.getenv('SITE_URL')
        await self.page.goto(site_url)
        
        # wait for page to load
        await self.page.wait_for_load_state('networkidle')
        
        # check for session message button and click if exists
        try:
            # check if "Click Here to Start Over" button exists
            start_over_button = await self.page.wait_for_selector('a:has-text("Click Here to Start Over")', timeout=3000)
            if start_over_button:
                logger.info("Session message found, clicking Start Over...")
                await self.page.click('a:has-text("Click Here to Start Over")')
                # wait for page to reload after clicking
                await self.page.wait_for_load_state('networkidle')
                logger.info("Session cleared, proceeding to login...")
        except:
            # button not found, continue normally
            logger.info("No session message, proceeding...")
        
        # check if login form exists (might already be logged in)
        try:
            # fills username and password
            username = os.getenv('SITE_USERNAME')
            await self.page.fill('input[aria-label="Username"]', username)
            
            password = os.getenv('SITE_PASSWORD')
            await self.page.fill('input[aria-label="Password"]', password)
            
            # login
            await self.page.click('button:has-text("Log in")')
            
            # wait for login to complete
            await self.page.wait_for_load_state('networkidle')
            
            # pause for manual 2FA if needed
            logger.info("If 2FA is required, please complete it in the browser...")
            logger.info("Press Enter when ready to continue...")
            #input()
            
            # save authentication state after successful login
            await self.save_auth_state()
            
        except:
            # might already be logged in if using saved auth
            logger.info("Login form not found - might already be logged in")
    
    async def navigate_to_home(self):
        # navigate back to home page to start new order
        logger.info("Navigating back to home page...")
        site_url = os.getenv('SITE_URL')
        await self.page.goto(site_url)
        await self.page.wait_for_load_state('networkidle')
    
    async def process_multiple_items(self, items):
        # navigate to order page once
        logger.info("Starting order process for multiple items...")
        
        # clicks on add/view retail orders
        logger.info("Clicking Add/View Retail Orders...")
        await self.page.click('span:has-text("Add/View Retail Orders")')
        
        # wait for page to load after clicking
        await self.page.wait_for_load_state('networkidle')
        logger.info("Navigated to retail orders page")
        
        # clicking add order
        logger.info("Clicking Add Order...")
        await self.page.click('span:has-text("Add Order")')
        
        # wait for page to load after clicking
        await self.page.wait_for_load_state('networkidle')
        logger.info("Navigated to add order page")
        
        # click on manually enter items button
        logger.info("Selecting manually enter item button...")
        await self.page.click('input[type="radio"][id="Dl-q"]')
        
        # click next button to go to item entry page
        logger.info("Clicking Next button...")
        await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
        
        # wait for page to load after clicking next
        await self.page.wait_for_load_state('networkidle')
        logger.info("Navigated to item entry page")
        
        # track items added to this order
        items_added_count = 0
        
        # process each item
        for item in items:
            item_number = item['item_number']
            quantity = item['quantity']
            
            if item.get('order_filled', '').lower() == 'yes':
                logger.info(f"Skipping item {item_number} - already filled")
                continue
            
            try:
                # input item number into search field
                logger.info(f"Entering item number: {item_number}")
                await self.page.fill('input[id="Dm-8"]', str(item_number))
                
                logger.info(f"Item number {item_number} entered in search field")
                
                # press Enter key after inputting item number
                logger.info("Pressing Enter key...")
                await self.page.press('input[id="Dm-8"]', 'Enter')
                
                # wait for page to load/update after pressing Enter
                await self.page.wait_for_load_state('networkidle')
                
                # check if Add Item button appears
                try:
                    # wait for Add Item button with short timeout
                    await self.page.wait_for_selector('span:has-text("Add Item")', timeout=3000)
                    
                    # read available quantity
                    try:
                        available_text = await self.page.text_content('span[id="fgvt_Dm-m-1"]')
                        # remove comma and convert to int
                        available_quantity = int(available_text.replace(',', ''))
                        logger.info(f"Available quantity: {available_quantity}")
                        
                        # adjust quantity if needed
                        if quantity > available_quantity:
                            # set quantity to 70% of available
                            adjusted_quantity = int(available_quantity * 0.7)
                            logger.info(f"Requested quantity {quantity} exceeds available. Adjusting to 70% of available: {adjusted_quantity}")
                            quantity = adjusted_quantity
                        else:
                            logger.info(f"Requested quantity {quantity} is available")
                            
                    except Exception as e:
                        logger.warning(f"Could not read available quantity: {e}")
                    
                    # clicking add item button
                    logger.info("Clicking Add Item button...")
                    await self.page.click('span:has-text("Add Item")')
                    
                    # wait for page to load/update after adding item
                    await self.page.wait_for_load_state('networkidle')
                    logger.info(f"Item {item_number} with quantity {quantity} added successfully")
                    
                    # input the quantity into the quantity field
                    logger.info(f"Entering quantity: {quantity}")
                    await self.page.fill('input[id="Dm_1-81"]', str(quantity))
                    
                    # hit ok button
                    logger.info("Clicking OK button...")
                    await self.page.click('button[data-event="AcceptDocModal"]')
                    
                    # wait for modal to close
                    await self.page.wait_for_load_state('networkidle')
                    
                    # mark item as processed
                    item['order_filled'] = 'yes'
                    items_added_count += 1
                    
                except:
                    logger.warning(f"Item {item_number} not found or Add Item button didn't appear. Moving to next item...")
                    # clear the search field for next item
                    await self.page.fill('input[id="Dm-8"]', '')
                    continue
                    
            except Exception as e:
                logger.error(f"Error processing item {item_number}: {e}")
                continue
        
        # only proceed with checkout if items were added
        if items_added_count > 0:
            logger.info(f"{items_added_count} items added. Proceeding to checkout...")
            
            # click Next button
            logger.info("Clicking Next button...")
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            
            # wait for page to load
            await self.page.wait_for_load_state('networkidle')
            
            # click Next button on the next page
            logger.info("Clicking Next button on second page...")
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            
            # wait for page to load
            await self.page.wait_for_load_state('networkidle')
            
            # click Next button on the third page
            logger.info("Clicking Next button on third page...")
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            
            # wait for page to load
            await self.page.wait_for_load_state('networkidle')
            
            # click Submit button on the final page
            logger.info("Clicking Submit button...")
            await self.page.click('span.ActionButtonCaptionText:has-text("Submit")')
            
            logger.info(f"Order submitted successfully with {items_added_count} items!")
            
            # wait for submission to complete
            await self.page.wait_for_load_state('networkidle')
            await asyncio.sleep(3)
            
            # navigate back to home page for next order
            await self.navigate_to_home()
            
            return True  # indicate that an order was submitted
        else:
            logger.info("No items were added to this order")
            return False  # no order was submitted
    
    async def save_auth_state(self):
        # save the authentication state for future runs
        logger.info("Saving authentication state...")
        await self.context.storage_state(path="auth_state.json")
        logger.info("Authentication state saved to auth_state.json")
    
    async def cleanup(self):
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

def read_csv_file(filename):
    # read csv file and return list of items
    items = []
    try:
        with open(filename, 'r', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                items.append({
                    'item_number': int(row['item_number']),
                    'quantity': int(row['quantity']),
                    'order_filled': row.get('order_filled', '').strip()
                })
        logger.info(f"Loaded {len(items)} items from {filename}")
        return items
    except Exception as e:
        logger.error(f"Error reading CSV file: {e}")
        return []

def update_csv_file(filename, items):
    # update csv file with order_filled status
    try:
        with open(filename, 'w', newline='') as file:
            fieldnames = ['item_number', 'quantity', 'order_filled']
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(items)
        logger.info(f"Updated {filename} with order status")
    except Exception as e:
        logger.error(f"Error updating CSV file: {e}")

def create_sample_csv():
    # create a sample csv file for reference
    sample_filename = 'orders_template.csv'
    with open(sample_filename, 'w', newline='') as file:
        fieldnames = ['item_number', 'quantity', 'order_filled']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([
            {'item_number': '12345', 'quantity': '10', 'order_filled': ''},
            {'item_number': '67890', 'quantity': '5', 'order_filled': ''},
            {'item_number': '11111', 'quantity': '20', 'order_filled': ''}
        ])
    logger.info(f"Created sample CSV template: {sample_filename}")
    return sample_filename

async def main():
    # check for orders.csv file
    csv_filename = 'orders.csv'
    
    if not Path(csv_filename).exists():
        logger.warning(f"{csv_filename} not found. Creating sample template...")
        create_sample_csv()
        logger.info(f"Please fill in orders_template.csv and rename it to {csv_filename}")
        return
    
    # read items from csv
    items = read_csv_file(csv_filename)
    
    if not items:
        logger.error("No items to process")
        return
    
    # create bot with visible browser
    bot = WebAutomationBot(headless=False)
    
    try:
        # run setup (use_saved_auth=True by default, set to False for first run)
        await bot.setup(use_saved_auth=True)
        
        # continuous loop to process orders
        while True:
            # filter items that haven't been filled
            unfilled_items = [item for item in items if item.get('order_filled', '').lower() != 'yes']
            
            if not unfilled_items:
                logger.info("All items have been filled. Checking for new items...")
                # reload csv to check for new items
                items = read_csv_file(csv_filename)
                unfilled_items = [item for item in items if item.get('order_filled', '').lower() != 'yes']
                
                if not unfilled_items:
                    logger.info("No new items. Waiting 2 seconds before checking again...")
                    await asyncio.sleep(2)
                    continue
            
            logger.info(f"Processing {len(unfilled_items)} unfilled items...")
            
            # process multiple items - will return True if order was submitted
            order_submitted = await bot.process_multiple_items(unfilled_items)
            
            # update csv with filled status
            update_csv_file(csv_filename, items)
            
            # if order was submitted and there are still unfilled items, continue immediately
            unfilled_items = [item for item in items if item.get('order_filled', '').lower() != 'yes']
            if order_submitted and unfilled_items:
                logger.info(f"Order submitted. {len(unfilled_items)} items remaining. Starting new order...")
                continue  # immediately start new order for remaining items
            
            # wait before next check
            logger.info("Waiting 30 seconds before checking for more items...")
            await asyncio.sleep(30)
            
    except KeyboardInterrupt:
        logger.info("Script interrupted by user")
    except Exception as e:
        logger.error(f"Error in main loop: {e}")
    finally:
        await bot.cleanup()

if __name__ == "__main__":
    asyncio.run(main())