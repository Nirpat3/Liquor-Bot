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
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
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
        """Initialize browser and login if needed"""
        # starts playwright
        self.playwright = await async_playwright().start()
        
        # launch browser with GUI
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=0  # no delay for faster checking
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
        
        # navigates to website
        site_url = os.getenv('SITE_URL')
        logger.info(f"Navigating to {site_url}")
        await self.page.goto(site_url)
        
        # wait for page to load
        await self.page.wait_for_load_state('networkidle')
        
        # check for session message button and click if exists
        try:
            start_over_button = await self.page.wait_for_selector('a:has-text("Click Here to Start Over")', timeout=3000)
            if start_over_button:
                logger.info("Session message found, clicking Start Over...")
                await self.page.click('a:has-text("Click Here to Start Over")')
                await self.page.wait_for_load_state('networkidle')
                logger.info("Session cleared, proceeding...")
        except:
            logger.info("No session message, proceeding...")
        
        # check if we need to login
        login_needed = False
        try:
            # check if username field exists (indicates login page)
            await self.page.wait_for_selector('input[aria-label="Username"]', timeout=3000)
            login_needed = True
            logger.info("Login required, entering credentials...")
            
            # fills username and password
            username = os.getenv('SITE_USERNAME')
            password = os.getenv('SITE_PASSWORD')
            
            await self.page.fill('input[aria-label="Username"]', username)
            await self.page.fill('input[aria-label="Password"]', password)
            
            # login
            await self.page.click('button:has-text("Log in")')
            
            # wait for login to complete
            await self.page.wait_for_load_state('networkidle')
            logger.info("Login submitted...")
            
            # wait longer for page to fully load after login
            await asyncio.sleep(3)
            
            # check if 2FA is needed or if we're logged in
            logged_in = False
            
            # Use the exact selector for the Add/View Retail Orders element
            main_selector = 'span.IconCaptionText:has-text("Add/View Retail Orders")'
            
            try:
                # First try with short timeout
                await self.page.wait_for_selector(main_selector, timeout=5000)
                logger.info("Login successful! Found Add/View Retail Orders button")
                logged_in = True
            except:
                # 2FA might be required or page is still loading
                logger.info("Waiting for page to load or 2FA completion...")
                logger.info("Please complete any 2FA in the browser if prompted.")
                
                # Wait with periodic checks and better feedback
                for i in range(60):  # 60 seconds total
                    try:
                        # Check if element exists
                        element = await self.page.query_selector(main_selector)
                        if element:
                            logger.info("✓ Login/2FA completed successfully!")
                            logged_in = True
                            break
                        
                        # Also check for any error messages
                        error_element = await self.page.query_selector('text=/error|invalid|incorrect/i')
                        if error_element:
                            logger.warning("Possible login error detected. Check credentials.")
                    except:
                        pass
                    
                    await asyncio.sleep(1)
                    if i % 10 == 0 and i > 0:
                        logger.info(f"Still waiting for login completion... ({60-i} seconds remaining)")
                
                if not logged_in:
                    # Take screenshot for debugging
                    logger.error("Could not find main page after login")
                    await self.page.screenshot(path="login_timeout.png")
                    logger.info("Screenshot saved as login_timeout.png for debugging")
                    
                    # Log current URL for debugging
                    current_url = self.page.url
                    logger.info(f"Current URL: {current_url}")
                    
                    # Try to log page content for debugging
                    try:
                        page_text = await self.page.text_content('body')
                        if page_text:
                            logger.info(f"Page contains text (first 200 chars): {page_text[:200]}...")
                    except:
                        pass
                    
                    raise Exception("Login failed or page structure unexpected. Check login_timeout.png")
            
            # save authentication state after successful login
            if logged_in:
                await self.save_auth_state()
            
        except Exception as e:
            if not login_needed:
                logger.info("Already logged in with saved authentication")
            else:
                logger.error(f"Login process error: {e}")
                raise
        
        # verify we're on the main page using the correct selector
        try:
            main_selector = 'span.IconCaptionText:has-text("Add/View Retail Orders")'
            await self.page.wait_for_selector(main_selector, timeout=10000)
            logger.info("✓ Bot is ready to process orders")
        except:
            logger.error("Could not verify main page loaded correctly")
            logger.info("Looking for debugging information...")
            
            # Try to find what's on the page
            try:
                all_spans = await self.page.query_selector_all('span.IconCaptionText')
                if all_spans:
                    logger.info(f"Found {len(all_spans)} IconCaptionText spans:")
                    for span in all_spans[:5]:  # Log first 5
                        text = await span.text_content()
                        logger.info(f"  - {text}")
            except:
                pass
            
            await self.page.screenshot(path="error_page.png")
            logger.info("Screenshot saved as error_page.png")
            raise Exception("Main page did not load correctly - check error_page.png")
    
    async def navigate_to_home(self):
        """Navigate back to home page to start new order"""
        logger.info("Navigating back to home page...")
        site_url = os.getenv('SITE_URL')
        await self.page.goto(site_url)
        await self.page.wait_for_load_state('networkidle')
        # wait for main page to load with correct selector
        main_selector = 'span.IconCaptionText:has-text("Add/View Retail Orders")'
        await self.page.wait_for_selector(main_selector, timeout=10000)
    
    async def start_order(self):
        """Navigate to the item entry page to start a new order"""
        # make sure we're on home page
        await self.navigate_to_home()
        
        # clicks on add/view retail orders using the correct selector
        logger.info("Starting new order...")
        main_selector = 'span.IconCaptionText:has-text("Add/View Retail Orders")'
        await self.page.click(main_selector)
        await self.page.wait_for_load_state('networkidle')
        
        # clicking add order
        await self.page.click('span:has-text("Add Order")')
        await self.page.wait_for_load_state('networkidle')
        
        # click on manually enter items button
        await self.page.click('input[type="radio"][id="Dl-q"]')
        
        # wait a moment for radio button to register
        await asyncio.sleep(0.5)
        
        # click next button to go to item entry page
        await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
        await self.page.wait_for_load_state('networkidle')
        logger.info("Ready to search for items")
    
    async def check_and_process_items(self, items):
        """Check all unfilled items and add any available ones to cart"""
        items_found = []
        
        for item in items:
            item_number = item['item_number']
            quantity = int(item['quantity'])
            
            # skip if already filled
            if item.get('order_filled', '').lower() == 'yes':
                continue
            
            # clear search field and enter item number
            logger.info(f"Checking item #{item_number}...")
            await self.page.fill('input[id="Dm-8"]', '')
            await asyncio.sleep(0.2)  # small delay to ensure field is cleared
            await self.page.fill('input[id="Dm-8"]', str(item_number))
            
            # press Enter to search
            await self.page.press('input[id="Dm-8"]', 'Enter')
            await self.page.wait_for_load_state('networkidle')
            
            # check if Add Item button appears (indicates item is available)
            try:
                await self.page.wait_for_selector('span:has-text("Add Item")', timeout=2000)
                logger.info(f"  ✓ Item #{item_number} is AVAILABLE!")
                
                # read available quantity and adjust if needed
                try:
                    available_text = await self.page.text_content('span[id="fgvt_Dm-m-1"]')
                    available_quantity = int(available_text.replace(',', ''))
                    logger.info(f"    Available quantity: {available_quantity}")
                    
                    if quantity > available_quantity:
                        adjusted_quantity = int(available_quantity * 0.7)
                        logger.info(f"    Adjusting quantity from {quantity} to {adjusted_quantity} (70% of available)")
                        quantity = adjusted_quantity
                    else:
                        logger.info(f"    Using requested quantity: {quantity}")
                        
                except Exception as e:
                    logger.warning(f"    Could not read available quantity, using requested: {quantity}")
                
                # click Add Item button
                await self.page.click('span:has-text("Add Item")')
                await self.page.wait_for_load_state('networkidle')
                
                # input quantity in the modal
                await self.page.fill('input[id="Dm_1-81"]', str(quantity))
                await asyncio.sleep(0.2)
                
                # click OK button to confirm
                await self.page.click('button[data-event="AcceptDocModal"]')
                await self.page.wait_for_load_state('networkidle')
                
                # mark item as processed
                item['order_filled'] = 'yes'
                items_found.append(item)
                logger.info(f"    ✓ Added to cart: {quantity} units")
                
            except:
                logger.info(f"  ✗ Item #{item_number} not available")
                # clear search field for next item
                await self.page.fill('input[id="Dm-8"]', '')
                await asyncio.sleep(0.1)
        
        return items_found
    
    async def submit_order(self):
        """Complete checkout and submit the current order"""
        logger.info("Proceeding to checkout...")
        
        try:
            # click Next button (first page)
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            await self.page.wait_for_load_state('networkidle')
            await asyncio.sleep(0.5)
            
            # click Next button (second page)
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            await self.page.wait_for_load_state('networkidle')
            await asyncio.sleep(0.5)
            
            # click Next button (third page)
            await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
            await self.page.wait_for_load_state('networkidle')
            await asyncio.sleep(0.5)
            
            # click Submit button (final page)
            await self.page.click('span.ActionButtonCaptionText:has-text("Submit")')
            await self.page.wait_for_load_state('networkidle')
            
            logger.info("✓ Order submitted successfully!")
            await asyncio.sleep(2)  # wait for confirmation
            
        except Exception as e:
            logger.error(f"Error during checkout: {e}")
            raise
    
    async def save_auth_state(self):
        """Save authentication state for future runs"""
        logger.info("Saving authentication state...")
        await self.context.storage_state(path="auth_state.json")
        logger.info("Authentication state saved")
    
    async def cleanup(self):
        """Clean up browser resources"""
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

def read_csv_file(filename):
    """Read CSV file and return list of items"""
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
        return items
    except Exception as e:
        logger.error(f"Error reading CSV file: {e}")
        return []

def update_csv_file(filename, items):
    """Update CSV file with order_filled status"""
    try:
        with open(filename, 'w', newline='') as file:
            fieldnames = ['item_number', 'quantity', 'order_filled']
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(items)
        logger.info("CSV file updated")
    except Exception as e:
        logger.error(f"Error updating CSV file: {e}")

def create_sample_csv():
    """Create a sample CSV template"""
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

async def main():
    """Main bot loop"""
    csv_filename = 'orders.csv'
    
    # check if CSV exists
    if not Path(csv_filename).exists():
        logger.warning(f"{csv_filename} not found. Creating sample template...")
        create_sample_csv()
        logger.info(f"Please fill in orders_template.csv and rename it to {csv_filename}")
        return
    
    # create bot instance
    bot = WebAutomationBot(headless=False)
    
    try:
        # setup browser and login
        logger.info("Initializing bot...")
        await bot.setup(use_saved_auth=True)
        
        logger.info("\n" + "="*60)
        logger.info("BOT STARTED - Continuously checking for items")
        logger.info("="*60 + "\n")
        
        while True:
            try:
                # reload CSV each iteration to get latest data
                items = read_csv_file(csv_filename)
                
                # get unfilled items
                unfilled_items = [item for item in items if item.get('order_filled', '').lower() != 'yes']
                
                if not unfilled_items:
                    logger.info("All items completed! Checking for new items in 5 seconds...")
                    await asyncio.sleep(5)
                    continue
                
                logger.info(f"\n--- Checking {len(unfilled_items)} items for availability ---")
                
                # Start a new order only if we are not already on the order page
                current_url = bot.page.url
                if not "itemEntry" in current_url:
                    await bot.start_order()
                
                # check all items and add available ones to cart
                items_found = await bot.check_and_process_items(items)
                
                if items_found:
                    # at least one item found - submit order
                    item_numbers = [str(item['item_number']) for item in items_found]
                    logger.info(f"\n✓ Found {len(items_found)} items: {', '.join(item_numbers)}")
                    
                    # submit the order
                    await bot.submit_order()
                    
                    # update CSV with ordered items
                    update_csv_file(csv_filename, items)
                    
                    # immediately check for remaining items (no delay)
                    logger.info("Immediately checking for remaining items...")
                    
                else:
                    # no items found - wait 1 second before checking again
                    logger.info("No items available. Checking again in 1 second...")
                    await asyncio.sleep(1)
            
            except Exception as e:
                logger.error(f"An error occurred during bot operation: {e}")
                import traceback
                traceback.print_exc()
                
                # Try to re-initialize the bot to recover from a potential crash
                logger.info("Attempting to recover by re-initializing the bot...")
                await bot.cleanup()
                bot = WebAutomationBot(headless=False)
                await bot.setup(use_saved_auth=True)
                
    except KeyboardInterrupt:
        logger.info("\nBot stopped by user")
    except Exception as e:
        logger.error(f"\nFATAL BOT ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        logger.info("Cleaning up...")
        await bot.cleanup()
        logger.info("Bot shutdown complete")

if __name__ == "__main__":
    print("\n" + "="*60)
    print("Mississippi DOR Order Bot")
    print("="*60)
    print("\nPress Ctrl+C to stop the bot\n")
    
    asyncio.run(main())