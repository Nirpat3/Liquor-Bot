import asyncio
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv
import logging
from typing import Optional
from pathlib import Path
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
           slow_mo = 1000 # slows actions by 1 second for now
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
           input()
           
           # save authentication state after successful login
           await self.save_auth_state()
           
       except:
           # might already be logged in if using saved auth
           logger.info("Login form not found - might already be logged in")
   
   async def order_item(self, item_number: int, quantity: int):
       # function to navigate and order items
       logger.info(f"Starting order process for item #{item_number} with quantity {quantity}...")
       
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
       
       # input item number into search field
       logger.info(f"Entering item number: {item_number}")
       await self.page.fill('input[id="Dm-8"]', str(item_number))
       
       logger.info(f"Item number {item_number} entered in search field")
       
       # press Enter key after inputting item number
       logger.info("Pressing Enter key...")
       await self.page.press('input[id="Dm-8"]', 'Enter')
       
       # wait for page to load/update after pressing Enter
       await self.page.wait_for_load_state('networkidle')
       
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
           # continue with original quantity if can't read available
       
       # TODO: Add code here to input the quantity into the appropriate field
       # await self.page.fill('input[id="quantity_field_id"]', str(quantity))
       
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
       
       # click Next button on the next page
       logger.info("Clicking Next button on second page...")
       await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
       
       # wait for page to load
       await self.page.wait_for_load_state('networkidle')
       
       # click Submit button on the next page
       logger.info("Clicking Next button on second page...")
       #await self.page.click('span.ActionButtonCaptionText:has-text("Submit")')
   
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

async def main():
   # create bot with visible browser
   bot = WebAutomationBot(headless=False)
   
   # run setup (use_saved_auth=True by default, set to False for first run)
   await bot.setup(use_saved_auth=True)
   
   # order an item for now
   item_to_order = 38178
   quantity_to_order = 100
   await bot.order_item(item_to_order, quantity_to_order)
   
   # wait to see the result
   await asyncio.sleep(5)
   
   await bot.cleanup()

if __name__ == "__main__":
   asyncio.run(main())