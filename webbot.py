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
  
  # wait to see the result
  await asyncio.sleep(5)
  
  await bot.cleanup()

if __name__ == "__main__":
  asyncio.run(main())