import asyncio
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv
import logging
from typing import Optional
# set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# loading env variables
load_dotenv()

class WebAutomationBot:
    def __init__(self, headless: bool = False):
        self.headless = headless
        self.browser = None
        self.page = None
        self.playwright = None
    
    async def setup(self):
        # starts playwright
        self.playwright = await async_playwright().start()
        
        # launch broswer with GUI
        self.browser = await self.playwright.chromium.launch(
            headless = self.headless,
            slow_mo = 1000 # slows actions by 1 second for now
        )
        # creates browser context
        context = await self.broswer.new_context()
        self.page = await context.new_page()
        
        # navigates to new website
        site_url = os.getenv('SITE_URL')
        await self.page.goto(site_url)
        
        # wait for page to load
        await self.page.wait_for_load_state('networkidle')
        
        # fills username and password
        username = os.getenv('SITE_USERNAME')
        await self.page.fill('input[aria-label="Username"]', username)
        
        password = os.getenv('SITE_PASSWORD')
        await self.page.fill('input[aria-label="Password"]', password)
        
        # login
        await self.page.click('button:has-text("Log in")')
    
    async def cleanup(self):
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

async def main():
   # create bot with visible browser
   bot = WebAutomationBot(headless=False)
   
   # run setup
   await bot.setup()
   
   # wait to see the result
   await asyncio.sleep(5)
   
   await bot.cleanup()

if __name__ == "__main__":
   asyncio.run(main())