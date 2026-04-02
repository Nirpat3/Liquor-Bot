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

if sys.platform == 'win32':
    import msvcrt
    def _lock_shared(f):
        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
    def _lock_exclusive(f):
        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
    def _unlock(f):
        try:
            msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
else:
    import fcntl
    def _lock_shared(f):
        fcntl.flock(f.fileno(), fcntl.LOCK_SH)
    def _lock_exclusive(f):
        fcntl.flock(f.fileno(), fcntl.LOCK_EX)
    def _unlock(f):
        fcntl.flock(f.fileno(), fcntl.LOCK_UN)

# set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# loading env variables
load_dotenv()


# ── CSV helpers with file locking ──

def read_csv_file(filename):
    """Read CSV file with file locking to prevent race conditions."""
    items = []
    try:
        with open(filename, 'r', newline='') as file:
            _lock_shared(file)
            try:
                reader = csv.DictReader(file)
                for row in reader:
                    items.append({
                        'item_number': int(row['item_number']),
                        'quantity': int(row['quantity']),
                        'name': row.get('name', ''),
                        'size': row.get('size', ''),
                        'units': row.get('units', ''),
                        'order_filled': row.get('order_filled', '').strip()
                    })
            finally:
                _unlock(file)
        return items
    except Exception as e:
        logger.error(f"Error reading CSV file: {e}")
        return []


def update_csv_file(filename, items):
    """Update CSV file with file locking to prevent race conditions."""
    try:
        with open(filename, 'w', newline='') as file:
            _lock_exclusive(file)
            try:
                fieldnames = ['item_number', 'quantity', 'name', 'size', 'units', 'order_filled']
                writer = csv.DictWriter(file, fieldnames=fieldnames, extrasaction='ignore')
                writer.writeheader()
                writer.writerows(items)
            finally:
                _unlock(file)
        logger.info("CSV file updated")
    except Exception as e:
        logger.error(f"Error updating CSV file: {e}")


def create_sample_csv():
    """Create a sample CSV template"""
    sample_filename = 'orders_template.csv'
    with open(sample_filename, 'w', newline='') as file:
        fieldnames = ['item_number', 'quantity', 'name', 'size', 'units', 'order_filled']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows([
            {'item_number': '12345', 'quantity': '10', 'name': '', 'size': '', 'units': '', 'order_filled': ''},
            {'item_number': '67890', 'quantity': '5', 'name': '', 'size': '', 'units': '', 'order_filled': ''},
            {'item_number': '11111', 'quantity': '20', 'name': '', 'size': '', 'units': '', 'order_filled': ''}
        ])
    logger.info(f"Created sample CSV template: {sample_filename}")


# ── Selector constants ──

# Selectors that appear on the site (Glide/ServiceNow-style dynamic IDs)
MAIN_PAGE_SELECTOR = 'span.IconCaptionText:has-text("Add/View Retail Orders")'
ADD_ITEM_SELECTORS = [
    'a:has-text("Add Item")',
    'td:has-text("Add Item")',
    'button:has-text("Add Item")',
    'span:has-text("Add Item")',
    'div[role="button"]:has-text("Add Item")',
    'div:has(> span:has-text("Add Item"))',
]
PASSWORD_SELECTORS = [
    'input#Dn-k',
    'input[name="Dn-k"]',
    'input.DocControlPassword',
    'input[type="password"]',
    'input[aria-label="Password"]',
]
SEARCH_INPUT_SELECTORS = [
    'input[id="Dm-8"]',
    'input[placeholder*="Item"]',
    'input[id^="Dm"]',
    'input[type="text"]',
]
QTY_AVAILABLE_SELECTORS = [
    'span[id="fgvt_Dm-m-1"]',
    'span[id^="fgvt_Dm"]',
    'span[id^="fgvt_"]',
]


class WebAutomationBot:
    def __init__(self, headless: bool = False):
        self.headless = headless
        self.browser = None
        self.context = None
        self.page = None
        self.playwright = None
        self._content_frame = None  # frame that has item entry (may be iframe)

    async def setup(self, use_saved_auth: bool = True):
        """Initialize browser and login if needed"""
        self.playwright = await async_playwright().start()

        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            slow_mo=0,
            args=['--start-maximized']
        )

        auth_file = "auth_state.json"
        if use_saved_auth and Path(auth_file).exists():
            logger.info("Loading saved authentication state...")
            self.context = await self.browser.new_context(
                storage_state=auth_file,
                no_viewport=True
            )
        else:
            logger.info("Creating new browser context...")
            self.context = await self.browser.new_context(no_viewport=True)

        self.context.set_default_timeout(30000)
        self.page = await self.context.new_page()

        site_url = os.getenv('SITE_URL')
        logger.info(f"Navigating to {site_url}")
        await self.page.goto(site_url)
        await self.page.wait_for_load_state('networkidle')

        # Handle "Start Over" session message
        await self._handle_session_message()

        # Login if needed
        login_needed = False
        try:
            await self.page.wait_for_selector('input[aria-label="Username"]', timeout=3000)
            login_needed = True
            logger.info("Login required, entering credentials...")

            username = os.getenv('SITE_USERNAME')
            password = os.getenv('SITE_PASSWORD')

            await self.page.fill('input[aria-label="Username"]', username)
            await self.page.fill('input[aria-label="Password"]', password)
            await self.page.click('button:has-text("Log in")')
            await self.page.wait_for_load_state('networkidle')
            await asyncio.sleep(3)

            logged_in = await self._wait_for_main_page(timeout=60)

            if not logged_in:
                await self.page.screenshot(path="login_timeout.png")
                logger.error(f"Login failed. Current URL: {self.page.url}")
                raise Exception("Login failed or page structure unexpected. Check login_timeout.png")

            await self.save_auth_state()

        except Exception as e:
            if not login_needed:
                logger.info("Already logged in with saved authentication")
            else:
                logger.error(f"Login process error: {e}")
                raise

        # Verify main page
        try:
            await self.page.wait_for_selector(MAIN_PAGE_SELECTOR, timeout=10000)
            logger.info("Bot is ready to process orders")
        except Exception:
            logger.error("Could not verify main page loaded correctly")
            await self._debug_page("error_page.png")
            raise Exception("Main page did not load correctly - check error_page.png")

    async def _handle_session_message(self):
        """Handle 'Click Here to Start Over' session message."""
        try:
            start_over = await self.page.wait_for_selector(
                'a:has-text("Click Here to Start Over")', timeout=3000
            )
            if start_over:
                logger.info("Session message found, clicking Start Over...")
                try:
                    async with self.context.expect_page(timeout=5000) as page_info:
                        await start_over.click()
                    new_page = await page_info.value
                    await self.page.close()
                    self.page = new_page
                    logger.info("Using single tab (closed old)")
                except Exception:
                    await self.page.wait_for_load_state('networkidle')
                await self.page.wait_for_load_state('networkidle')
                logger.info("Session cleared, proceeding...")
        except Exception:
            logger.info("No session message, proceeding...")

    async def _wait_for_main_page(self, timeout: int = 60) -> bool:
        """Wait for main page to load after login/2FA, with periodic checks."""
        try:
            await self.page.wait_for_selector(MAIN_PAGE_SELECTOR, timeout=5000)
            logger.info("Login successful! Found Add/View Retail Orders button")
            return True
        except Exception:
            pass

        logger.info("Waiting for page to load or 2FA completion...")
        logger.info("Please complete any 2FA in the browser if prompted.")

        for i in range(timeout):
            try:
                element = await self.page.query_selector(MAIN_PAGE_SELECTOR)
                if element:
                    logger.info("Login/2FA completed successfully!")
                    return True

                error_element = await self.page.query_selector('text=/error|invalid|incorrect/i')
                if error_element:
                    logger.warning("Possible login error detected. Check credentials.")
            except Exception:
                pass

            await asyncio.sleep(1)
            if i % 10 == 0 and i > 0:
                logger.info(f"Still waiting for login completion... ({timeout - i}s remaining)")

        return False

    async def _debug_page(self, screenshot_name: str):
        """Take screenshot and log page info for debugging."""
        try:
            await self.page.screenshot(path=screenshot_name)
            logger.info(f"Screenshot saved as {screenshot_name}")
        except Exception:
            pass
        try:
            all_spans = await self.page.query_selector_all('span.IconCaptionText')
            if all_spans:
                logger.info(f"Found {len(all_spans)} IconCaptionText spans:")
                for span in all_spans[:5]:
                    text = await span.text_content()
                    logger.info(f"  - {text}")
        except Exception:
            pass

    async def navigate_to_home(self):
        """Navigate back to home page to start new order"""
        logger.info("Navigating back to home page...")
        site_url = os.getenv('SITE_URL')
        await self.page.goto(site_url)
        await self.page.wait_for_load_state('networkidle')
        await self.page.wait_for_selector(MAIN_PAGE_SELECTOR, timeout=10000)

    async def start_order(self):
        """Navigate to the item entry page to start a new order"""
        await self.navigate_to_home()

        logger.info("Starting new order...")
        await self.page.click(MAIN_PAGE_SELECTOR)
        await self.page.wait_for_load_state('networkidle')

        await self.page.click('span:has-text("Add Order")')
        await self.page.wait_for_load_state('networkidle')
        await asyncio.sleep(1)

        # Click "Manually Enter Items" radio
        manually_clicked = await self._click_manually_enter()
        if not manually_clicked:
            raise Exception("Could not find/click 'Manually Enter Items' option")

        await asyncio.sleep(0.5)

        # Click Next to go to item entry page
        await self._scroll_to_bottom()
        await asyncio.sleep(0.2)
        await self._scroll_to_bottom()
        await self.page.click('span.ActionButtonCaptionText:has-text("Next")')
        await self.page.wait_for_load_state('networkidle')
        logger.info("Ready to search for items")

    async def _click_manually_enter(self) -> bool:
        """Click the 'Manually Enter Items' radio button."""
        # Try Playwright robust locators first
        for locator in [
            self.page.get_by_label("Manually Enter Items"),
            self.page.get_by_label("Manually enter items"),
            self.page.get_by_role("radio", name="Manually Enter Items"),
            self.page.get_by_role("radio", name="Manually enter items"),
        ]:
            try:
                if await locator.count() > 0:
                    await locator.first.click()
                    logger.info("Selected manually enter items (label/role)")
                    return True
            except Exception:
                continue

        # Try CSS selectors
        for selector in [
            'input[type="radio"][id="Dl-q"]',
            'label:has-text("Manually Enter")',
            'label:has-text("Manually enter")',
            'label:has-text("Manual Entry")',
            'label:has-text("Manually")',
            'span:has-text("Manually Enter")',
            'span:has-text("Manual Entry")',
        ]:
            try:
                el = await self.page.wait_for_selector(selector, timeout=2000)
                if el:
                    await el.click()
                    logger.info(f"Selected manually enter items via: {selector}")
                    return True
            except Exception:
                continue

        # Fallback: click first radio button
        try:
            radios = await self.page.query_selector_all('input[type="radio"]')
            if radios:
                await radios[0].click()
                logger.info("Selected first radio option (manually enter)")
                return True
        except Exception:
            pass

        return False

    async def _get_search_input(self):
        """Get the item search input - try main page and iframes."""
        if self._content_frame:
            for selector in SEARCH_INPUT_SELECTORS:
                try:
                    el = await self._content_frame.query_selector(selector)
                    if el and await el.is_visible():
                        return el
                except Exception:
                    continue

        for frame in [self.page] + list(self.page.frames):
            try:
                for selector in SEARCH_INPUT_SELECTORS:
                    try:
                        el = await frame.wait_for_selector(selector, timeout=500)
                        if el and await el.is_visible():
                            self._content_frame = frame
                            return el
                    except Exception:
                        continue
            except Exception:
                continue
        raise Exception("Could not find item search input")

    async def _click_add_item(self, hint_frame=None):
        """Click Add Item — restored original 8-method approach from working version.
        When hint_frame is provided, tries that frame first."""

        async def try_click_in_frame(frame):
            """Locator-based click across multiple selectors in a single frame."""
            for sel in ADD_ITEM_SELECTORS:
                try:
                    loc = frame.locator(sel).first
                    if await loc.count() == 0:
                        continue
                    await loc.scroll_into_view_if_needed()
                    await loc.click(force=True, timeout=2000)
                    logger.info(f"    [AddItem] Locator click succeeded: {sel}")
                    return True
                except Exception:
                    continue
            return False

        async def try_click():
            # === FAST PATH: hint_frame from availability check ===
            if hint_frame:
                logger.info(f"    [AddItem] Trying hint frame first...")
                try:
                    if await try_click_in_frame(hint_frame):
                        return True
                except Exception:
                    pass

            # Try content frame, then main page, then all other frames
            if self._content_frame and self._content_frame != hint_frame:
                if await try_click_in_frame(self._content_frame):
                    return True
            if await try_click_in_frame(self.page):
                return True
            for fr in self.page.frames:
                if fr != self.page.main_frame:
                    try:
                        if await try_click_in_frame(fr):
                            return True
                    except Exception:
                        pass

            # Use the frame where content was found for JS-based methods
            eval_page = hint_frame or self._content_frame or self.page

            # Method 1: JS — find "Add Item", get clickable parent (data-event/td), mouse events at center
            try:
                r = await eval_page.evaluate("""() => {
                    const all = document.querySelectorAll('span, div, td, button, a');
                    for (const e of all) {
                        const t = e.textContent && e.textContent.trim();
                        if ((t === 'Add Item' || t.includes('Add Item')) && t.length < 20 && e.offsetParent) {
                            const target = e.closest('[data-event]') || e.closest('td') || e.closest('div[role="button"]') || e.parentElement || e;
                            target.scrollIntoView({block: 'center'});
                            const rect = target.getBoundingClientRect();
                            const x = rect.left + rect.width/2, y = rect.top + rect.height/2;
                            ['mousedown','mouseup','click'].forEach(n => target.dispatchEvent(new MouseEvent(n, {bubbles:true,cancelable:true,view:window,clientX:x,clientY:y})));
                            return true;
                        }
                    }
                    return false;
                }""")
                if r:
                    logger.info("    [AddItem] Method 1 (JS data-event parent + mouse coords) succeeded")
                    return True
            except Exception:
                pass

            # Method 2: JS — direct .click() on element
            try:
                r = await eval_page.evaluate("""() => {
                    const all = document.querySelectorAll('span, div, td, button');
                    for (const e of all) {
                        const t = e.textContent && e.textContent.trim();
                        if ((t === 'Add Item' || t.includes('Add Item')) && t.length < 20 && e.offsetParent) {
                            e.scrollIntoView({block: 'center'});
                            e.click();
                            return true;
                        }
                    }
                    return false;
                }""")
                if r:
                    logger.info("    [AddItem] Method 2 (JS direct click) succeeded")
                    return True
            except Exception:
                pass

            # Method 3: Playwright — div[data-event] containing Add Item, mouse.click at center
            try:
                el = await eval_page.query_selector('div[data-event]:has(span:has-text("Add Item"))')
                if el:
                    await el.scroll_into_view_if_needed()
                    box = await el.bounding_box()
                    if box:
                        await self.page.mouse.click(box['x'] + box['width']/2, box['y'] + box['height']/2)
                        logger.info("    [AddItem] Method 3 (data-event div mouse.click) succeeded")
                        return True
            except Exception:
                pass

            # Method 4: Playwright — td containing Add Item
            try:
                el = await eval_page.query_selector('td:has(span:has-text("Add Item"))')
                if el:
                    await el.scroll_into_view_if_needed()
                    await el.click(force=True)
                    logger.info("    [AddItem] Method 4 (td click) succeeded")
                    return True
            except Exception:
                pass

            # Method 5: Playwright — span.ActionButtonCaptionText (Glide specific)
            try:
                el = await eval_page.query_selector('span.ActionButtonCaptionText:has-text("Add Item")')
                if el:
                    await el.scroll_into_view_if_needed()
                    await el.click(force=True)
                    logger.info("    [AddItem] Method 5 (ActionButtonCaptionText) succeeded")
                    return True
            except Exception:
                pass

            # Method 6: Playwright — any span with Add Item, mouse.click at bounding box center
            try:
                el = await eval_page.query_selector('span:has-text("Add Item")')
                if el:
                    await el.scroll_into_view_if_needed()
                    box = await el.bounding_box()
                    if box:
                        await self.page.mouse.click(box['x'] + box['width']/2, box['y'] + box['height']/2)
                        logger.info("    [AddItem] Method 6 (span mouse.click at coords) succeeded")
                        return True
            except Exception:
                pass

            # Method 7: Playwright getByText — flexible text match
            try:
                loc = eval_page.locator('text=Add Item')
                if await loc.count() > 0:
                    await loc.first.scroll_into_view_if_needed()
                    await loc.first.click(force=True)
                    logger.info("    [AddItem] Method 7 (getByText) succeeded")
                    return True
            except Exception:
                pass

            # Method 8: JS — case-insensitive "add" + "item" with parent walk-up
            try:
                r = await eval_page.evaluate("""() => {
                    const all = document.querySelectorAll('span, div, td, button, a');
                    for (const e of all) {
                        const t = (e.textContent || '').toLowerCase();
                        if (t.includes('add') && t.includes('item') && t.length < 20 && e.offsetParent) {
                            const target = e.closest('[data-event]') || e.closest('td') || e.parentElement || e;
                            target.scrollIntoView({block: 'center'});
                            target.click();
                            return true;
                        }
                    }
                    return false;
                }""")
                if r:
                    logger.info("    [AddItem] Method 8 (case-insensitive fallback) succeeded")
                    return True
            except Exception:
                pass

            logger.warning("    [AddItem] All 8 methods failed")
            return False

        try:
            result = await asyncio.wait_for(try_click(), timeout=15.0)
        except asyncio.TimeoutError:
            logger.warning("    _click_add_item timed out after 15s")
            result = False

        if result:
            await asyncio.sleep(0.6)
        return result

    async def _check_item_availability(self, item_number: int) -> dict:
        """Check if an item is available and return availability info.

        Returns dict with keys:
            available: bool - whether item can be added
            quantity: int - available quantity (0 if not available)
            reason: str - why item is not available (if not available)
            frame: the frame where qty/button was found (for direct clicking)
        """
        frames = self._get_search_frames()

        # Try each selector across all frames — fast check first (no wait)
        for ctx in frames:
            for sel in QTY_AVAILABLE_SELECTORS:
                try:
                    el = await ctx.query_selector(sel)
                    if el:
                        available_text = await el.text_content() or '0'
                        available_quantity = int(available_text.replace(',', '').strip())
                        if available_quantity == 0:
                            return {'available': False, 'quantity': 0, 'reason': 'available qty is 0', 'frame': None}
                        logger.info(f"    Found qty {available_quantity} via {sel}")
                        return {'available': True, 'quantity': available_quantity, 'reason': '', 'frame': ctx}
                except Exception:
                    continue

        # Brief wait in case page is still rendering (max 500ms per selector)
        for ctx in frames:
            for sel in QTY_AVAILABLE_SELECTORS:
                try:
                    el = await ctx.wait_for_selector(sel, timeout=500)
                    if el:
                        available_text = await el.text_content() or '0'
                        available_quantity = int(available_text.replace(',', '').strip())
                        if available_quantity == 0:
                            return {'available': False, 'quantity': 0, 'reason': 'available qty is 0', 'frame': None}
                        logger.info(f"    Found qty {available_quantity} via {sel} (waited)")
                        return {'available': True, 'quantity': available_quantity, 'reason': '', 'frame': ctx}
                except Exception:
                    continue

        # Last resort: check if Add Item button is visible (item exists but qty selector unknown)
        for ctx in frames:
            for sel in ADD_ITEM_SELECTORS:
                try:
                    el = await ctx.query_selector(sel)
                    if el and await el.is_visible():
                        logger.info(f"    Qty selector not found but Add Item button visible — proceeding")
                        return {'available': True, 'quantity': -1, 'reason': '', 'frame': ctx}
                except Exception:
                    continue

        return {'available': False, 'quantity': 0, 'reason': 'qty element not found', 'frame': None}

    def _get_search_frames(self):
        """Get frames to search in priority order."""
        frames = []
        if self._content_frame:
            frames.append(self._content_frame)
        frames.append(self.page)
        for fr in self.page.frames:
            if fr != self.page.main_frame and fr not in frames:
                frames.append(fr)
        return frames

    async def check_and_process_items(self, items):
        """Check all unfilled items and add available ones to cart.
        Skips items with order_filled='yes'. Short-circuits on zero-qty items."""
        items_found = []
        total_qty_added = 0
        self.trace_log = getattr(self, 'trace_log', [])

        for item in items:
            item_number = item['item_number']
            quantity = int(item['quantity'])

            if item.get('order_filled', '').lower() == 'yes':
                continue

            t_item_start = time.time()
            trace = {'item': item_number, 'steps': {}, 'result': '', 'total_ms': 0}

            # Human-like delay between searches (1-3s random)
            import random
            delay = random.uniform(1.0, 3.0)
            logger.info(f"Checking item #{item_number} (requested qty: {quantity}, status: {item.get('order_filled', 'pending')})... (waiting {delay:.1f}s)")
            await asyncio.sleep(delay)

            t0 = time.time()
            search_input = await self._get_search_input()
            await search_input.click()
            await search_input.fill('')
            await search_input.type(str(item_number), delay=15)
            trace['steps']['type_item'] = round((time.time() - t0) * 1000)

            t0 = time.time()
            await search_input.press('Enter')
            try:
                await self.page.wait_for_load_state('networkidle', timeout=5000)
            except Exception:
                pass
            # Extra settle time for Glide/ServiceNow to render results
            await asyncio.sleep(0.5)
            trace['steps']['search_wait'] = round((time.time() - t0) * 1000)

            # ORIGINAL FLOW: Check if Add Item button appears (= item is available)
            t0 = time.time()
            add_item_visible = None
            found_frame = None
            search_frames = ([self._content_frame] if self._content_frame else []) + [self.page] + list(self.page.frames)
            for frame in search_frames:
                if not frame:
                    continue
                for sel in ADD_ITEM_SELECTORS:
                    try:
                        add_item_visible = await frame.wait_for_selector(sel, timeout=2000)
                        if add_item_visible:
                            found_frame = frame
                            break
                    except Exception:
                        continue
                if add_item_visible:
                    break
            trace['steps']['check_button'] = round((time.time() - t0) * 1000)

            if not add_item_visible:
                t0 = time.time()
                logger.info(f"  x Item #{item_number} — Add Item button not found, skipping")
                try:
                    search_input = await self._get_search_input()
                    await search_input.triple_click()
                except Exception:
                    pass
                trace['steps']['clear_input'] = round((time.time() - t0) * 1000)
                trace['result'] = "SKIP (no Add Item button)"
                trace['total_ms'] = round((time.time() - t_item_start) * 1000)
                self.trace_log.append(trace)
                logger.info(f"  [TRACE] Item #{item_number}: {trace['total_ms']}ms — {' | '.join(f'{k}:{v}ms' for k,v in trace['steps'].items())}")
                continue

            # Item is available — Add Item button found
            logger.info(f"  + Item #{item_number} is AVAILABLE!")
            ctx = found_frame or self._content_frame or self.page

            # Read available quantity and adjust if needed
            t0 = time.time()
            available_quantity = -1
            backorder_qty = 0
            for sel in QTY_AVAILABLE_SELECTORS:
                try:
                    el = await ctx.wait_for_selector(sel, timeout=3000)
                    if el:
                        available_text = await el.text_content() or '0'
                        available_quantity = int(available_text.replace(',', '').strip())
                        logger.info(f"    Available quantity: {available_quantity} (via {sel})")
                        break
                except Exception:
                    continue
            trace['steps']['read_qty'] = round((time.time() - t0) * 1000)

            if available_quantity == 0:
                logger.info(f"  x Item #{item_number} available qty is 0, skipping")
                try:
                    search_input = await self._get_search_input()
                    await search_input.triple_click()
                except Exception:
                    pass
                trace['result'] = "SKIP (available qty is 0)"
                trace['total_ms'] = round((time.time() - t_item_start) * 1000)
                self.trace_log.append(trace)
                continue

            if available_quantity > 0:
                if quantity > available_quantity:
                    backorder_qty = quantity - available_quantity
                    logger.info(f"    Ordering {available_quantity} of {quantity} requested ({backorder_qty} → backorder)")
                    quantity = available_quantity
                else:
                    logger.info(f"    Using requested quantity: {quantity}")
            else:
                logger.warning(f"    Could not read available quantity, using requested: {quantity}")

            # Click Add Item button (original 8-method approach, 15s timeout)
            try:
                t0 = time.time()
                add_clicked = await self._click_add_item(hint_frame=found_frame)
                if add_clicked:
                    logger.info("    Add Item clicked, waiting for quantity modal...")
                if not add_clicked:
                    await self.page.screenshot(path="add_item_fail.png")
                    try:
                        html_snippet = await ctx.evaluate("""() => {
                            const spans = document.querySelectorAll('span');
                            return Array.from(spans).filter(s => s.textContent && s.textContent.includes('Add')).slice(0,10).map(s => ({text: s.textContent.trim().substring(0,50), tag: s.tagName, cls: s.className, id: s.id}));
                        }""")
                        logger.error(f"    Add Item elements on page: {html_snippet}")
                    except Exception:
                        pass
                    raise Exception("Failed to click Add Item - see add_item_fail.png")

                await self.page.wait_for_load_state('networkidle')
                await asyncio.sleep(0.4)
                trace['steps']['click_add'] = round((time.time() - t0) * 1000)

                # Enter quantity in modal
                t0 = time.time()
                await self._enter_quantity_in_modal(quantity, item_number)
                trace['steps']['enter_qty'] = round((time.time() - t0) * 1000)

                # Mark item as processed — update qty to what was actually ordered
                original_qty = int(item['quantity'])
                item['order_filled'] = 'yes'
                item['quantity'] = quantity
                items_found.append(item)
                total_qty_added += quantity

                # Create backorder entry for remaining qty
                if backorder_qty > 0:
                    backorder_item = {
                        'item_number': item_number,
                        'quantity': backorder_qty,
                        'name': item.get('name', ''),
                        'size': item.get('size', ''),
                        'units': item.get('units', ''),
                        'order_filled': 'backorder'
                    }
                    items.append(backorder_item)
                    logger.info(f"    + Backorder created: {backorder_qty} units for item #{item_number}")
                    trace['result'] = f"ADDED {quantity} units (+{backorder_qty} backorder)"
                else:
                    trace['result'] = f"ADDED {quantity} units"
                logger.info(f"    + Added to cart: {quantity} units (total: {total_qty_added})")

            except Exception as e:
                logger.error(f"  x Item #{item_number} failed: {e}")
                trace['result'] = f"FAILED: {e}"
                try:
                    # Dismiss any open modal/dialog
                    await self.page.keyboard.press('Escape')
                    await asyncio.sleep(0.2)
                    search_input = await self._get_search_input()
                    await search_input.triple_click()
                except Exception:
                    pass

            trace['total_ms'] = round((time.time() - t_item_start) * 1000)
            self.trace_log.append(trace)
            logger.info(f"  [TRACE] Item #{item_number}: {trace['total_ms']}ms — {' | '.join(f'{k}:{v}ms' for k,v in trace['steps'].items())}")

        return items_found, total_qty_added

    async def _enter_quantity_in_modal(self, quantity: int, item_number: int):
        """Enter quantity into the modal dialog and confirm."""
        qty_str = str(int(quantity))
        logger.info(f"    Entering quantity {qty_str} for item #{item_number}")

        all_frames = [self.page] + list(self.page.frames)

        # Method 1: JS-based (handles Glide framework)
        for frame in all_frames:
            try:
                result = await frame.evaluate(f"""() => {{
                    const qtyInput = document.querySelector('input[id="Ds_1-81"]')
                        || document.querySelector('input[id^="Ds_1"]')
                        || document.querySelector('input.DocControlQuantity')
                        || document.querySelector('input[name="Ds_1-81"]');
                    if (!qtyInput || !qtyInput.offsetParent) return false;
                    qtyInput.focus();
                    qtyInput.select();
                    qtyInput.value = '';
                    qtyInput.value = '{qty_str}';
                    qtyInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    qtyInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    qtyInput.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                    const dialog = document.querySelector('[role="dialog"]')
                        || document.querySelector('[class*="modal"]') || document;
                    const btns = dialog.querySelectorAll('button, span, div[role="button"], a');
                    for (const b of btns) {{
                        const t = (b.textContent || '').trim();
                        if (t === 'Add Item' || t === 'OK') {{
                            b.click();
                            return true;
                        }}
                    }}
                    return false;
                }}""")
                if result:
                    logger.info(f"    Entered quantity {qty_str} and clicked Add Item")
                    await self.page.wait_for_load_state('networkidle')
                    await asyncio.sleep(0.2)
                    return
            except Exception:
                continue

        # Method 2: Playwright-based
        for frame in all_frames:
            try:
                qty_input = (
                    await frame.query_selector('input[id="Ds_1-81"]')
                    or await frame.query_selector('input[id^="Ds_1"]')
                    or await frame.query_selector('input.DocControlQuantity')
                )
                if qty_input and await qty_input.is_visible():
                    await qty_input.click()
                    await qty_input.fill('')
                    await qty_input.type(qty_str, delay=20)
                    await asyncio.sleep(0.1)
                    add_btn = (
                        await frame.query_selector('[role="dialog"] button:has-text("Add Item"), [role="dialog"] span:has-text("Add Item")')
                        or await frame.query_selector('[role="dialog"] button:has-text("OK"), [role="dialog"] span:has-text("OK")')
                    )
                    if add_btn:
                        await add_btn.click()
                        logger.info(f"    Entered quantity {qty_str} and clicked Add Item")
                        await self.page.wait_for_load_state('networkidle')
                        await asyncio.sleep(0.2)
                        return
            except Exception:
                continue

        raise Exception("Could not enter quantity and click Add Item")

    async def _scroll_to_bottom(self):
        """Scroll page and all frames to bottom."""
        async def do_scroll(ctx):
            try:
                await ctx.evaluate('''() => {
                    window.scrollTo(0, document.body.scrollHeight);
                    const d = document.scrollingElement || document.documentElement;
                    if (d) d.scrollTop = d.scrollHeight;
                    document.querySelectorAll('[style*="overflow"], .Overflown, [class*="Scroll"]').forEach(el => {
                        if (el.scrollHeight > el.clientHeight) el.scrollTop = el.scrollHeight;
                    });
                }''')
            except Exception:
                pass

        await do_scroll(self.page)
        if self._content_frame:
            await do_scroll(self._content_frame)
        for fr in self.page.frames:
            if fr != self.page.main_frame:
                try:
                    await do_scroll(fr)
                except Exception:
                    pass

    async def _click_next_or_submit(self, text: str, timeout_ms: int = 3000):
        """Click Next or Submit button. Scrolls to bottom first."""
        await self._scroll_to_bottom()
        await asyncio.sleep(0.1)
        await self._scroll_to_bottom()

        for ctx in self._get_search_frames():
            for sel in [
                f'span.ActionButtonCaptionText:has-text("{text}")',
                f'span:has-text("{text}")',
                f'button:has-text("{text}")',
            ]:
                try:
                    btn = await ctx.wait_for_selector(sel, timeout=timeout_ms)
                    if btn and await btn.is_visible():
                        await btn.scroll_into_view_if_needed()
                        await btn.click()
                        return True
                except Exception:
                    continue

            # Try clicking parent via JS
            try:
                span = await ctx.wait_for_selector(
                    f'span.ActionButtonCaptionText:has-text("{text}")', timeout=timeout_ms
                )
                if span and await span.is_visible():
                    clicked = await span.evaluate("""el => {
                        const p = el.closest('button, div[role="button"], [class*="ActionButton"], [class*="Button"]') || el.parentElement;
                        if (p && p.offsetParent) { p.click(); return true; }
                        return false;
                    }""")
                    if clicked:
                        return True
            except Exception:
                pass
        return False

    async def _discover_password_inputs(self, ctx=None) -> list:
        """Discover all password-like input fields on the current page."""
        if ctx is None:
            ctx = self.page
        try:
            inputs = await ctx.evaluate("""() => {
                const results = [];
                for (const inp of document.querySelectorAll('input')) {
                    const type = (inp.type || '').toLowerCase();
                    const id = inp.id || '';
                    const name = inp.name || '';
                    const className = inp.className || '';
                    const ariaLabel = inp.getAttribute('aria-label') || '';
                    const placeholder = inp.getAttribute('placeholder') || '';
                    const visible = !!(inp.offsetParent || inp.offsetWidth || inp.offsetHeight);
                    const isPasswordLike = (
                        type === 'password' ||
                        className.toLowerCase().includes('password') ||
                        name.toLowerCase().includes('password') ||
                        id.toLowerCase().includes('password') ||
                        ariaLabel.toLowerCase().includes('password') ||
                        placeholder.toLowerCase().includes('password')
                    );
                    if (isPasswordLike && visible) {
                        results.push({
                            id, name, type, className, ariaLabel, placeholder,
                            selector: id ? `input#${id}` : (name ? `input[name="${name}"]` : `input.${className.split(' ')[0]}`)
                        });
                    }
                }
                return results;
            }""")
            return inputs or []
        except Exception as e:
            logger.warning(f"Password field discovery failed: {e}")
            return []

    async def _verify_password_entered(self, ctx=None) -> bool:
        """Verify that the password field actually has a value."""
        if ctx is None:
            ctx = self.page
        try:
            return await ctx.evaluate("""() => {
                const selectors = ['input#Dn-k', 'input[name="Dn-k"]', 'input.DocControlPassword', 'input[type="password"]'];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el && el.offsetParent && el.value && el.value.length > 0) return true;
                }
                return false;
            }""")
        except Exception:
            return False

    async def _fill_password_field(self, password: str) -> bool:
        """Enter confirmation password using multiple strategies with retry.

        Strategies:
        1. fill() — same as login
        2. JS focus + keyboard.type
        3. Playwright click + keyboard.type
        4. Agentic discovery — scan DOM
        5. Frame search
        6. Tab-into-field fallback
        """
        logger.info("Attempting password entry...")

        # Strategy 1: fill()
        logger.info("  Strategy 1: fill()...")
        for sel in PASSWORD_SELECTORS:
            try:
                el = await self.page.wait_for_selector(sel, timeout=2000)
                if el and await el.is_visible():
                    await el.scroll_into_view_if_needed()
                    await el.click()
                    await self.page.fill(sel, password)
                    await asyncio.sleep(0.1)
                    if await self._verify_password_entered():
                        logger.info(f"  Password entered via fill() on {sel}")
                        return True
            except Exception:
                continue

        # Strategy 2: JS focus + keyboard.type
        logger.info("  Strategy 2: JS focus + keyboard.type...")
        try:
            focused = await self.page.evaluate("""() => {
                const selectors = ['#Dn-k', 'input[name="Dn-k"]', 'input.DocControlPassword', 'input[type="password"]'];
                for (const sel of selectors) {
                    const el = document.querySelector(sel);
                    if (el && el.offsetParent) {
                        el.scrollIntoView({block: 'center'});
                        el.focus();
                        el.click();
                        el.value = '';
                        el.dispatchEvent(new Event('focus', {bubbles: true}));
                        return sel;
                    }
                }
                return null;
            }""")
            if focused:
                await asyncio.sleep(0.05)
                await self.page.keyboard.type(password, delay=10)
                await asyncio.sleep(0.1)
                if await self._verify_password_entered():
                    logger.info(f"  Password entered via JS focus + keyboard.type on {focused}")
                    return True
        except Exception as e:
            logger.warning(f"  Strategy 2 failed: {e}")

        # Strategy 3: Playwright click + keyboard.type
        logger.info("  Strategy 3: Playwright click + keyboard.type...")
        for sel in PASSWORD_SELECTORS:
            try:
                el = await self.page.wait_for_selector(sel, timeout=2000)
                if el and await el.is_visible():
                    await el.scroll_into_view_if_needed()
                    await el.click(force=True)
                    await asyncio.sleep(0.05)
                    await self.page.keyboard.press('Control+a')
                    await self.page.keyboard.press('Delete')
                    await self.page.keyboard.type(password, delay=10)
                    await asyncio.sleep(0.1)
                    if await self._verify_password_entered():
                        logger.info(f"  Password entered via click + keyboard.type on {sel}")
                        return True
            except Exception:
                continue

        # Strategy 4: Agentic discovery
        logger.info("  Strategy 4: Agentic discovery...")
        discovered = await self._discover_password_inputs()
        if discovered:
            logger.info(f"  Discovered {len(discovered)} password-like input(s)")
            for field_info in discovered:
                sel = field_info.get('selector', '')
                if not sel or sel.startswith('input.'):
                    if field_info.get('id'):
                        sel = f"input#{field_info['id']}"
                    elif field_info.get('name'):
                        sel = f"input[name=\"{field_info['name']}\"]"
                    else:
                        continue
                try:
                    el = await self.page.wait_for_selector(sel, timeout=2000)
                    if el and await el.is_visible():
                        await el.scroll_into_view_if_needed()
                        await el.click()
                        await self.page.fill(sel, password)
                        await asyncio.sleep(0.1)
                        if await self._verify_password_entered():
                            logger.info(f"  Password entered via discovered {sel}")
                            return True
                        # Try keyboard.type
                        await el.click(force=True)
                        await self.page.keyboard.press('Control+a')
                        await self.page.keyboard.press('Delete')
                        await self.page.keyboard.type(password, delay=10)
                        await asyncio.sleep(0.1)
                        if await self._verify_password_entered():
                            return True
                except Exception:
                    continue

        # Strategy 5: Search all frames
        logger.info("  Strategy 5: Searching all frames...")
        for fr in list(self.page.frames):
            if fr == self.page.main_frame:
                continue
            frame_discovered = await self._discover_password_inputs(ctx=fr)
            for sel in PASSWORD_SELECTORS + [d.get('selector', '') for d in frame_discovered]:
                if not sel:
                    continue
                try:
                    el = await fr.wait_for_selector(sel, timeout=1500)
                    if el and await el.is_visible():
                        await el.scroll_into_view_if_needed()
                        try:
                            await el.click()
                            await fr.fill(sel, password)
                            await asyncio.sleep(0.1)
                            if await self._verify_password_entered(ctx=fr):
                                logger.info(f"  Password entered via frame fill() on {sel}")
                                return True
                        except Exception:
                            pass
                        await el.click(force=True)
                        await asyncio.sleep(0.05)
                        await self.page.keyboard.press('Control+a')
                        await self.page.keyboard.press('Delete')
                        await self.page.keyboard.type(password, delay=10)
                        await asyncio.sleep(0.1)
                        if await self._verify_password_entered(ctx=fr):
                            logger.info(f"  Password entered via frame keyboard.type on {sel}")
                            return True
                except Exception:
                    continue

        # Strategy 6: Tab fallback
        logger.info("  Strategy 6: Tab into field fallback...")
        try:
            await self.page.keyboard.press('Tab')
            await asyncio.sleep(0.05)
            await self.page.keyboard.type(password, delay=10)
            await asyncio.sleep(0.1)
            if await self._verify_password_entered():
                logger.info("  Password entered via Tab + keyboard.type")
                return True
        except Exception as e:
            logger.warning(f"  Tab fallback failed: {e}")

        # All failed — diagnostics
        logger.error("All password entry strategies failed. Running diagnostics...")
        await self._debug_page("password_fail.png")
        try:
            all_inputs = await self.page.evaluate("""() => {
                return Array.from(document.querySelectorAll('input')).filter(e => e.offsetParent).map(e => ({
                    id: e.id, name: e.name, type: e.type,
                    class: e.className.substring(0, 60),
                    ariaLabel: e.getAttribute('aria-label') || '',
                    placeholder: e.placeholder || ''
                }));
            }""")
            logger.info(f"  Visible inputs on page: {all_inputs}")
        except Exception:
            pass

        return False

    async def _check_session_expired(self) -> bool:
        """Check if the session has expired."""
        try:
            el = await self.page.query_selector('a:has-text("Click Here to Start Over"), :has-text("session has expired")')
            if el and await el.is_visible():
                return True
        except Exception:
            pass
        return False

    async def _detect_page_state(self) -> str:
        """Fast JS-based page state detection.
        Returns: 'next', 'submit', 'password', 'ach', 'session_expired', 'unknown'"""
        try:
            return await self.page.evaluate("""() => {
                const body = document.body ? document.body.textContent || '' : '';
                if (body.includes('session has expired') || body.includes('Start Over')) return 'session_expired';

                const pwSelectors = ['#Dn-k', 'input[name="Dn-k"]', 'input.DocControlPassword', 'input[type="password"]'];
                for (const sel of pwSelectors) {
                    const el = document.querySelector(sel);
                    if (el && el.offsetParent) return 'password';
                }

                const allSpans = document.querySelectorAll('span.ActionButtonCaptionText, span');
                let hasNext = false, hasSubmit = false;
                for (const s of allSpans) {
                    const t = (s.textContent || '').trim();
                    if (t === 'Next' && s.offsetParent) hasNext = true;
                    if (t === 'Submit' && s.offsetParent) hasSubmit = true;
                }

                if (body.includes('ACH Debit Bank')) {
                    if (hasNext) return 'ach_with_next';
                    return 'ach';
                }

                if (hasSubmit && !hasNext) return 'submit';
                if (hasNext) return 'next';
                if (hasSubmit) return 'submit';

                return 'unknown';
            }""")
        except Exception:
            return 'unknown'

    async def submit_order(self):
        """Complete checkout and submit the full order.

        Agentic flow using fast JS-based page state detection:
        1. Navigate through checkout pages (Next/ACH/Submit)
        2. Enter confirmation password
        3. Final Submit
        """
        logger.info("Proceeding to checkout - placing full order...")

        try:
            max_steps = 12
            step = 0
            ach_selected = False

            # Phase 1: Navigate through checkout pages
            while step < max_steps:
                step += 1

                try:
                    await self.page.wait_for_load_state('networkidle', timeout=8000)
                except Exception:
                    pass
                await asyncio.sleep(0.3)

                state = await self._detect_page_state()
                logger.info(f"  Step {step}: Page state = {state}")

                if state == 'session_expired':
                    raise Exception("Session expired during checkout")

                elif state == 'password':
                    logger.info(f"  Step {step}: Password field detected")
                    break

                elif state in ('ach', 'ach_with_next'):
                    if not ach_selected:
                        for ctx in self._get_search_frames():
                            for sel in ['text=ACH Debit Bank', ':has-text("ACH Debit Bank")', 'span:has-text("ACH Debit")', 'label:has-text("ACH")']:
                                try:
                                    ach = await ctx.wait_for_selector(sel, timeout=1000)
                                    if ach and await ach.is_visible():
                                        await ach.click()
                                        logger.info(f"  Step {step}: Selected ACH Debit Bank")
                                        ach_selected = True
                                        await asyncio.sleep(0.2)
                                        break
                                except Exception:
                                    continue
                            if ach_selected:
                                break
                    if state == 'ach_with_next' or await self._detect_page_state() in ('next', 'ach_with_next'):
                        logger.info(f"  Step {step}: Clicking Next after ACH...")
                        await self._click_next_or_submit("Next", timeout_ms=3000)
                    continue

                elif state == 'next':
                    logger.info(f"  Step {step}: Clicking Next...")
                    if not await self._click_next_or_submit("Next", timeout_ms=3000):
                        raise Exception(f"Next detected but click failed at step {step}")
                    continue

                elif state == 'submit':
                    logger.info(f"  Step {step}: Clicking Submit...")
                    if not await self._click_next_or_submit("Submit", timeout_ms=3000):
                        raise Exception("Submit detected but click failed")
                    continue

                else:  # unknown
                    logger.warning(f"  Step {step}: Unknown page state, waiting...")
                    await asyncio.sleep(2)
                    state = await self._detect_page_state()
                    if state == 'session_expired':
                        raise Exception("Session expired during checkout")
                    if state == 'unknown':
                        await self._debug_page("checkout_stuck.png")
                        raise Exception(f"Checkout stuck at step {step}")
                    step -= 1
                    continue

            if step >= max_steps:
                raise Exception(f"Checkout exceeded {max_steps} steps")

            # Phase 2: Password confirmation
            logger.info("Entering confirmation password...")
            password = os.getenv('SITE_PASSWORD')
            if not password:
                raise Exception("SITE_PASSWORD not set")

            max_password_attempts = 3
            pw_filled = False
            for attempt in range(1, max_password_attempts + 1):
                logger.info(f"Password attempt {attempt}/{max_password_attempts}...")

                await self._scroll_to_bottom()
                await asyncio.sleep(0.2)
                await self._scroll_to_bottom()

                # Click "scroll for more" if present
                try:
                    scroll_more = await self.page.query_selector('a.ScrollForMoreLink, a[data-event="ScrollForMore"]')
                    if scroll_more and await scroll_more.is_visible():
                        await scroll_more.click()
                        await asyncio.sleep(0.2)
                except Exception:
                    pass

                pw_filled = await self._fill_password_field(password)
                if pw_filled:
                    break

                logger.warning(f"  Password attempt {attempt} failed, recovering...")
                await asyncio.sleep(1.0 * attempt)
                try:
                    await self.page.mouse.click(10, 10)
                    await asyncio.sleep(0.2)
                except Exception:
                    pass

            if not pw_filled:
                raise Exception("Could not enter password after all attempts")

            await asyncio.sleep(0.1)

            # Phase 3: Final Submit
            logger.info("Clicking final Submit after password...")
            if not await self._click_next_or_submit("Submit"):
                raise Exception("Could not find final Submit button")

            try:
                await self.page.wait_for_load_state('networkidle', timeout=10000)
            except Exception:
                pass

            logger.info("Order submitted successfully!")
            await asyncio.sleep(1)
            return True

        except Exception as e:
            logger.error(f"Error during checkout: {e}")
            raise

    async def process_multiple_items(self, items):
        """Process unfilled items: start order, add available ones, submit if found.
        Returns dict with success, items_ordered, message."""
        unfilled_items = [i for i in items if i.get('order_filled', '').lower() != 'yes']
        if not unfilled_items:
            return {'success': False, 'items_ordered': [], 'message': 'No unfilled items'}

        current_url = self.page.url
        if 'itemEntry' not in current_url:
            await self.start_order()

        items_found, total_qty_added = await self.check_and_process_items(items)

        if not items_found:
            logger.info("No items available at this time")
            return {'success': False, 'items_ordered': [], 'message': 'No items were available'}

        if total_qty_added < 10:
            logger.warning(f"Need min 10 qty (have {total_qty_added}). Not submitting.")
            for item in items_found:
                item['order_filled'] = ''
            return {'success': False, 'items_ordered': [], 'message': f'Need min 10 qty (have {total_qty_added})'}

        item_numbers = [str(i['item_number']) for i in items_found]
        logger.info(f"Found {len(items_found)} items, {total_qty_added} total qty: {', '.join(item_numbers)}")

        try:
            await self.submit_order()
            logger.info(f"ORDER SUCCESS: Placed order for items {', '.join(item_numbers)}")
            return {
                'success': True,
                'items_ordered': item_numbers,
                'message': f'Order placed for {len(items_found)} item(s): {", ".join(item_numbers)}'
            }
        except Exception as e:
            logger.error(f"ORDER FAILED: {e}")
            for item in items_found:
                item['order_filled'] = ''
            return {
                'success': False,
                'items_ordered': item_numbers,
                'message': f'Checkout failed: {e}'
            }

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


async def main():
    """Main bot loop"""
    csv_filename = 'orders.csv'

    if not Path(csv_filename).exists():
        logger.warning(f"{csv_filename} not found. Creating sample template...")
        create_sample_csv()
        logger.info(f"Please fill in orders_template.csv and rename it to {csv_filename}")
        return

    bot = WebAutomationBot(headless=False)

    try:
        logger.info("Initializing bot...")
        await bot.setup(use_saved_auth=True)

        logger.info("\n" + "=" * 60)
        logger.info("BOT STARTED - Continuously checking for items")
        logger.info("=" * 60 + "\n")

        consecutive_errors = 0
        on_item_entry = False
        while True:
            try:
                items = read_csv_file(csv_filename)
                unfilled_items = [item for item in items if item.get('order_filled', '').lower() != 'yes']

                if not unfilled_items:
                    logger.info("All items completed! Checking for new items in 5 seconds...")
                    await asyncio.sleep(5)
                    on_item_entry = False
                    continue

                if not on_item_entry or 'itemEntry' not in bot.page.url:
                    await bot.start_order()
                    on_item_entry = True

                logger.info(f"\n--- Checking {len(unfilled_items)} items for availability ---")

                items_found, total_qty_added = await bot.check_and_process_items(items)
                consecutive_errors = 0

                if items_found and total_qty_added >= 10:
                    item_numbers = [str(item['item_number']) for item in items_found]
                    logger.info(f"\n+ Found {len(items_found)} items, {total_qty_added} total qty: {', '.join(item_numbers)}")
                    await bot.submit_order()
                    update_csv_file(csv_filename, items)
                    on_item_entry = False

                    remaining = [i for i in items if i.get('order_filled', '').lower() != 'yes']
                    if not remaining:
                        logger.info("All items filled!")
                elif items_found and total_qty_added < 10:
                    logger.warning(f"Need min 10 qty total (have {total_qty_added}). Reverting - will retry.")
                    for item in items_found:
                        item['order_filled'] = ''
                else:
                    logger.info("No items available. Re-checking all items...")
                    await asyncio.sleep(1)

            except Exception as e:
                consecutive_errors += 1
                logger.error(f"Error (attempt {consecutive_errors}): {e}")

                if consecutive_errors < 3:
                    try:
                        logger.info("Attempting to recover...")
                        await bot.start_order()
                        on_item_entry = True
                        logger.info("Recovered, resuming item checks...")
                        continue
                    except Exception:
                        pass

                logger.info("Re-initializing bot (full login)...")
                consecutive_errors = 0
                on_item_entry = False
                try:
                    await bot.cleanup()
                except Exception:
                    pass
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
    print("\n" + "=" * 60)
    print("Mississippi DOR Order Bot")
    print("=" * 60)
    print("\nPress Ctrl+C to stop the bot\n")

    asyncio.run(main())
