"""
Run once to log in to Elsevier PageBurst and save the session.
After this, fetch_elsevier.py will reuse the saved session automatically.
"""
import sys
import asyncio
import os
import time
from playwright.async_api import async_playwright

sys.stdout.reconfigure(line_buffering=True)

PROFILE_DIR = r"c:\Users\JasonOng\AST_Agent\.playwright_profile"
LOGIN_URL   = "https://pageburstls.elsevier.com"
SENTINEL    = r"c:\Users\JasonOng\AST_Agent\go.txt"
LOG_FILE    = r"c:\Users\JasonOng\AST_Agent\elsevier_log.txt"

def log(msg):
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

if os.path.exists(SENTINEL):
    os.remove(SENTINEL)

os.makedirs(PROFILE_DIR, exist_ok=True)

async def main():
    log("=== elsevier_login.py — one-time session setup ===")

    async with async_playwright() as p:
        # Persistent context saves cookies/storage to PROFILE_DIR
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR,
            headless=False,
            slow_mo=50,
            channel="chrome",
        )

        page = context.pages[0] if context.pages else await context.new_page()

        log(f"Opening login page...")
        await page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=30000)
        log("Login page loaded.")
        log("")
        log(">>> BROWSER IS READY. Enter your username and password, then click Sign In.")
        log(">>> Once you are on the book page, tell Claude 'ready'.")
        log("")
        log(f"Waiting for signal ({SENTINEL})...")

        while not os.path.exists(SENTINEL):
            time.sleep(2)
        os.remove(SENTINEL)
        log("Signal received. Verifying session...")

        # Check we are actually logged in (not on login page)
        current_url = page.url
        log(f"Current URL: {current_url}")

        if "login" in current_url.lower() or "signin" in current_url.lower():
            log("WARNING: Still on login page — session may not have been saved correctly.")
        else:
            log("Session verified. Cookies saved to profile directory.")

        # Take screenshot as confirmation
        await page.screenshot(path=r"c:\Users\JasonOng\AST_Agent\elsevier_session_confirmed.png")
        log("Screenshot saved: elsevier_session_confirmed.png")
        log("")
        log("Session saved! Future runs will open already logged in.")

        await context.close()

asyncio.run(main())
