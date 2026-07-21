"""
Diagnostic: dumps sidebar HTML and all clickable elements to understand the reader structure.
"""
import sys, asyncio, os, time, json
from playwright.async_api import async_playwright

sys.stdout.reconfigure(line_buffering=True)

URL         = "https://pageburstls.elsevier.com/reader/books/9780443117282/epubcfi/6/2390[%3Bvnd.vst.idref%3DCH0117_2008-2039_B9780443116575001178_02]!/4[sec-s0015]/4/14/26/20[p1025]/1:482[sio%2Cn.]"
PROFILE_DIR = r"c:\Users\JasonOng\AST_Agent\.playwright_profile"
LOG_FILE    = r"c:\Users\JasonOng\AST_Agent\elsevier_log.txt"

def log(msg):
    print(msg, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

async def main():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("=== elsevier_inspect.py ===\n")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, slow_mo=50, channel="chrome"
        )
        page = context.pages[0] if context.pages else await context.new_page()

        log("Loading reader...")
        await page.goto(URL, wait_until="domcontentloaded", timeout=30000)

        # Wait for content to appear
        for i in range(20):
            await page.wait_for_timeout(2000)
            txt = await page.evaluate("() => document.body.innerText.trim()")
            log(f"  {(i+1)*2}s: {len(txt)} chars")
            if len(txt) > 500:
                log("  Loaded.")
                break

        # Take full screenshot
        await page.screenshot(path=r"c:\Users\JasonOng\AST_Agent\elsevier_full.png", full_page=True)
        log("Full screenshot saved.")

        # Dump body inner HTML (first 50KB)
        html = await page.evaluate("() => document.body.innerHTML.substring(0, 50000)")
        with open(r"c:\Users\JasonOng\AST_Agent\elsevier_body.html", "w", encoding="utf-8") as f:
            f.write(html)
        log("Body HTML (first 50KB) saved -> elsevier_body.html")

        # Get all clickable/interactive elements with their text
        clickables = await page.evaluate("""() => {
            const results = [];
            const els = document.querySelectorAll(
                'button, [role="button"], [role="link"], [role="treeitem"], [role="menuitem"], [tabindex], [onclick], [class*="toc"], [class*="chapter"], [class*="sidebar"]'
            );
            els.forEach(el => {
                const text = (el.innerText || el.textContent || '').trim().substring(0, 100);
                const role = el.getAttribute('role') || el.tagName.toLowerCase();
                const cls = typeof el.className === 'string' ? el.className.substring(0, 80) : '';
                const tabindex = el.getAttribute('tabindex') || '';
                if (text.length > 0) {
                    results.push({ text, role, cls, tabindex });
                }
            });
            return results;
        }""")
        with open(r"c:\Users\JasonOng\AST_Agent\elsevier_clickables.json", "w", encoding="utf-8") as f:
            json.dump(clickables, f, indent=2, ensure_ascii=False)
        log(f"Clickable elements ({len(clickables)}) saved -> elsevier_clickables.json")

        # Also dump all text nodes to see structure
        body_text = await page.evaluate("() => document.body.innerText")
        with open(r"c:\Users\JasonOng\AST_Agent\elsevier_bodytext.txt", "w", encoding="utf-8") as f:
            f.write(body_text)
        log(f"Body text ({len(body_text)} chars) saved -> elsevier_bodytext.txt")

        log("\nDone. You can now close the browser.")
        await page.wait_for_timeout(5000)
        await context.close()

asyncio.run(main())
