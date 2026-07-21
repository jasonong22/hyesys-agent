"""Inspect page 4 DOM to find correct selectors for date and time fields."""
import asyncio
from playwright.async_api import async_playwright

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdB-N9-5o03e3m5h6fIrDb_kxQFEdcpi9h34Fe0sQUBI-9BqA/viewform"

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=400)
        page = await browser.new_page(viewport={"width": 1200, "height": 900})
        await page.goto(FORM_URL, wait_until="networkidle")

        # Page 1
        dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
        await dropdown.click()
        await page.wait_for_timeout(800)
        await page.get_by_role("option", name="Rachel Teo").dispatch_event("click")
        await page.wait_for_timeout(400)
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(500)

        # Page 2
        await page.locator("[data-value='NUH/AH/NTFGH (IM Postings)']").click()
        await page.wait_for_timeout(300)
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(500)

        # Page 3
        dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
        await dropdown.click()
        await page.wait_for_timeout(800)
        await page.get_by_role("option", name="Medicine Teaching").dispatch_event("click")
        await page.wait_for_timeout(300)
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(500)

        # Page 4 — dump all inputs
        print("=== ALL INPUTS ON PAGE 4 ===")
        all_inputs = await page.evaluate("""() => {
            const inputs = document.querySelectorAll('input, textarea, [role=textbox], [contenteditable=true]');
            return Array.from(inputs).map(el => ({
                tag: el.tagName,
                type: el.type || 'n/a',
                name: el.name || 'n/a',
                id: el.id || 'n/a',
                placeholder: el.placeholder || 'n/a',
                ariaLabel: el.getAttribute('aria-label') || 'n/a',
                jsname: el.getAttribute('jsname') || 'n/a',
                class: el.className.substring(0, 60)
            }));
        }""")
        for i, inp in enumerate(all_inputs):
            print(f"  [{i}] {inp}")

        await page.screenshot(path="inspect_page4.png")
        print("Screenshot: inspect_page4.png")
        await browser.close()

asyncio.run(main())
