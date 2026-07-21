"""Explore NUHS PTT form — open each dropdown to reveal options, screenshot everything."""
import asyncio
from playwright.async_api import async_playwright

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdB-N9-5o03e3m5h6fIrDb_kxQFEdcpi9h34Fe0sQUBI-9BqA/viewform"

async def click_dropdown_and_screenshot(page, name):
    dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
    await dropdown.click()
    await page.wait_for_timeout(800)
    await page.screenshot(path=f"explore_{name}_open.png")
    print(f"Screenshot: explore_{name}_open.png")
    # Get all option texts
    options = await page.locator("[role='option']").all_text_contents()
    print(f"Options: {options}")
    return options

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=400)
        page = await browser.new_page(viewport={"width": 1200, "height": 900})
        await page.goto(FORM_URL, wait_until="networkidle")

        # Page 1 — Select Rachel Teo
        dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
        await dropdown.click()
        await page.wait_for_timeout(800)
        await page.get_by_role("option", name="Rachel Teo").dispatch_event("click")
        await page.wait_for_timeout(500)
        print("Page 1: Rachel Teo selected")

        # Next to page 2
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(600)

        # Page 2 — Select NUH/AH/NTFGH IM Postings
        await page.locator("[data-value='NUH/AH/NTFGH (IM Postings)']").click()
        await page.wait_for_timeout(400)
        print("Page 2: NUH/AH/NTFGH selected")

        # Next to page 3
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(600)

        # Page 3 — Select Medicine Teaching
        dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
        await dropdown.click()
        await page.wait_for_timeout(800)
        await page.get_by_role("option", name="Medicine Teaching").dispatch_event("click")
        await page.wait_for_timeout(400)
        print("Page 3: Medicine Teaching selected")

        # Next to page 4
        await page.get_by_role("button", name="Next").click()
        await page.wait_for_load_state("networkidle")
        await page.wait_for_timeout(600)
        await page.screenshot(path="explore_page4.png")
        print("Screenshot: explore_page4.png")

        # Print all field labels on page 4
        labels = await page.locator(".M7eMe, .iLrMGb, .z12JJ, .oEqjAb").all_text_contents()
        print(f"Page 4 labels: {labels}")

        # Check for any dropdowns on page 4 and reveal their options
        dropdowns = await page.locator("[role='listbox'], [aria-haspopup='listbox']").count()
        print(f"Page 4 dropdowns: {dropdowns}")
        if dropdowns > 0:
            for i in range(dropdowns):
                d = page.locator("[role='listbox'], [aria-haspopup='listbox']").nth(i)
                await d.click()
                await page.wait_for_timeout(600)
                opts = await page.locator("[role='option']").all_text_contents()
                print(f"  Dropdown {i} options: {opts}")
                await page.screenshot(path=f"explore_page4_dropdown{i}.png")
                await page.keyboard.press("Escape")
                await page.wait_for_timeout(300)

        input("\nPress Enter to close browser...")
        await browser.close()

asyncio.run(main())
