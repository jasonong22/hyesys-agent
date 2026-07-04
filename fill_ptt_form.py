"""Fill NUHS PTT form for Rachel Teo — 6 entries from Telegram, approval before each submit."""
import asyncio
import ctypes
from playwright.async_api import async_playwright

MB_YESNOCANCEL = 0x03
IDYES = 6
IDNO  = 7

def ask_submit(i, total, entry):
    msg = (
        f"Entry {i}/{total}\n\n"
        f"Date:  {entry['date_day']}/{entry['date_month']}/{entry['date_year']}\n"
        f"Time:  {entry['hour']}:{entry['minute']}\n"
        f"Title: {entry['title']}\n"
        f"Tutor: {entry['tutor']}\n\n"
        f"Submit this entry?"
    )
    result = ctypes.windll.user32.MessageBoxW(0, msg, f"PTT Form — Entry {i}/{total}", MB_YESNOCANCEL)
    return result == IDYES

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSdB-N9-5o03e3m5h6fIrDb_kxQFEdcpi9h34Fe0sQUBI-9BqA/viewform"

COMMON = {
    "name": "Rachel Teo",
    "posting": "NUH/AH/NTFGH (IM Postings)",
    "type_of_teaching": "Medicine Teaching",
    "duration": "1",
}

ENTRIES = [
    {"date_day": "22", "date_month": "06", "date_year": "2026", "hour": "12", "minute": "00",
     "title": "Rehab CME", "tutor": "Rehab Dept"},
    {"date_day": "22", "date_month": "06", "date_year": "2026", "hour": "14", "minute": "00",
     "title": "Gastro M&M and teaching", "tutor": "Gastro Dept"},
    {"date_day": "25", "date_month": "06", "date_year": "2026", "hour": "08", "minute": "00",
     "title": "Respi GWR", "tutor": "Respi 53 team"},
    {"date_day": "29", "date_month": "06", "date_year": "2026", "hour": "14", "minute": "00",
     "title": "G&H teaching on resistant nausea and vomiting", "tutor": "G&H Dept"},
    {"date_day": "01", "date_month": "07", "date_year": "2026", "hour": "12", "minute": "30",
     "title": "Neuro CME meeting", "tutor": "Neuro Dept"},
    {"date_day": "02", "date_month": "07", "date_year": "2026", "hour": "08", "minute": "00",
     "title": "Respi GWR", "tutor": "MICU team A"},
]


async def fill_entry(page, entry):
    a = {**COMMON, **entry}

    await page.goto(FORM_URL, wait_until="networkidle")
    await page.wait_for_timeout(600)

    # Page 1 — Name
    dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
    await dropdown.click()
    await page.wait_for_timeout(800)
    await page.get_by_role("option", name=a["name"]).dispatch_event("click")
    await page.wait_for_timeout(400)
    print(f"  [OK] Name: {a['name']}")
    await page.get_by_role("button", name="Next").click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(500)

    # Page 2 — Posting
    await page.locator(f"[data-value='{a['posting']}']").click()
    await page.wait_for_timeout(300)
    print(f"  [OK] Posting: {a['posting']}")
    await page.get_by_role("button", name="Next").click()
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(500)

    # Page 3 — Type of teaching
    dropdown = page.locator("[role='listbox'], [aria-haspopup='listbox']").first
    await dropdown.click()
    await page.wait_for_timeout(800)
    await page.get_by_role("option", name=a["type_of_teaching"]).dispatch_event("click")
    await page.wait_for_timeout(300)
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(300)
    print(f"  [OK] Type: {a['type_of_teaching']}")
    await page.get_by_role("button", name="Next").click(force=True)
    await page.wait_for_load_state("networkidle")
    await page.wait_for_timeout(500)

    # Page 4 — Details
    date_iso  = f"{a['date_year']}-{a['date_month']}-{a['date_day']}"
    day_int   = str(int(a["date_day"]))
    month_int = str(int(a["date_month"]))
    await page.evaluate(f"""() => {{
        const input = document.querySelector("input[type='date']");
        if (!input) return;
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(input, '{date_iso}');
        ['input', 'change', 'blur'].forEach(e =>
            input.dispatchEvent(new Event(e, {{bubbles: true}})));
        const y = document.querySelector("input[name='entry.372757847_year']");
        const m = document.querySelector("input[name='entry.372757847_month']");
        const d = document.querySelector("input[name='entry.372757847_day']");
        if (y) y.value = '{a['date_year']}';
        if (m) m.value = '{month_int}';
        if (d) d.value = '{day_int}';
    }}""")
    await page.wait_for_timeout(400)
    print(f"  [OK] Date: {a['date_day']}/{a['date_month']}/{a['date_year']}")

    await page.get_by_label("Hour").fill(a["hour"])
    await page.get_by_label("Minute").fill(a["minute"])
    print(f"  [OK] Time: {a['hour']}:{a['minute']}")

    text_inputs = page.locator("input[type='text']")
    await text_inputs.nth(2).fill(a["title"])
    print(f"  [OK] Title: {a['title']}")
    await text_inputs.nth(3).fill(a["tutor"])
    print(f"  [OK] Tutor: {a['tutor']}")

    await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    await page.wait_for_timeout(500)

    duration_option = page.locator(f"[data-value='{a['duration']}']").last
    await duration_option.click()
    await page.wait_for_timeout(300)
    print(f"  [OK] Duration: {a['duration']} hr")

    # Scroll back to top so full form is visible for review
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(400)


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=400)
        page = await browser.new_page(viewport={"width": 1200, "height": 900})

        for i, entry in enumerate(ENTRIES, 1):
            print(f"\n{'='*55}")
            print(f"  Entry {i}/{len(ENTRIES)}: {entry['date_day']}/{entry['date_month']} "
                  f"{entry['hour']}:{entry['minute']}  |  {entry['title']}  |  {entry['tutor']}")
            print(f"{'='*55}")

            await fill_entry(page, entry)

            print("\n  >> Browser is open — check the popup to approve or skip.")
            approved = ask_submit(i, len(ENTRIES), entry)

            if not approved:
                print(f"  [SKIPPED] Entry {i}")
                continue

            await page.get_by_role("button", name="Submit").click()
            await page.wait_for_timeout(3000)
            print(f"  [SUBMITTED] Entry {i} done.")
            await page.wait_for_timeout(1000)

        print("\n All entries processed. Closing browser.")
        await browser.close()

asyncio.run(main())
