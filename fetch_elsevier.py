"""
Full extraction of all 32 chapters from Elsevier PageBurst.
- Clicks each TOC button, pages through, saves per chapter to output directory.
- Fully robust: catches all exceptions per chapter/page, never stops early.
- Incremental saves: each chapter saved immediately after extraction.
"""
import sys, asyncio, os, time, json, traceback
from playwright.async_api import async_playwright

sys.stdout.reconfigure(line_buffering=True)

URL         = "https://pageburstls.elsevier.com/reader/books/9780443117282/epubcfi/6/2390[%3Bvnd.vst.idref%3DCH0117_2008-2039_B9780443116575001178_02]!/4[sec-s0015]/4/14/26/20[p1025]/1:482[sio%2Cn.]"
PROFILE_DIR = r"c:\Users\JasonOng\AST_Agent\.playwright_profile"
LOG_FILE    = r"c:\Users\JasonOng\AST_Agent\elsevier_log.txt"
OUT_DIR     = r"C:\Users\JasonOng\Desktop\local docs\personal\rach"
SENTINEL    = r"c:\Users\JasonOng\AST_Agent\go.txt"

os.makedirs(OUT_DIR, exist_ok=True)
for s in [SENTINEL]:
    if os.path.exists(s): os.remove(s)

def log(msg):
    print(msg, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass

CHAPTERS = [
    (5,  "Ch100_Anatomy_SmallLargeIntestine"),
    (6,  "Ch101_SmallIntestinal_Motor_Sensory"),
    (7,  "Ch102_Colonic_Motor_Sensory"),
    (8,  "Ch103_Electrolyte_Absorption_Secretion"),
    (9,  "Ch104_Digestion_Absorption_Macro"),
    (10, "Ch105_Micronutrients_Absorption"),
    (11, "Ch106_Maldigestion_Malabsorption"),
    (12, "Ch107_Small_Intestinal_Bacterial_Overgrowth"),
    (13, "Ch108_Short_Bowel_Syndrome"),
    (14, "Ch109_Celiac_Disease"),
    (15, "Ch110_Tropical_Diarrhea_Malabsorption"),
    (16, "Ch111_Whipple_Disease"),
    (17, "Ch112_Infectious_Enteritis_Proctocolitis"),
    (18, "Ch113_Food_Poisoning"),
    (19, "Ch114_Cdiff_AAD"),
    (20, "Ch115_Intestinal_Protozoa"),
    (21, "Ch116_Intestinal_Worms"),
    (22, "Ch117_IBD_Epidemiology_Diagnosis"),
    (28, "Ch118_IBD_Management"),
    (29, "Ch119_Ileostomies_Colostomies"),
    (30, "Ch120_Intestinal_Ischemia"),
    (31, "Ch121_Intestinal_Ulcerations"),
    (32, "Ch122_Appendicitis"),
    (33, "Ch123_Diverticular_Disease"),
    (34, "Ch124_IBS"),
    (35, "Ch125_Intestinal_Obstruction"),
    (36, "Ch126_Ileus_PseudoObstruction"),
    (37, "Ch127_Small_Bowel_Tumors"),
    (38, "Ch128_Colonic_Polyps_Polyposis"),
    (39, "Ch129_Colorectal_Cancer"),
    (40, "Ch130_Other_Colon_Diseases"),
    (41, "Ch131_Anal_Diseases"),
]

async def get_content_from_frames(page):
    best = ""
    for frame in page.frames:
        try:
            text = await frame.evaluate("""() => {
                const selectors = ['[class*="page"]','[class*="content"]','[class*="chapter"]',
                                   '[class*="text"]','article','section','main'];
                for (const sel of selectors) {
                    let combined = '';
                    for (const el of document.querySelectorAll(sel)) {
                        const t = el.innerText ? el.innerText.trim() : '';
                        if (t.length > 100) combined += t + '\\n\\n';
                    }
                    if (combined.length > 300) return combined.trim();
                }
                return document.body ? document.body.innerText.trim() : '';
            }""")
            if text and len(text) > len(best):
                best = text
        except Exception:
            pass
    return best

async def click_toc_button(page, btn_index):
    try:
        clicked = await page.evaluate(f"""() => {{
            const btns = document.querySelectorAll('button.sc-eBHJIF');
            if (btns[{btn_index}]) {{ btns[{btn_index}].click(); return true; }}
            return false;
        }}""")
        return clicked
    except Exception as e:
        log(f"  TOC click error: {e}")
        return False

async def click_next(page):
    for sel in [
        'button[aria-label*="Next"]', 'button[aria-label*="next"]',
        '[title="Next page"]', '[title="Next"]',
        'button:has-text("Next")',
    ]:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=1200):
                await btn.click()
                return True
        except Exception:
            pass
    try:
        await page.keyboard.press("ArrowRight")
        return True
    except Exception:
        pass
    return False

async def get_page_indicator(page):
    try:
        return await page.evaluate("""() => {
            const candidates = [
                document.querySelector('[class*="pageNum"]'),
                document.querySelector('[class*="page-num"]'),
                document.querySelector('input[aria-label*="page"]'),
                document.querySelector('[class*="pagination"] input'),
            ];
            for (const el of candidates) {
                if (el) return (el.value || el.innerText || '').trim();
            }
            return '';
        }""")
    except Exception:
        return "?"

async def extract_chapter(page, btn_index, chapter_name):
    log(f"\n{'='*55}")
    log(f"  {chapter_name}  (TOC btn #{btn_index})")
    log(f"{'='*55}")

    # Click the TOC entry to navigate to chapter start
    try:
        clicked = await click_toc_button(page, btn_index)
        if not clicked:
            log("  WARNING: TOC button not found — skipping chapter")
            return []
    except Exception as e:
        log(f"  TOC click failed: {e}")
        return []

    await page.wait_for_timeout(4500)

    pages_text = []
    prev_content = ""
    same_count   = 0
    empty_count  = 0
    max_pages    = 80  # safety cap per chapter

    for pg in range(1, max_pages + 1):
        try:
            content = await get_content_from_frames(page)
        except Exception as e:
            log(f"  Page {pg}: extraction error — {e}")
            content = ""

        try:
            pg_ind = await get_page_indicator(page)
        except Exception:
            pg_ind = "?"

        if content and len(content) > 150:
            if content == prev_content:
                same_count += 1
                log(f"  Page {pg} (pg {pg_ind}): duplicate #{same_count}")
                if same_count >= 3:
                    log("  3 identical pages in a row — chapter complete.")
                    break
            else:
                same_count  = 0
                empty_count = 0
                pages_text.append(f"--- Reader page {pg_ind} ---\n{content}")
                log(f"  Page {pg} (pg {pg_ind}): {len(content)} chars  [total: {len(pages_text)} pages]")
            prev_content = content
        else:
            empty_count += 1
            log(f"  Page {pg} (pg {pg_ind}): empty/short ({len(content)} chars) [{empty_count}]")
            if empty_count >= 5:
                log("  5 empty pages — chapter complete.")
                break

        # Attempt to go to next page
        try:
            clicked = await click_next(page)
            if not clicked:
                log(f"  No Next button at page {pg} — chapter complete.")
                break
        except Exception as e:
            log(f"  Next click error: {e}")
            break

        await page.wait_for_timeout(2500)

    log(f"  -> Chapter done: {len(pages_text)} pages extracted.")
    return pages_text


async def main():
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write("=== fetch_elsevier.py — full 32-chapter extraction ===\n")
    log(f"Output dir : {OUT_DIR}")
    log(f"Profile    : {PROFILE_DIR}")
    log(f"Chapters   : {len(CHAPTERS)}")

    async with async_playwright() as p:
        context = await p.chromium.launch_persistent_context(
            PROFILE_DIR, headless=False, slow_mo=30, channel="chrome"
        )
        page = context.pages[0] if context.pages else await context.new_page()

        # Load reader
        log("\nLoading reader...")
        try:
            await page.goto(URL, wait_until="domcontentloaded", timeout=30000)
        except Exception as e:
            log(f"  goto error: {e}")

        for i in range(20):
            await page.wait_for_timeout(2000)
            try:
                txt = await page.evaluate("() => document.body ? document.body.innerText.trim() : ''")
            except Exception:
                txt = ""
            log(f"  {(i+1)*2}s: {len(txt)} chars")
            if len(txt) > 500:
                log("  Reader ready.\n")
                break

        log("Waiting for start signal (go.txt)...")
        while not os.path.exists(SENTINEL):
            time.sleep(2)
        os.remove(SENTINEL)
        log("START signal received — beginning extraction of all 32 chapters.\n")

        # Progress tracker
        progress_file = os.path.join(OUT_DIR, "00_PROGRESS.txt")

        completed = []
        failed    = []

        for i, (btn_idx, chapter_name) in enumerate(CHAPTERS):
            log(f"\n[{i+1}/{len(CHAPTERS)}] Starting: {chapter_name}")

            try:
                pages = await extract_chapter(page, btn_idx, chapter_name)
            except Exception as e:
                log(f"  FATAL ERROR in chapter {chapter_name}: {e}")
                log(traceback.format_exc())
                pages = []
                failed.append(chapter_name)

            # Save chapter file regardless of content
            out_file = os.path.join(OUT_DIR, f"{chapter_name}.txt")
            try:
                with open(out_file, "w", encoding="utf-8") as f:
                    f.write(f"CHAPTER: {chapter_name}\n")
                    f.write(f"Pages extracted: {len(pages)}\n")
                    f.write("=" * 60 + "\n\n")
                    for j, pg_text in enumerate(pages):
                        f.write(f"\n{'—'*40}\nPage {j+1}\n{'—'*40}\n{pg_text}\n")
                log(f"  Saved -> {out_file}")
                completed.append((chapter_name, len(pages)))
            except Exception as e:
                log(f"  Save error: {e}")
                failed.append(chapter_name)

            # Update progress file after every chapter
            try:
                with open(progress_file, "w", encoding="utf-8") as f:
                    f.write(f"Extraction progress — {i+1}/{len(CHAPTERS)} chapters done\n\n")
                    f.write("COMPLETED:\n")
                    for ch, pgs in completed:
                        f.write(f"  [OK] {ch}  ({pgs} pages)\n")
                    if failed:
                        f.write("\nFAILED:\n")
                        for ch in failed:
                            f.write(f"  [FAIL] {ch}\n")
                    remaining = [c[1] for c in CHAPTERS[i+1:]]
                    if remaining:
                        f.write("\nREMAINING:\n")
                        for ch in remaining:
                            f.write(f"  [ ] {ch}\n")
            except Exception:
                pass

        log("\n\n========================================")
        log(f"ALL DONE — {len(completed)}/{len(CHAPTERS)} chapters extracted")
        log(f"Output: {OUT_DIR}")
        for ch, pgs in completed:
            log(f"  [OK] {ch}: {pgs} pages")
        if failed:
            log("FAILED chapters:")
            for ch in failed:
                log(f"  [FAIL] {ch}")
        log("========================================")

        await context.close()

asyncio.run(main())
