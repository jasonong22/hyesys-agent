"""
Build PCS 测试清单 testing checklist - Master v2.xlsx
Adds 6 upstream meter columns (KW A/B/C + KVAR A/B/C) per PCS block.

Final column layout (50 cols):
  A-C  (1-3)  : metadata
  D-I  (4-9)  : command reference (KW A/B/C, KVAR A/B/C)
  J    (10)   : PCS-1  Within Range?
  K-P  (11-16): PCS-1  readings
  Q-V  (17-22): Upstream Meter 1  readings  [NEW]
  W    (23)   : PCS-2  Within Range?
  X-AC (24-29): PCS-2  readings
  AD-AI(30-35): Upstream Meter 2  readings  [NEW]
  AJ   (36)   : PCS-3  Within Range?
  AK-AP(37-42): PCS-3  readings
  AQ-AV(43-48): Upstream Meter 3  readings  [NEW]
  AW   (49)   : Note
  AX   (50)   : Reminder
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
import openpyxl.formatting.formatting

FAT_DIR = Path(
    r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD"
    r"\HyESys Dept\3. Hardware (PCS.BATT)\PCS"
    r"\sinosoar - current v2.1\FAT - PCS checklist"
)
SRC = FAT_DIR / "PCS 测试清单 testing checklist - Master.xlsx"
DST = FAT_DIR / "PCS 测试清单 testing checklist - Master v2.xlsx"

DATA_START = 14
DATA_END   = 85

# ── Fills (8-char AARRGGBB) ───────────────────────────────────────────────────
AMBER        = PatternFill("solid", fgColor="FFFFC000")
GREEN_INPUT  = PatternFill("solid", fgColor="FF92D050")
UM_HDR_FILL  = PatternFill("solid", fgColor="FF7030A0")   # purple
UM_SUB1_FILL = PatternFill("solid", fgColor="FFB09CC8")   # medium purple
UM_SUB2_FILL = PatternFill("solid", fgColor="FFDDD0EA")   # light purple
GREEN_CF     = PatternFill("solid", fgColor="FF00B050")
RED_CF       = PatternFill("solid", fgColor="FFFF0000")

WHITE_BOLD = Font(bold=True, color="FFFFFF")
DARK_BOLD  = Font(bold=True)
CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER_NW  = Alignment(horizontal="center", vertical="center")


def sc(ws, row, col, value=None, fill=None, font=None, align=None):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if fill is not None:
        c.fill = fill
    if font is not None:
        c.font = font
    if align is not None:
        c.alignment = align
    return c


def safe_merge(ws, r1, c1, r2, c2, value, fill, font, align=CENTER):
    """Merge cells, removing any overlapping existing merges first."""
    # Collect and remove any existing merged ranges that overlap our target
    to_remove = []
    for mr in list(ws.merged_cells.ranges):
        # Check overlap
        if (mr.min_row <= r2 and mr.max_row >= r1 and
                mr.min_col <= c2 and mr.max_col >= c1):
            to_remove.append(str(mr))
    for r in to_remove:
        try:
            ws.unmerge_cells(r)
        except Exception:
            pass
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1)
    if value is not None:
        cell.value = value
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if align is not None:
        cell.alignment = align


def add_um_section(ws, meter_idx, col_start):
    """Add all headers for one upstream meter block (rows 4-12 + data fills)."""
    c = col_start

    # ── Rows 4-9: label + topic placeholder ──
    safe_merge(ws, 4, c, 4, c+5,
               f"上游电表 {meter_idx}\nUpstream Meter {meter_idx}\nInsert MQTT Topic",
               UM_HDR_FILL, WHITE_BOLD)
    for row in range(5, 10):
        safe_merge(ws, row, c, row, c+5, "", UM_HDR_FILL, WHITE_BOLD)

    # ── Row 10: section header ──
    safe_merge(ws, 10, c, 10, c+5,
               f"上游电表 {meter_idx} Readings\nUpstream Meter {meter_idx} Readings",
               UM_HDR_FILL, WHITE_BOLD)

    # ── Row 11: KW (cols c–c+2) | KVAR (cols c+3–c+5) ──
    safe_merge(ws, 11, c,   11, c+2, "有功KW\nActive Power KW",       UM_SUB1_FILL, DARK_BOLD)
    safe_merge(ws, 11, c+3, 11, c+5, "无功KVAR\nReactive Power KVAR", UM_SUB1_FILL, DARK_BOLD)

    # ── Row 12: A/B/C × 2 ──
    for i, letter in enumerate("ABCABC"):
        # Each is a standalone cell — no merge needed
        sc(ws, 12, c+i, letter, UM_SUB2_FILL, DARK_BOLD, CENTER)

    # ── Data rows 14-85: light-purple fill ──
    for row in range(DATA_START, DATA_END + 1):
        for offset in range(6):
            sc(ws, row, c+offset, fill=UM_SUB2_FILL, align=CENTER_NW)


def cf_within_range_green(read_start, cmd_start=4, row=DATA_START):
    rs = [get_column_letter(read_start + i) for i in range(6)]
    cs = [get_column_letter(cmd_start  + i) for i in range(6)]
    or_part = ",".join(f"ISNUMBER({r}{row})" for r in rs)
    ifs = ",".join(
        f"IF(ISNUMBER({cs[i]}{row}),AND(ISNUMBER({rs[i]}{row}),ABS({cs[i]}{row}-{rs[i]}{row})<=2),TRUE)"
        for i in range(6)
    )
    return f"AND(OR({or_part}),{ifs})"

def cf_within_range_red(read_start, row=DATA_START):
    rs = [get_column_letter(read_start + i) for i in range(6)]
    return f"OR({','.join(f'ISNUMBER({r}{row})' for r in rs)})"

def cf_reading_red(read_col, cmd_col, row=DATA_START):
    r, c = get_column_letter(read_col), get_column_letter(cmd_col)
    return f"AND(ISNUMBER({r}{row}),ISNUMBER({c}{row}),ABS({c}{row}-{r}{row})>2)"

def cf_um_green(read_col, cmd_col, row=DATA_START):
    r, c = get_column_letter(read_col), get_column_letter(cmd_col)
    return f"AND(ISNUMBER({r}{row}),ISNUMBER({c}{row}),ABS({c}{row}-{r}{row})<=2)"

def cf_um_red(read_col, cmd_col, row=DATA_START):
    r, c = get_column_letter(read_col), get_column_letter(cmd_col)
    return f"AND(ISNUMBER({r}{row}),ISNUMBER({c}{row}),ABS({c}{row}-{r}{row})>2)"

def rng(col_a, col_b=None):
    ca = get_column_letter(col_a) if isinstance(col_a, int) else col_a
    if col_b is None:
        return f"{ca}{DATA_START}:{ca}{DATA_END}"
    cb = get_column_letter(col_b) if isinstance(col_b, int) else col_b
    return f"{ca}{DATA_START}:{cb}{DATA_END}"


# ── MAIN ──────────────────────────────────────────────────────────────────────
print(f"Loading: {SRC.name}")
wb = openpyxl.load_workbook(SRC)
ws = wb.active

# Step 1: Insert upstream meter columns (right → left to keep indices stable)
ws.insert_cols(31, 6)   # UM-3: after original AD (col 30), before old AE (31)
ws.insert_cols(24, 6)   # UM-2: after original W  (col 23), before old X  (24)
ws.insert_cols(17, 6)   # UM-1: after original P  (col 16), before old Q  (17)
print(f"  Inserted 18 columns. Total: {ws.max_column} ({get_column_letter(ws.max_column)})")

# Step 2: Column widths for the 18 new upstream meter columns
for col in [*range(17, 23), *range(30, 36), *range(43, 49)]:
    ws.column_dimensions[get_column_letter(col)].width = 6.5

# Step 3: Ensure PCS Within Range? cells get the right fill for data rows
# (auto-shift should preserve fills; but re-apply to be safe)
for row in range(DATA_START, DATA_END + 1):
    sc(ws, row, 10, fill=AMBER,       align=CENTER_NW)   # J = PCS-1
    sc(ws, row, 23, fill=AMBER,       align=CENTER_NW)   # W = PCS-2
    sc(ws, row, 36, fill=AMBER,       align=CENTER_NW)   # AJ = PCS-3
    for offset in range(6):
        sc(ws, row, 11+offset, fill=GREEN_INPUT, align=CENTER_NW)  # K-P PCS-1
        sc(ws, row, 24+offset, fill=GREEN_INPUT, align=CENTER_NW)  # X-AC PCS-2
        sc(ws, row, 37+offset, fill=GREEN_INPUT, align=CENTER_NW)  # AK-AP PCS-3

# Step 4: Add the 3 upstream meter sections (headers + data fills)
add_um_section(ws, 1, 17)   # Q-V  (17-22)
add_um_section(ws, 2, 30)   # AD-AI (30-35)
add_um_section(ws, 3, 43)   # AQ-AV (43-48)

# Step 5: Verify / ensure Within Range? row 10-12 merges exist and have correct values
# (auto-shift moves J10:J12 → stays, Q10:Q12 → W10:W12, X10:X12 → AJ10:AJ12)
for (wc, label) in [(10, "J"), (23, "W"), (36, "AJ")]:
    safe_merge(ws, 10, wc, 12, wc, "Within\nRange?", AMBER, DARK_BOLD)

# Step 6: Rebuild all conditional formatting
ws.conditional_formatting = openpyxl.formatting.formatting.ConditionalFormattingList()

# PCS-1 Within Range? J (col 10) reads K-P (11-16)
ws.conditional_formatting.add(rng(10), FormulaRule([cf_within_range_green(11)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(10), FormulaRule([cf_within_range_red(11)],   fill=RED_CF,   stopIfTrue=True))
# PCS-1 readings K-P (11-16): RED if out of range
ws.conditional_formatting.add(rng(11, 16), FormulaRule([cf_reading_red(11, 4)], fill=RED_CF, stopIfTrue=True))
# Upstream Meter 1 Q-V (17-22): GREEN/RED per cell
ws.conditional_formatting.add(rng(17, 22), FormulaRule([cf_um_green(17, 4)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(17, 22), FormulaRule([cf_um_red(17, 4)],   fill=RED_CF,   stopIfTrue=True))

# PCS-2 Within Range? W (col 23) reads X-AC (24-29)
ws.conditional_formatting.add(rng(23), FormulaRule([cf_within_range_green(24)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(23), FormulaRule([cf_within_range_red(24)],   fill=RED_CF,   stopIfTrue=True))
# PCS-2 readings X-AC (24-29): RED if out of range
ws.conditional_formatting.add(rng(24, 29), FormulaRule([cf_reading_red(24, 4)], fill=RED_CF, stopIfTrue=True))
# Upstream Meter 2 AD-AI (30-35): GREEN/RED per cell
ws.conditional_formatting.add(rng(30, 35), FormulaRule([cf_um_green(30, 4)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(30, 35), FormulaRule([cf_um_red(30, 4)],   fill=RED_CF,   stopIfTrue=True))

# PCS-3 Within Range? AJ (col 36) reads AK-AP (37-42)
ws.conditional_formatting.add(rng(36), FormulaRule([cf_within_range_green(37)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(36), FormulaRule([cf_within_range_red(37)],   fill=RED_CF,   stopIfTrue=True))
# PCS-3 readings AK-AP (37-42): RED if out of range
ws.conditional_formatting.add(rng(37, 42), FormulaRule([cf_reading_red(37, 4)], fill=RED_CF, stopIfTrue=True))
# Upstream Meter 3 AQ-AV (43-48): GREEN/RED per cell
ws.conditional_formatting.add(rng(43, 48), FormulaRule([cf_um_green(43, 4)], fill=GREEN_CF, stopIfTrue=True))
ws.conditional_formatting.add(rng(43, 48), FormulaRule([cf_um_red(43, 4)],   fill=RED_CF,   stopIfTrue=True))

# Step 7: Save
wb.save(DST)
print(f"Saved: {DST.name}")
print()
print("Column layout:")
print("  J(10)=PCS-1 Within Range?  K-P(11-16)=PCS-1  Q-V(17-22)=UM-1")
print("  W(23)=PCS-2 Within Range?  X-AC(24-29)=PCS-2  AD-AI(30-35)=UM-2")
print("  AJ(36)=PCS-3 Within Range?  AK-AP(37-42)=PCS-3  AQ-AV(43-48)=UM-3")
print("  AW(49)=Note  AX(50)=Reminder")
