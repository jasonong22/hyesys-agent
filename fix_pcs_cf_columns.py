"""
fix_pcs_cf_columns.py
=====================
For each of the 3 PCS FAT Excel files:
  1. Delete columns J (10), W (23), AJ (36) — right-to-left to avoid index shift.
  2. Clear ALL existing conditional formatting rules.
  3. Add new per-column GREEN / RED CF rules for the 3 result blocks
     (6 cols each × 3 sites = 18 columns), rows 13:200.
  4. Save the file.

New column positions after deletion:
  Site 1 results : cols 10-15  (J:O)   — cmd cols D:I (4-9)
  Site 2 results : cols 22-27  (V:AA)  — cmd cols D:I (4-9)
  Site 3 results : cols 34-39  (AH:AM) — cmd cols D:I (4-9)

GREEN: =AND(<res_col><row><>"", ABS(<res_col><row>-$<cmd_col><row>)<=5)
RED  : =AND(<res_col><row><>"", ABS(<res_col><row>-$<cmd_col><row>)>5)
"""

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

# ── file list ──────────────────────────────────────────────────────────────────
BASE = (r"C:\Users\JasonOng\OneDrive - Advancer Global Ltd\AST BD\HyESys Dept"
        r"\3. Hardware (PCS.BATT)\PCS\sinosoar - current v2.1\FAT - PCS checklist")
TESTED = BASE + r"\tested files"

FILES = {
    "Master": BASE + r"\PCS 测试清单 testing checklist - Master v2.xlsx",
    "Run 1":  TESTED + r"\PCS FAT Multi 20261014BY001+20261014BY002+26022703840003 2026-10-23 11-56.xlsx",
    "Run 2":  TESTED + r"\PCS FAT Multi 20261014BY001+20261014BY002+26022703840003 2026-10-23 12-49.xlsx",
}

# ── constants ─────────────────────────────────────────────────────────────────
# columns to delete (1-indexed), delete RIGHT to LEFT
DEL_COLS = [36, 23, 10]   # AJ, W, J

# After deletion, result block start columns (1-indexed)
RESULT_BLOCKS = [10, 22, 34]   # Site 1 = cols 10-15, Site 2 = 22-27, Site 3 = 34-39

# Command columns (1-indexed) — UNCHANGED (D:I = 4-9)
CMD_COLS = [4, 5, 6, 7, 8, 9]   # D, E, F, G, H, I  → offsets 0-5

# CF row range
CF_ROW_START = 13
CF_ROW_END   = 200

# Fill colours
GREEN_HEX = "00B050"
RED_HEX   = "FF0000"

# ── helpers ───────────────────────────────────────────────────────────────────
def make_fill(hex_colour: str) -> PatternFill:
    return PatternFill(start_color=hex_colour, end_color=hex_colour, fill_type="solid")

GREEN_FILL = make_fill(GREEN_HEX)
RED_FILL   = make_fill(RED_HEX)


def print_existing_cf(ws, label: str):
    """Print a summary of existing CF rules before we wipe them."""
    print(f"\n  [{label}] Existing CF rule sets: {len(ws.conditional_formatting._cf_rules)}")
    for sqref, rules in ws.conditional_formatting._cf_rules.items():
        print(f"    Range {sqref}  ({len(rules)} rule(s))")
        for r in rules:
            fml = getattr(r, 'formula', None)
            if fml:
                print(f"      formula: {fml[0][:80]}{'...' if len(fml[0]) > 80 else ''}")


def delete_column_safe(ws, col_idx: int):
    """Delete a single column by 1-based index, preserving merged cell integrity."""
    ws.delete_cols(col_idx)


def clear_cf(ws):
    """Remove all conditional formatting rules from the worksheet."""
    ws.conditional_formatting._cf_rules.clear()


def add_cf_rules(ws):
    """
    Add GREEN / RED FormulaRule to each of the 18 result columns.
    Result blocks: Site1=cols10-15, Site2=cols22-27, Site3=cols34-39
    Command cols:  D(4)=off0, E(5)=off1, F(6)=off2, G(7)=off3, H(8)=off4, I(9)=off5
    """
    total_rules = 0
    for block_start in RESULT_BLOCKS:
        for offset, cmd_col in enumerate(CMD_COLS):
            res_col    = block_start + offset          # 1-indexed result column
            res_letter = get_column_letter(res_col)
            cmd_letter = get_column_letter(cmd_col)

            # Range string e.g. "J13:J200"
            range_str = f"{res_letter}{CF_ROW_START}:{res_letter}{CF_ROW_END}"

            # Formulae reference the TOP cell of the applied range
            top = CF_ROW_START
            green_formula = (f"AND({res_letter}{top}<>\"\","
                             f"ABS({res_letter}{top}-${cmd_letter}{top})<=5)")
            red_formula   = (f"AND({res_letter}{top}<>\"\","
                             f"ABS({res_letter}{top}-${cmd_letter}{top})>5)")

            ws.conditional_formatting.add(
                range_str,
                FormulaRule(formula=[green_formula], fill=GREEN_FILL, stopIfTrue=False)
            )
            ws.conditional_formatting.add(
                range_str,
                FormulaRule(formula=[red_formula],   fill=RED_FILL,   stopIfTrue=False)
            )
            total_rules += 2
            print(f"      CF added → {range_str}  "
                  f"GREEN: {green_formula[:55]}  |  RED: {red_formula[:55]}")
    print(f"      Total rules added: {total_rules}")


# ── main loop ─────────────────────────────────────────────────────────────────
for label, filepath in FILES.items():
    print(f"\n{'='*70}")
    print(f"  Processing: {label}")
    print(f"  File: {filepath}")
    print(f"{'='*70}")

    wb = openpyxl.load_workbook(filepath, keep_vba=False)
    ws = wb.active
    print(f"  Sheet: '{ws.title}'  |  Max col: {ws.max_column}  |  Max row: {ws.max_row}")

    # ── Step 1: Print existing CF ──────────────────────────────────────────
    print_existing_cf(ws, label)

    # ── Step 2A: Delete columns (right → left) ────────────────────────────
    print(f"\n  Deleting columns (right→left): AJ(36), W(23), J(10)")
    for col_idx in DEL_COLS:
        col_ltr = get_column_letter(col_idx)
        print(f"    Deleting col {col_idx} ({col_ltr}) ...")
        delete_column_safe(ws, col_idx)
    print(f"  After deletion — max col: {ws.max_column}")

    # ── Step 2C: Clear existing CF rules ─────────────────────────────────
    old_count = len(ws.conditional_formatting._cf_rules)
    clear_cf(ws)
    print(f"\n  Cleared {old_count} CF rule sets.")

    # ── Step 2D: Add new CF rules ─────────────────────────────────────────
    print(f"\n  Adding new CF rules (rows {CF_ROW_START}:{CF_ROW_END}):")
    add_cf_rules(ws)

    # ── Step 2E: Save ─────────────────────────────────────────────────────
    wb.save(filepath)
    print(f"\n  SAVED: {filepath}")

print(f"\n{'='*70}")
print("  ALL DONE — 3 files processed successfully.")
print(f"{'='*70}\n")
