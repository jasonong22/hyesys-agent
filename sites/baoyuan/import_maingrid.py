"""
import_maingrid.py — Baoyuan Main Grid History Importer

Imports the periodic Main Grid Excel file (sent by Baoyuan after each
experiment period) into the maingrid_history table in baoyuan.db.
Applies the same multiplying factors used in the v3 analysis:
  kW / kVAr × 4000 | Current × 40 | Voltage × 100

Usage:
    python sites/baoyuan/import_maingrid.py --file "path/to/Main Grid file.xlsx"
    python sites/baoyuan/import_maingrid.py --file "path/to/file.xlsx" --dry-run

Options:
    --file      Path to the Main Grid Excel file (required)
    --dry-run   Print rows that would be imported without writing to DB
    --mf-power  Override multiplying factor for kW/kVAr (default: 4000)
    --mf-current Override multiplying factor for current (default: 40)
    --mf-voltage Override multiplying factor for voltage (default: 100)
"""

import argparse
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERROR: pandas and numpy required. Run: pip install pandas openpyxl")
    sys.exit(1)

DB_PATH = Path(__file__).parent / "data" / "baoyuan.db"


# ─────────────────────────────────────────────
# LOAD & TRANSFORM
# ─────────────────────────────────────────────
def load_maingrid(file_path: Path, mf_power: float, mf_current: float, mf_voltage: float) -> pd.DataFrame:
    df_raw = pd.read_excel(file_path)

    # Parse timestamp — strip tabs and whitespace from date column
    date_col = "日期"
    if date_col not in df_raw.columns:
        # Try first column if header not found
        date_col = df_raw.columns[0]

    df_raw[date_col] = df_raw[date_col].astype(str).str.replace(r"\t", "", regex=True).str.strip()
    df_raw["timestamp"] = pd.to_datetime(df_raw[date_col], errors="coerce")
    df_raw = df_raw.dropna(subset=["timestamp"]).sort_values("timestamp").reset_index(drop=True)

    # Rename known columns (same mapping as baoyuan_analysis_v3.py)
    rename = {
        "瞬时有功":   "kW_raw",
        "瞬时无功":   "kVAr_raw",
        "A相电流":    "I_A_raw",
        "C相":        "I_C_raw",
        "A相电压":    "V_A_raw",
        "C相.1":      "V_C_raw",
        "总功率因数":  "PF",
        "正向有功总":  "kWh_raw",
        "A相瞬时有功": "kW_A_raw",
        "C相瞬时":    "kW_C_raw",
    }
    df = df_raw.rename(columns=rename)

    for col in ["kW_raw", "kVAr_raw", "I_A_raw", "I_C_raw",
                "V_A_raw", "V_C_raw", "PF", "kWh_raw", "kW_A_raw", "kW_C_raw"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Apply multiplying factors
    df["kW"]   = df["kW_raw"]   * mf_power   if "kW_raw"   in df.columns else np.nan
    df["kVAr"] = df["kVAr_raw"] * mf_power   if "kVAr_raw" in df.columns else np.nan
    df["I_A"]  = df["I_A_raw"]  * mf_current if "I_A_raw"  in df.columns else np.nan
    df["I_C"]  = df["I_C_raw"]  * mf_current if "I_C_raw"  in df.columns else np.nan
    df["V_A"]  = df["V_A_raw"]  * mf_voltage if "V_A_raw"  in df.columns else np.nan
    df["V_C"]  = df["V_C_raw"]  * mf_voltage if "V_C_raw"  in df.columns else np.nan
    df["kWh"]  = df["kWh_raw"]  * mf_power   if "kWh_raw"  in df.columns else np.nan
    df["kW_A"] = df["kW_A_raw"] * mf_power   if "kW_A_raw" in df.columns else np.nan
    df["kW_C"] = df["kW_C_raw"] * mf_power   if "kW_C_raw" in df.columns else np.nan

    return df[["timestamp", "kW", "kVAr", "PF", "I_A", "I_C", "V_A", "V_C", "kWh", "kW_A", "kW_C"]]


# ─────────────────────────────────────────────
# IMPORT TO DB
# ─────────────────────────────────────────────
def import_to_db(df: pd.DataFrame, dry_run: bool) -> tuple[int, int]:
    imported_at = datetime.now(timezone.utc).isoformat()
    inserted = 0
    skipped  = 0

    if dry_run:
        print(f"\n[DRY RUN] Would import {len(df)} rows into maingrid_history")
        print(df.to_string(index=False, max_rows=20))
        return len(df), 0

    conn = sqlite3.connect(str(DB_PATH))
    for _, row in df.iterrows():
        ts = row["timestamp"]
        if pd.isna(ts):
            skipped += 1
            continue
        ts_str = ts.isoformat()
        try:
            conn.execute(
                """INSERT OR IGNORE INTO maingrid_history
                   (timestamp, kW, kVAr, PF, I_A, I_C, V_A, V_C, kWh, kW_A, kW_C, imported_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    ts_str,
                    _safe(row, "kW"),   _safe(row, "kVAr"),
                    _safe(row, "PF"),   _safe(row, "I_A"),  _safe(row, "I_C"),
                    _safe(row, "V_A"),  _safe(row, "V_C"),  _safe(row, "kWh"),
                    _safe(row, "kW_A"), _safe(row, "kW_C"),
                    imported_at,
                ),
            )
            inserted += 1
        except sqlite3.IntegrityError:
            skipped += 1

    conn.commit()
    conn.close()
    return inserted, skipped


def _safe(row, col):
    import math
    v = row.get(col)
    if v is None:
        return None
    try:
        f = float(v)
        return None if math.isnan(f) else round(f, 4)
    except (TypeError, ValueError):
        return None


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Import Baoyuan Main Grid history")
    parser.add_argument("--file",        required=True, help="Path to Main Grid Excel file")
    parser.add_argument("--dry-run",     action="store_true", help="Preview without writing")
    parser.add_argument("--mf-power",    type=float, default=4000.0, help="kW/kVAr multiplying factor")
    parser.add_argument("--mf-current",  type=float, default=40.0,   help="Current multiplying factor")
    parser.add_argument("--mf-voltage",  type=float, default=100.0,  help="Voltage multiplying factor")
    args = parser.parse_args()

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"ERROR: File not found: {file_path}")
        sys.exit(1)

    if not args.dry_run and not DB_PATH.exists():
        print(f"ERROR: Database not found: {DB_PATH}")
        print("Run agent1.py first to initialise the database.")
        sys.exit(1)

    print(f"Loading: {file_path.name}")
    print(f"Multiplying factors: kW/kVAr ×{args.mf_power}  I ×{args.mf_current}  V ×{args.mf_voltage}")

    df = load_maingrid(file_path, args.mf_power, args.mf_current, args.mf_voltage)
    print(f"Rows loaded: {len(df)}  ({df['timestamp'].min()} → {df['timestamp'].max()})")

    inserted, skipped = import_to_db(df, args.dry_run)

    if not args.dry_run:
        print(f"Inserted: {inserted}  Skipped (duplicate/null): {skipped}")
        print(f"Database: {DB_PATH}")


if __name__ == "__main__":
    main()
