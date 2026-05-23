"""
Baoyuan Historical Data Analysis
Main Grid (26 Apr – 6 May, 15-min intervals) vs HyESys (30 Apr – 6 May, 1-min intervals)
HyESys activation: 2026-05-06 16:29:31
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import numpy as np
from pathlib import Path

DATA_DIR = Path(r"C:\Users\JasonOng\Desktop\Data analytics\Baoyuan China\LV room\1-6may")
HYESYS_ACTIVATION = pd.Timestamp("2026-05-06 16:29:31")

# ─────────────────────────────────────────────
# 1. LOAD FILES
# ─────────────────────────────────────────────
print("=" * 60)
print("LOADING FILES")
print("=" * 60)

# Main Grid — headers in row 0
mg_raw = pd.read_excel(DATA_DIR / "Main Grid 26apr-6may.xlsx")
print(f"Main Grid raw shape: {mg_raw.shape}")

# HyESys — row 0 is blank/merged, row 1 has actual Chinese headers
hy_raw = pd.read_excel(DATA_DIR / "HyESys 30apr-6may.xlsx", header=1)
print(f"HyESys raw shape:    {hy_raw.shape}")

# ─────────────────────────────────────────────
# 2. CLEAN MAIN GRID
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("MAIN GRID — COLUMN MAPPING")
print("=" * 60)

mg = mg_raw.copy()

# Strip tab artifacts from timestamp and asset columns
mg["日期"] = mg["日期"].astype(str).str.replace(r"\t", "", regex=True).str.strip()
mg["timestamp"] = pd.to_datetime(mg["日期"], errors="coerce")
mg = mg.dropna(subset=["timestamp"])
mg = mg.sort_values("timestamp").reset_index(drop=True)

# Rename key columns to English
mg_rename = {
    "瞬时有功":   "kW_total",
    "瞬时无功":   "kVAr_total",
    "A相电流":    "I_A",
    "B相":        "I_B",
    "C相":        "I_C",
    "零线":       "I_N",
    "A相电压":    "V_A",
    "B相.1":      "V_B",
    "C相.1":      "V_C",
    "总功率因数":  "PF_total",
    "A相瞬时有功": "kW_A",
    "B相瞬时":    "kW_B",
    "C相瞬时":    "kW_C",
    "正向有功总":  "kWh_fwd",
}
mg = mg.rename(columns=mg_rename)

# Force numeric
for col in ["kW_total", "kVAr_total", "I_A", "I_B", "I_C", "I_N",
            "V_A", "V_B", "V_C", "PF_total", "kW_A", "kW_B", "kW_C"]:
    mg[col] = pd.to_numeric(mg[col], errors="coerce")

# Compute RMS current (average of three phases)
mg["I_rms_avg"] = (mg["I_A"] + mg["I_B"] + mg["I_C"]) / 3

print(f"Main Grid cleaned: {len(mg)} records")
print(f"  Period: {mg['timestamp'].min()} → {mg['timestamp'].max()}")
print(f"  Interval check (most common gap): {mg['timestamp'].diff().mode()[0]}")

# ─────────────────────────────────────────────
# 3. CLEAN HYESYS
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("HYESYS — COLUMN MAPPING")
print("=" * 60)

hy = hy_raw.copy()
print("HyESys raw columns:")
for i, c in enumerate(hy.columns):
    print(f"  [{i}] {c}")

# Find timestamp column
ts_col = [c for c in hy.columns if "时间" in str(c) or "timestamp" in str(c).lower()]
print(f"\nTimestamp column detected: {ts_col}")

hy["timestamp"] = pd.to_datetime(hy[ts_col[0]], errors="coerce")
hy = hy.dropna(subset=["timestamp"])
hy = hy.sort_values("timestamp").reset_index(drop=True)

# Rename HyESys key columns
hy_rename = {
    "A相电流(A)":       "hy_I_A",
    "B相电流(A)":       "hy_I_B",
    "C相电流(A)":       "hy_I_C",
    "A相电压(V)":       "hy_V_A",
    "B相电压(V)":       "hy_V_B",
    "C相电压(V)":       "hy_V_C",
    "总输出有功功率(kW)": "hy_kW_out",
    "总输出无功功率(kVar)": "hy_kVAr_out",
    "总输出视在功率(kVA)": "hy_kVA_out",
    "输出总相功率因素":   "hy_PF_out",
    "3PAmpimb":          "hy_amp_imbalance",
    "N":                 "hy_N",
    "电网频率(Hz)":      "hy_freq",
    "输入功率(kW)":      "hy_kW_in",
    "温度(℃)":           "hy_temp",
}
hy = hy.rename(columns={k: v for k, v in hy_rename.items() if k in hy.columns})

numeric_hy = ["hy_I_A", "hy_I_B", "hy_I_C", "hy_V_A", "hy_V_B", "hy_V_C",
              "hy_kW_out", "hy_kVAr_out", "hy_kVA_out", "hy_PF_out",
              "hy_amp_imbalance", "hy_N", "hy_freq", "hy_kW_in", "hy_temp"]
for col in numeric_hy:
    if col in hy.columns:
        hy[col] = pd.to_numeric(hy[col], errors="coerce")

print(f"\nHyESys cleaned: {len(hy)} records")
print(f"  Period: {hy['timestamp'].min()} → {hy['timestamp'].max()}")
print(f"  Interval check (most common gap): {hy['timestamp'].diff().mode()[0]}")

# ─────────────────────────────────────────────
# 4. SPLIT: PRE vs POST ACTIVATION
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("ACTIVATION SPLIT")
print("=" * 60)

mg_pre  = mg[mg["timestamp"] <  HYESYS_ACTIVATION].copy()
mg_post = mg[mg["timestamp"] >= HYESYS_ACTIVATION].copy()

hy_pre  = hy[hy["timestamp"] <  HYESYS_ACTIVATION].copy()
hy_post = hy[hy["timestamp"] >= HYESYS_ACTIVATION].copy()

print(f"Main Grid — PRE  activation: {len(mg_pre):4d} records  ({mg_pre['timestamp'].min()} → {mg_pre['timestamp'].max()})")
print(f"Main Grid — POST activation: {len(mg_post):4d} records  ({mg_post['timestamp'].min()} → {mg_post['timestamp'].max()})")
print(f"HyESys    — PRE  activation: {len(hy_pre):4d} records  ({hy_pre['timestamp'].min() if len(hy_pre) else 'N/A'} → {hy_pre['timestamp'].max() if len(hy_pre) else 'N/A'})")
print(f"HyESys    — POST activation: {len(hy_post):4d} records  ({hy_post['timestamp'].min() if len(hy_post) else 'N/A'} → {hy_post['timestamp'].max() if len(hy_post) else 'N/A'})")

# ─────────────────────────────────────────────
# 5. BASELINE STATISTICS (PRE-ACTIVATION)
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 1 — BASELINE ELECTRICAL PATTERN (PRE-ACTIVATION)")
print("=" * 60)

def stats(series, label, unit):
    s = series.dropna()
    if len(s) == 0:
        print(f"  {label}: NO DATA")
        return
    print(f"  {label:30s}  mean={s.mean():.3f}  median={s.median():.3f}  min={s.min():.3f}  max={s.max():.3f}  std={s.std():.3f}  [{unit}]")

stats(mg_pre["kW_total"],    "Active Power",         "kW")
stats(mg_pre["kVAr_total"],  "Reactive Power",       "kVAr")
stats(mg_pre["PF_total"],    "Power Factor",         "PF")
stats(mg_pre["I_A"],         "Current Phase A",      "A")
stats(mg_pre["I_B"],         "Current Phase B",      "A")
stats(mg_pre["I_C"],         "Current Phase C",      "A")
stats(mg_pre["I_rms_avg"],   "Current RMS avg",      "A")
stats(mg_pre["I_N"],         "Neutral Current",      "A")
stats(mg_pre["V_A"],         "Voltage Phase A",      "V")
stats(mg_pre["V_B"],         "Voltage Phase B",      "V")
stats(mg_pre["V_C"],         "Voltage Phase C",      "V")

# Phase imbalance (pre)
i_std_pre = mg_pre[["I_A","I_B","I_C"]].std(axis=1)
i_mean_pre = mg_pre[["I_A","I_B","I_C"]].mean(axis=1)
imbalance_pre = (i_std_pre / i_mean_pre.replace(0, np.nan)) * 100
print(f"  {'Phase Imbalance %':30s}  mean={imbalance_pre.mean():.2f}%  max={imbalance_pre.max():.2f}%")

# ─────────────────────────────────────────────
# 6. POST-ACTIVATION STATISTICS
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 2 — POST-ACTIVATION MAIN GRID (HYESYS RUNNING)")
print("=" * 60)

stats(mg_post["kW_total"],   "Active Power",         "kW")
stats(mg_post["kVAr_total"], "Reactive Power",       "kVAr")
stats(mg_post["PF_total"],   "Power Factor",         "PF")
stats(mg_post["I_A"],        "Current Phase A",      "A")
stats(mg_post["I_B"],        "Current Phase B",      "A")
stats(mg_post["I_C"],        "Current Phase C",      "A")
stats(mg_post["I_rms_avg"],  "Current RMS avg",      "A")
stats(mg_post["I_N"],        "Neutral Current",      "A")
stats(mg_post["V_A"],        "Voltage Phase A",      "V")
stats(mg_post["V_B"],        "Voltage Phase B",      "V")
stats(mg_post["V_C"],        "Voltage Phase C",      "V")

i_std_post = mg_post[["I_A","I_B","I_C"]].std(axis=1)
i_mean_post = mg_post[["I_A","I_B","I_C"]].mean(axis=1)
imbalance_post = (i_std_post / i_mean_post.replace(0, np.nan)) * 100
print(f"  {'Phase Imbalance %':30s}  mean={imbalance_post.mean():.2f}%  max={imbalance_post.max():.2f}%")

# ─────────────────────────────────────────────
# 7. DELTA: PRE vs POST
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 3 — DELTA: PRE vs POST (mean values)")
print("=" * 60)

def delta(col, unit, scale=1):
    pre_mean  = mg_pre[col].dropna().mean()  * scale
    post_mean = mg_post[col].dropna().mean() * scale
    d = post_mean - pre_mean
    pct = (d / pre_mean * 100) if pre_mean != 0 else float("nan")
    direction = "DECREASE" if d < 0 else "INCREASE"
    print(f"  {col:20s}  pre={pre_mean:.3f}  post={post_mean:.3f}  delta={d:+.3f} {unit}  ({pct:+.1f}%)  [{direction}]")

delta("kW_total",   "kW")
delta("kVAr_total", "kVAr")
delta("PF_total",   "PF")
delta("I_A",        "A")
delta("I_B",        "A")
delta("I_C",        "A")
delta("I_rms_avg",  "A")
delta("I_N",        "A")

# Imbalance delta
print(f"  {'Phase Imbalance %':20s}  pre={imbalance_pre.mean():.2f}%  post={imbalance_post.mean():.2f}%  delta={imbalance_post.mean()-imbalance_pre.mean():+.2f}%")

# ─────────────────────────────────────────────
# 8. HYESYS OUTPUT DURING ACTIVE PERIOD
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 4 — HYESYS OUTPUT STATS (POST-ACTIVATION)")
print("=" * 60)

stats(hy_post["hy_kVAr_out"],      "HyESys kVAr Output",   "kVAr")
stats(hy_post["hy_kW_out"],        "HyESys kW Output",     "kW")
stats(hy_post["hy_kVA_out"],       "HyESys kVA Output",    "kVA")
stats(hy_post["hy_PF_out"],        "HyESys PF Output",     "PF")
stats(hy_post["hy_amp_imbalance"], "HyESys Amp Imbalance", "%")
stats(hy_post["hy_I_A"],           "HyESys Current A",     "A")
stats(hy_post["hy_I_B"],           "HyESys Current B",     "A")
stats(hy_post["hy_I_C"],           "HyESys Current C",     "A")
stats(hy_post["hy_temp"],          "Temperature",          "°C")

# ─────────────────────────────────────────────
# 9. TIMESTAMP ALIGNMENT — 15-MIN AGGREGATION
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 5 — ALIGNED COMPARISON (15-MIN WINDOWS, POST-ACTIVATION)")
print("=" * 60)

# Resample HyESys to 15-min means to match Main Grid cadence
hy_post_idx = hy_post.set_index("timestamp")
hy_15 = hy_post_idx[["hy_kVAr_out", "hy_kW_out", "hy_kVA_out",
                       "hy_PF_out", "hy_amp_imbalance",
                       "hy_I_A", "hy_I_B", "hy_I_C"]].resample("15min").mean()

mg_post_idx = mg_post.set_index("timestamp")
mg_15 = mg_post_idx[["kW_total", "kVAr_total", "PF_total",
                       "I_A", "I_B", "I_C", "I_rms_avg", "I_N"]].copy()

# Merge on nearest 15-min timestamp
merged = mg_15.join(hy_15, how="inner")
print(f"Aligned 15-min windows (post-activation): {len(merged)} records")

if len(merged) > 1:
    # Correlation: HyESys kVAr output vs Main Grid kVAr
    corr_kvar = merged[["hy_kVAr_out", "kVAr_total"]].dropna().corr().iloc[0, 1]
    corr_kw   = merged[["hy_kW_out", "kW_total"]].dropna().corr().iloc[0, 1]
    corr_pf   = merged[["hy_kVAr_out", "PF_total"]].dropna().corr().iloc[0, 1]
    print(f"\n  Correlation (HyESys kVAr out vs Main Grid kVAr):  r = {corr_kvar:.3f}")
    print(f"  Correlation (HyESys kW out  vs Main Grid kW):      r = {corr_kw:.3f}")
    print(f"  Correlation (HyESys kVAr out vs Main Grid PF):     r = {corr_pf:.3f}")

    # Per-row comparison
    print(f"\n  Aligned 15-min window detail (all post-activation records):")
    print(f"  {'Timestamp':<22} {'MG_kW':>7} {'MG_kVAr':>8} {'MG_PF':>6} {'MG_IA':>6} {'MG_IB':>6} {'MG_IC':>6} | {'HY_kVAr':>8} {'HY_kW':>7} {'HY_PF':>6}")
    print("  " + "-" * 95)
    for ts, row in merged.iterrows():
        print(f"  {str(ts):<22} "
              f"{row.get('kW_total', float('nan')):>7.3f} "
              f"{row.get('kVAr_total', float('nan')):>8.3f} "
              f"{row.get('PF_total', float('nan')):>6.3f} "
              f"{row.get('I_A', float('nan')):>6.2f} "
              f"{row.get('I_B', float('nan')):>6.2f} "
              f"{row.get('I_C', float('nan')):>6.2f} | "
              f"{row.get('hy_kVAr_out', float('nan')):>8.3f} "
              f"{row.get('hy_kW_out', float('nan')):>7.3f} "
              f"{row.get('hy_PF_out', float('nan')):>6.3f}")

# ─────────────────────────────────────────────
# 10. HOURLY PATTERN ANALYSIS (PRE vs POST by hour)
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 6 — HOURLY PATTERN (PRE vs POST by hour-of-day)")
print("=" * 60)

mg["hour"] = mg["timestamp"].dt.hour
mg_pre["hour"]  = mg_pre["timestamp"].dt.hour
mg_post["hour"] = mg_post["timestamp"].dt.hour

print(f"\n  {'Hour':>5} | {'PRE_kW':>8} {'PRE_kVAr':>9} {'PRE_I_avg':>10} {'PRE_PF':>7} | {'POST_kW':>8} {'POST_kVAr':>9} {'POST_I_avg':>10} {'POST_PF':>7}")
print("  " + "-" * 80)
for h in range(24):
    p  = mg_pre[mg_pre["hour"] == h]
    q  = mg_post[mg_post["hour"] == h]
    pre_kw   = p["kW_total"].mean()   if len(p) else float("nan")
    pre_kvar = p["kVAr_total"].mean() if len(p) else float("nan")
    pre_i    = p["I_rms_avg"].mean()  if len(p) else float("nan")
    pre_pf   = p["PF_total"].mean()   if len(p) else float("nan")
    pst_kw   = q["kW_total"].mean()   if len(q) else float("nan")
    pst_kvar = q["kVAr_total"].mean() if len(q) else float("nan")
    pst_i    = q["I_rms_avg"].mean()  if len(q) else float("nan")
    pst_pf   = q["PF_total"].mean()   if len(q) else float("nan")
    print(f"  {h:>5}h | {pre_kw:>8.3f} {pre_kvar:>9.3f} {pre_i:>10.3f} {pre_pf:>7.3f} | {pst_kw:>8.3f} {pst_kvar:>9.3f} {pst_i:>10.3f} {pst_pf:>7.3f}")

# ─────────────────────────────────────────────
# 11. ENERGY SAVINGS ESTIMATE
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 7 — ENERGY SAVINGS ESTIMATE")
print("=" * 60)

pre_I  = mg_pre["I_rms_avg"].dropna().mean()
post_I = mg_post["I_rms_avg"].dropna().mean()

if pre_I > 0:
    loss_fraction = 1 - (post_I / pre_I) ** 2
    print(f"\n  I_before (mean RMS avg, pre-activation):  {pre_I:.3f} A")
    print(f"  I_after  (mean RMS avg, post-activation): {post_I:.3f} A")
    print(f"  Current reduction:                        {pre_I - post_I:.3f} A  ({(pre_I - post_I)/pre_I*100:.1f}%)")
    print(f"  I²R loss reduction fraction:              {loss_fraction*100:.2f}%")
    print(f"\n  Note: Loss fraction = 1 - (I_after/I_before)² — R cancels out.")
    print(f"  This is the upper-bound distribution loss saving assuming all current")
    print(f"  reduction is attributable to HyESys reactive/imbalance correction.")

pre_kvar  = mg_pre["kVAr_total"].dropna().mean()
post_kvar = mg_post["kVAr_total"].dropna().mean()
print(f"\n  kVAr_before (mean, pre-activation):  {pre_kvar:.3f} kVAr")
print(f"  kVAr_after  (mean, post-activation): {post_kvar:.3f} kVAr")
print(f"  kVAr reduction:                      {pre_kvar - post_kvar:.3f} kVAr  ({(pre_kvar-post_kvar)/pre_kvar*100 if pre_kvar else float('nan'):.1f}%)")

pre_kw  = mg_pre["kW_total"].dropna().mean()
post_kw = mg_post["kW_total"].dropna().mean()
print(f"\n  kW_before (mean, pre-activation):  {pre_kw:.3f} kW")
print(f"  kW_after  (mean, post-activation): {post_kw:.3f} kW")
print(f"  kW reduction:                      {pre_kw - post_kw:.3f} kW  ({(pre_kw-post_kw)/pre_kw*100 if pre_kw else float('nan'):.1f}%)")

# ─────────────────────────────────────────────
# 12. DATA QUALITY FLAGS
# ─────────────────────────────────────────────
print("\n" + "=" * 60)
print("SECTION 8 — DATA QUALITY FLAGS")
print("=" * 60)

# Zero/missing current rows in MG
zero_I = mg[(mg["I_A"] == 0) | (mg["I_B"] == 0) | (mg["I_C"] == 0)]
print(f"\n  Main Grid rows with any zero-phase current: {len(zero_I)}")

# Gaps in Main Grid
mg_sorted = mg.sort_values("timestamp")
mg_gaps = mg_sorted["timestamp"].diff()
large_gaps = mg_gaps[mg_gaps > pd.Timedelta("30min")]
print(f"  Main Grid gaps > 30 min: {len(large_gaps)}")
if len(large_gaps) > 0:
    for ts, gap in zip(mg_sorted["timestamp"][large_gaps.index], large_gaps):
        print(f"    Gap of {gap} ending at {ts}")

# Gaps in HyESys
hy_sorted = hy.sort_values("timestamp")
hy_gaps = hy_sorted["timestamp"].diff()
large_hy_gaps = hy_gaps[hy_gaps > pd.Timedelta("10min")]
print(f"\n  HyESys gaps > 10 min: {len(large_hy_gaps)}")
if len(large_hy_gaps) > 0:
    for ts, gap in zip(hy_sorted["timestamp"][large_hy_gaps.index], large_hy_gaps):
        print(f"    Gap of {gap} ending at {ts}")

# HyESys pre-activation rows (what was it doing before 16:29:31?)
print(f"\n  HyESys records BEFORE activation ({HYESYS_ACTIVATION}): {len(hy_pre)}")
if len(hy_pre) > 0:
    print(f"  Period: {hy_pre['timestamp'].min()} → {hy_pre['timestamp'].max()}")
    stats(hy_pre["hy_kVAr_out"], "  kVAr out (pre-activation)", "kVAr")
    stats(hy_pre["hy_kW_out"],   "  kW out (pre-activation)",   "kW")

# B相 and 零线 in Main Grid — check if consistently 0 (meter issue)
print(f"\n  Main Grid I_B — unique non-zero values: {(mg['I_B'] != 0).sum()}")
print(f"  Main Grid I_N — NaN count: {mg['I_N'].isna().sum()} / {len(mg)}")

print("\n" + "=" * 60)
print("ANALYSIS COMPLETE")
print("=" * 60)
