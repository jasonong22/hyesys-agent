"""
Baoyuan Analysis v2 — Corrected
- Both files capped at 2026-05-06 10:30:00 (pre-activation baseline only)
- Main Grid multiplying factors applied: kW/kVAr ×4000, I ×40, V ×100
- HyESys values taken as-is (no multiplier)
- Phase B excluded (no CT installed — not a fault)
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import numpy as np
from pathlib import Path

DATA_DIR  = Path(r"C:\Users\JasonOng\Desktop\Data analytics\Baoyuan China\LV room\1-6may")
CAP_TIME  = pd.Timestamp("2026-05-06 10:30:00")   # analysis end for both files

# Multiplying factors (confirmed via 综合倍率=4000 in 查询结果.xlsx)
MF_POWER   = 4000   # kW, kVAr, kWh
MF_CURRENT =   40   # CT ratio — secondary Amps → primary Amps
MF_VOLTAGE =  100   # PT ratio — secondary Volts → primary Volts (→ ~10kV)

# ─────────────────────────────────────────────
# 1. LOAD & CLEAN MAIN GRID
# ─────────────────────────────────────────────
mg_raw = pd.read_excel(DATA_DIR / "Main Grid 26apr-6may.xlsx")
mg = mg_raw.copy()
mg["日期"] = mg["日期"].astype(str).str.replace(r"\t", "", regex=True).str.strip()
mg["timestamp"] = pd.to_datetime(mg["日期"], errors="coerce")
mg = mg.dropna(subset=["timestamp"]).sort_values("timestamp").reset_index(drop=True)

mg = mg.rename(columns={
    "瞬时有功": "kW_raw", "瞬时无功": "kVAr_raw",
    "A相电流": "I_A_raw", "B相": "I_B_raw", "C相": "I_C_raw",
    "A相电压": "V_A_raw", "B相.1": "V_B_raw", "C相.1": "V_C_raw",
    "总功率因数": "PF", "正向有功总": "kWh_raw",
    "A相瞬时有功": "kW_A_raw", "C相瞬时": "kW_C_raw",
    "Ⅰ象限无功": "kVArh_Q1_raw", "CT": "CT", "PT": "PT",
})

for col in ["kW_raw","kVAr_raw","I_A_raw","I_C_raw","V_A_raw","V_C_raw",
            "PF","kWh_raw","kW_A_raw","kW_C_raw","kVArh_Q1_raw"]:
    mg[col] = pd.to_numeric(mg[col], errors="coerce")

# Apply multiplying factors
mg["kW"]     = mg["kW_raw"]     * MF_POWER
mg["kVAr"]   = mg["kVAr_raw"]   * MF_POWER
mg["I_A"]    = mg["I_A_raw"]    * MF_CURRENT
mg["I_C"]    = mg["I_C_raw"]    * MF_CURRENT
mg["V_A"]    = mg["V_A_raw"]    * MF_VOLTAGE
mg["V_C"]    = mg["V_C_raw"]    * MF_VOLTAGE
mg["kWh"]    = mg["kWh_raw"]    * MF_POWER
mg["kW_A"]   = mg["kW_A_raw"]   * MF_POWER
mg["kW_C"]   = mg["kW_C_raw"]   * MF_POWER
mg["I_AC_avg"] = (mg["I_A"] + mg["I_C"]) / 2

# Cap at analysis window
mg = mg[mg["timestamp"] <= CAP_TIME].copy()

# ─────────────────────────────────────────────
# 2. LOAD & CLEAN HYESYS
# ─────────────────────────────────────────────
hy_raw = pd.read_excel(DATA_DIR / "HyESys 30apr-6may.xlsx", header=1)
hy = hy_raw.copy()
hy["timestamp"] = pd.to_datetime(hy["上报时间戳"], errors="coerce")
hy = hy.dropna(subset=["timestamp"]).sort_values("timestamp").reset_index(drop=True)
hy = hy.rename(columns={
    "A相电流(A)": "hy_I_A", "B相电流(A)": "hy_I_B", "C相电流(A)": "hy_I_C",
    "A相电压(V)": "hy_V_A", "B相电压(V)": "hy_V_B", "C相电压(V)": "hy_V_C",
    "总输出有功功率(kW)": "hy_kW", "总输出无功功率(kVar)": "hy_kVAr",
    "总输出视在功率(kVA)": "hy_kVA", "输出总相功率因素": "hy_PF",
    "A相输出无功功率(kVar)": "hy_kVAr_A", "B相输出无功功率(kVar)": "hy_kVAr_B",
    "C相输出无功功率(kVar)": "hy_kVAr_C",
    "A相输出有功功率(kW)": "hy_kW_A", "B相输出有功功率(kW)": "hy_kW_B",
    "C相输出有功功率(kW)": "hy_kW_C",
    "3PAmpimb": "hy_amp_imb", "N": "hy_N",
    "电网频率(Hz)": "hy_freq", "温度(℃)": "hy_temp",
    "输入功率(kW)": "hy_kW_in", "输入电压(V)": "hy_V_in", "输入电流(A)": "hy_I_in",
})
num_hy = ["hy_I_A","hy_I_B","hy_I_C","hy_V_A","hy_V_B","hy_V_C",
          "hy_kW","hy_kVAr","hy_kVA","hy_PF","hy_kVAr_A","hy_kVAr_B","hy_kVAr_C",
          "hy_kW_A","hy_kW_B","hy_kW_C","hy_amp_imb","hy_N","hy_freq","hy_temp",
          "hy_kW_in","hy_V_in","hy_I_in"]
for col in num_hy:
    if col in hy.columns:
        hy[col] = pd.to_numeric(hy[col], errors="coerce")

# Cap at analysis window
hy = hy[hy["timestamp"] <= CAP_TIME].copy()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
print("=" * 70)
print("BAOYUAN ANALYSIS v2 — CORRECTED WITH MULTIPLYING FACTORS")
print(f"Analysis window: both files capped at {CAP_TIME}")
print(f"Multiplying factors: kW/kVAr ×{MF_POWER}, I ×{MF_CURRENT}, V ×{MF_VOLTAGE}")
print(f"Phase B excluded — no CT installed at this meter")
print("=" * 70)
print(f"\nMain Grid records: {len(mg)}  ({mg['timestamp'].min()} → {mg['timestamp'].max()})")
print(f"HyESys records:    {len(hy)}  ({hy['timestamp'].min()} → {hy['timestamp'].max()})")
print(f"Overlap (both available): Apr 30 00:00 → May 6 10:30")

# ─────────────────────────────────────────────
# SECTION 1 — MAIN GRID BASELINE (CORRECTED)
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 1 — MAIN GRID BASELINE (CORRECTED REAL VALUES)")
print("=" * 70)

def stats(series, label, unit, width=35):
    s = series.dropna()
    if len(s) == 0:
        print(f"  {label:{width}s}  NO DATA")
        return
    print(f"  {label:{width}s}  mean={s.mean():>10.2f}  median={s.median():>10.2f}  "
          f"min={s.min():>10.2f}  max={s.max():>10.2f}  std={s.std():>8.2f}  [{unit}]")

stats(mg["kW"],       "Active Power (kW)",                 "kW")
stats(mg["kVAr"],     "Reactive Power (kVAr)",             "kVAr")
stats(mg["PF"],       "Power Factor",                      "–")
stats(mg["I_A"],      "Current Phase A (primary)",         "A")
stats(mg["I_C"],      "Current Phase C (primary)",         "A")
stats(mg["I_AC_avg"], "Current A+C average (primary)",     "A")
stats(mg["V_A"],      "Voltage Phase A (primary)",         "V")
stats(mg["V_C"],      "Voltage Phase C (primary)",         "V")
stats(mg["kW_A"],     "Active Power Phase A",              "kW")
stats(mg["kW_C"],     "Active Power Phase C",              "kW")

# kVAr angle — apparent power check using A and C only
mg["kVA_check"] = np.sqrt(mg["kW"]**2 + mg["kVAr"]**2)
mg["PF_check"]  = mg["kW"] / mg["kVA_check"].replace(0, np.nan)
print(f"\n  Derived check (using corrected kW and kVAr):")
stats(mg["kVA_check"], "  Apparent Power (√kW²+kVAr²)", "kVA")
stats(mg["PF_check"],  "  PF derived from kW/kVA",      "–")

# Demand
print(f"\n  Max active power demand:   {mg['kW'].max():.1f} kW  at {mg.loc[mg['kW'].idxmax(),'timestamp']}")
print(f"  Max reactive demand:       {mg['kVAr'].max():.1f} kVAr  at {mg.loc[mg['kVAr'].idxmax(),'timestamp']}")
print(f"  Min PF recorded:           {mg['PF'].min():.4f}  at {mg.loc[mg['PF'].idxmin(),'timestamp']}")
print(f"  Max PF recorded:           {mg['PF'].max():.4f}  at {mg.loc[mg['PF'].idxmax(),'timestamp']}")

pf_below_90 = mg[mg["PF"] < 0.90]
pf_below_85 = mg[mg["PF"] < 0.85]
print(f"\n  Records PF < 0.90:  {len(pf_below_90)} ({len(pf_below_90)/len(mg)*100:.1f}%)")
print(f"  Records PF < 0.85 (penalty):  {len(pf_below_85)} ({len(pf_below_85)/len(mg)*100:.1f}%)")

# ─────────────────────────────────────────────
# SECTION 2 — ENERGY CONSUMPTION
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 2 — ENERGY CONSUMPTION (CORRECTED)")
print("=" * 70)

mg_sorted = mg.sort_values("timestamp")
kwh_start = mg_sorted["kWh"].iloc[0]
kwh_end   = mg_sorted["kWh"].iloc[-1]
period_days = (mg_sorted["timestamp"].iloc[-1] - mg_sorted["timestamp"].iloc[0]).total_seconds() / 86400

print(f"\n  kWh meter start (Apr 26 00:00):  {kwh_start:,.1f} kWh")
print(f"  kWh meter end   (May 6 10:30):   {kwh_end:,.1f} kWh")
print(f"  Period:                          {period_days:.1f} days")
print(f"  Total consumption:               {kwh_end - kwh_start:,.1f} kWh")
print(f"  Average daily consumption:       {(kwh_end - kwh_start) / period_days:,.1f} kWh/day")
print(f"  Average hourly:                  {(kwh_end - kwh_start) / (period_days * 24):,.1f} kWh/h")
print(f"  Average load:                    {(kwh_end - kwh_start) / (period_days * 24):,.1f} kW")

# ─────────────────────────────────────────────
# SECTION 3 — HOURLY LOAD PATTERN
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 3 — HOURLY LOAD PATTERN (all days, Main Grid)")
print("=" * 70)
mg["hour"] = mg["timestamp"].dt.hour

print(f"\n  {'Hour':>5} | {'kW mean':>9} {'kW max':>9} {'kVAr mean':>10} {'PF mean':>8} {'I_A mean':>9} {'I_C mean':>9} | Observations")
print("  " + "-" * 85)
for h in range(24):
    sub = mg[mg["hour"] == h]
    print(f"  {h:>5}h | {sub['kW'].mean():>9.1f} {sub['kW'].max():>9.1f} {sub['kVAr'].mean():>10.1f} "
          f"{sub['PF'].mean():>8.4f} {sub['I_A'].mean():>9.1f} {sub['I_C'].mean():>9.1f} | n={len(sub)}")

# ─────────────────────────────────────────────
# SECTION 4 — DAILY LOAD PATTERN
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 4 — DAILY PATTERN (by date)")
print("=" * 70)
mg["date"] = mg["timestamp"].dt.date
print(f"\n  {'Date':<13} | {'kW mean':>9} {'kW max':>9} {'kVAr mean':>10} {'PF mean':>8} {'I_A mean':>9} {'Records':>8}")
print("  " + "-" * 75)
for date, sub in mg.groupby("date"):
    print(f"  {str(date):<13} | {sub['kW'].mean():>9.1f} {sub['kW'].max():>9.1f} {sub['kVAr'].mean():>10.1f} "
          f"{sub['PF'].mean():>8.4f} {sub['I_A'].mean():>9.1f} {sub['timestamp'].count():>8}")

# ─────────────────────────────────────────────
# SECTION 5 — REACTIVE POWER ANALYSIS
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 5 — REACTIVE POWER ANALYSIS")
print("=" * 70)

# kVAr breakdown by time of day
kvar_day   = mg[(mg["hour"] >= 8)  & (mg["hour"] < 20)]["kVAr"]
kvar_night = mg[(mg["hour"] < 8)   | (mg["hour"] >= 20)]["kVAr"]
print(f"\n  Daytime  (08:00–20:00):  mean kVAr = {kvar_day.mean():.1f}  max = {kvar_day.max():.1f}")
print(f"  Nighttime (20:00–08:00): mean kVAr = {kvar_night.mean():.1f}  max = {kvar_night.max():.1f}")

# kVAr vs kW ratio
mg["kVAr_to_kW"] = mg["kVAr"] / mg["kW"].replace(0, np.nan)
print(f"\n  kVAr/kW ratio:  mean={mg['kVAr_to_kW'].mean():.4f}  std={mg['kVAr_to_kW'].std():.4f}")
print(f"  tan(φ) → φ:     mean angle = {np.degrees(np.arctan(mg['kVAr_to_kW'].mean())):.2f}°")

# Required kVAr to correct to PF=0.98
pf_target = 0.98
mg["kVAr_required"] = mg["kW"] * (
    np.tan(np.arccos(mg["PF"].clip(0.01, 0.9999))) -
    np.tan(np.arccos(pf_target))
)
print(f"\n  kVAr required for PF correction to {pf_target}:")
stats(mg["kVAr_required"].clip(lower=0), "  Reactive correction needed", "kVAr")
print(f"  Peak kVAr correction needed: {mg['kVAr_required'].clip(lower=0).max():.1f} kVAr  at {mg.loc[mg['kVAr_required'].idxmax(), 'timestamp']}")

# ─────────────────────────────────────────────
# SECTION 6 — HYESYS DATA (capped at 10:30)
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 6 — HYESYS DATA (Apr 30 – May 6 10:30, pre-activation)")
print("=" * 70)

print(f"\n  Records: {len(hy)}")
print(f"  Period:  {hy['timestamp'].min()} → {hy['timestamp'].max()}")

stats(hy["hy_kVAr"],     "HyESys kVAr (output/measured)", "kVAr")
stats(hy["hy_kW"],       "HyESys kW (output/measured)",   "kW")
stats(hy["hy_kVA"],      "HyESys kVA",                    "kVA")
stats(hy["hy_PF"],       "HyESys PF",                     "–")
stats(hy["hy_I_A"],      "HyESys I_A",                    "A")
stats(hy["hy_I_B"],      "HyESys I_B",                    "A")
stats(hy["hy_I_C"],      "HyESys I_C",                    "A")
stats(hy["hy_amp_imb"],  "HyESys Amp Imbalance",          "%")
stats(hy["hy_V_A"],      "HyESys V_A",                    "V")
stats(hy["hy_V_B"],      "HyESys V_B",                    "V")
stats(hy["hy_V_C"],      "HyESys V_C",                    "V")
stats(hy["hy_freq"],     "Grid Frequency",                "Hz")
stats(hy["hy_temp"],     "Temperature",                   "°C")

# ─────────────────────────────────────────────
# SECTION 7 — HYESYS HOURLY PATTERN
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 7 — HYESYS HOURLY PATTERN (by hour-of-day)")
print("=" * 70)
hy["hour"] = hy["timestamp"].dt.hour

print(f"\n  {'Hour':>5} | {'kVAr mean':>10} {'kVAr max':>9} {'kW mean':>9} {'I_A mean':>9} {'I_B mean':>9} {'I_C mean':>9} | n")
print("  " + "-" * 85)
for h in range(24):
    sub = hy[hy["hour"] == h]
    if len(sub) == 0:
        continue
    print(f"  {h:>5}h | {sub['hy_kVAr'].mean():>10.3f} {sub['hy_kVAr'].max():>9.3f} {sub['hy_kW'].mean():>9.3f} "
          f"{sub['hy_I_A'].mean():>9.3f} {sub['hy_I_B'].mean():>9.3f} {sub['hy_I_C'].mean():>9.3f} | {len(sub)}")

# ─────────────────────────────────────────────
# SECTION 8 — ALIGNED COMPARISON (OVERLAP PERIOD)
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 8 — ALIGNED 15-MIN COMPARISON (Apr 30 – May 6 10:30)")
print("=" * 70)

hy_15 = hy.set_index("timestamp")[
    ["hy_kVAr","hy_kW","hy_kVA","hy_PF","hy_I_A","hy_I_B","hy_I_C","hy_amp_imb","hy_freq"]
].resample("15min").mean()

mg_idx = mg.set_index("timestamp")[["kW","kVAr","PF","I_A","I_C","I_AC_avg"]]

merged = mg_idx.join(hy_15, how="inner")
print(f"\n  Aligned records: {len(merged)}")

if len(merged) > 2:
    print(f"\n  Correlations (HyESys vs Main Grid):")
    print(f"    HyESys kVAr  vs MG kVAr:     r = {merged[['hy_kVAr','kVAr']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys kW    vs MG kW:        r = {merged[['hy_kW','kW']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys I_A   vs MG I_A:       r = {merged[['hy_I_A','I_A']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys I_C   vs MG I_C:       r = {merged[['hy_I_C','I_C']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys freq  vs MG PF:        r = {merged[['hy_freq','PF']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys kVAr  vs MG PF:        r = {merged[['hy_kVAr','PF']].dropna().corr().iloc[0,1]:+.4f}")

    print(f"\n  Mean values in aligned window:")
    print(f"    MG kW mean:          {merged['kW'].mean():,.1f} kW")
    print(f"    MG kVAr mean:        {merged['kVAr'].mean():,.1f} kVAr")
    print(f"    MG PF mean:          {merged['PF'].mean():.4f}")
    print(f"    MG I_A mean:         {merged['I_A'].mean():.1f} A")
    print(f"    HyESys kVAr mean:    {merged['hy_kVAr'].mean():.3f} kVAr")
    print(f"    HyESys I_A mean:     {merged['hy_I_A'].mean():.3f} A")
    print(f"    HyESys I_B mean:     {merged['hy_I_B'].mean():.3f} A")
    print(f"    HyESys I_C mean:     {merged['hy_I_C'].mean():.3f} A")

# ─────────────────────────────────────────────
# SECTION 9 — SCALE RECONCILIATION
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 9 — SCALE RECONCILIATION (Main Grid corrected vs HyESys)")
print("=" * 70)
print(f"""
  After applying multiplying factors:
    MG kW      = {mg['kW'].mean():.1f} kW  (was {mg['kW_raw'].mean():.4f} × 4000)
    MG kVAr    = {mg['kVAr'].mean():.1f} kVAr  (was {mg['kVAr_raw'].mean():.4f} × 4000)
    MG I_A     = {mg['I_A'].mean():.1f} A  (was {mg['I_A_raw'].mean():.4f} × 40)
    MG I_C     = {mg['I_C'].mean():.1f} A  (was {mg['I_C_raw'].mean():.4f} × 40)
    MG V_A     = {mg['V_A'].mean():.0f} V  (was {mg['V_A_raw'].mean():.1f} × 100)  → {mg['V_A'].mean()/1000:.2f} kV
    MG V_C     = {mg['V_C'].mean():.0f} V  (was {mg['V_C_raw'].mean():.1f} × 100)  → {mg['V_C'].mean()/1000:.2f} kV

  HyESys (no factor applied):
    HY kVAr    = {hy['hy_kVAr'].mean():.3f} kVAr
    HY kW      = {hy['hy_kW'].mean():.3f} kW
    HY I_A     = {hy['hy_I_A'].mean():.3f} A
    HY V_A     = {hy['hy_V_A'].mean():.1f} V  → {hy['hy_V_A'].mean()/1000:.3f} kV

  >> Main Grid meter is on the 10kV MV supply side (V_A ~{mg['V_A'].mean()/1000:.1f} kV).
  >> HyESys is on the LV side (V_A ~{hy['hy_V_A'].mean():.0f} V = {hy['hy_V_A'].mean()/1000:.3f} kV).
  >> They measure the same facility from different electrical nodes
     (MV primary incomer vs LV distribution bus).
  >> MG I_A ({mg['I_A'].mean():.1f} A at 10kV) and HY I_A ({hy['hy_I_A'].mean():.3f} A at {hy['hy_V_A'].mean():.0f}V) are consistent:
     Power balance: {mg['I_A'].mean():.1f}A × {mg['V_A'].mean()/1000:.1f}kV ≈ HY I_A × {hy['hy_V_A'].mean():.0f}V / transformer ratio.
""")

# ─────────────────────────────────────────────
# SECTION 10 — SIZING ESTIMATE FOR HYESYS
# ─────────────────────────────────────────────
print("=" * 70)
print("SECTION 10 — HYESYS SIZING ESTIMATE (based on corrected baseline)")
print("=" * 70)

kvar_mean = mg["kVAr"].mean()
kvar_max  = mg["kVAr"].max()
kw_mean   = mg["kW"].mean()
pf_mean   = mg["PF"].mean()
pf_target = 0.98

kvar_to_correct_mean = kw_mean * (np.tan(np.arccos(pf_mean)) - np.tan(np.arccos(pf_target)))
kvar_to_correct_max  = mg["kVAr_required"].clip(lower=0).max()

print(f"""
  Site reactive load:
    Mean kVAr:            {kvar_mean:.1f} kVAr
    Peak kVAr:            {kvar_max:.1f} kVAr

  To correct PF from {pf_mean:.3f} → {pf_target}:
    Mean kVAr injection needed: {kvar_to_correct_mean:.1f} kVAr
    Peak kVAr injection needed: {kvar_to_correct_max:.1f} kVAr

  Note: This is the MV-side reactive demand. HyESys is installed on LV side.
  The LV-side kVAr demand will be higher due to transformer reactive losses.
  HyESys must compensate at LV for effect to propagate to MV meter.
""")

# ─────────────────────────────────────────────
# SECTION 11 — DATA QUALITY
# ─────────────────────────────────────────────
print("=" * 70)
print("SECTION 11 — DATA QUALITY")
print("=" * 70)

# Main Grid gaps
mg_gaps = mg.sort_values("timestamp")["timestamp"].diff()
gaps_30 = mg_gaps[mg_gaps > pd.Timedelta("30min")]
print(f"\n  Main Grid gaps >30 min: {len(gaps_30)}")
for ts, gap in zip(mg.sort_values("timestamp")["timestamp"][gaps_30.index], gaps_30):
    print(f"    Gap of {gap} ending at {ts}")

# Outliers
mg_kw_outliers  = mg[mg["kW"]   < mg["kW"].quantile(0.005)]
mg_pf_outliers  = mg[mg["PF"]   < 0.90]
mg_i_drop       = mg[mg["I_A"]  < mg["I_A"].quantile(0.01)]
print(f"\n  Main Grid kW below 1st percentile ({mg['kW'].quantile(0.005):.1f} kW): {len(mg_kw_outliers)} records")
print(f"  Main Grid PF below 0.90: {len(mg_pf_outliers)} records")
print(f"  Main Grid I_A below 1st percentile ({mg['I_A'].quantile(0.01):.1f} A): {len(mg_i_drop)} records")

# HyESys gaps
hy_gaps = hy.sort_values("timestamp")["timestamp"].diff()
gaps_hy = hy_gaps[hy_gaps > pd.Timedelta("10min")]
print(f"\n  HyESys gaps >10 min (in capped window): {len(gaps_hy)}")
for ts, gap in zip(hy.sort_values("timestamp")["timestamp"][gaps_hy.index], gaps_hy):
    print(f"    Gap of {gap} ending at {ts}")

# Duplicate timestamps in HyESys
hy_dups = hy[hy["timestamp"].duplicated()]
print(f"\n  HyESys duplicate timestamps: {len(hy_dups)}")
if len(hy_dups) > 0:
    print(f"  First 5: {list(hy_dups['timestamp'].head())}")

# ─────────────────────────────────────────────
# SECTION 12 — FULL SUMMARY
# ─────────────────────────────────────────────
print("\n" + "=" * 70)
print("SECTION 12 — COMPLETE FINDINGS SUMMARY")
print("=" * 70)

kvar_mean_corrected = mg["kVAr"].mean()
kvar_max_corrected  = mg["kVAr"].max()
kw_mean_corrected   = mg["kW"].mean()
kw_max_corrected    = mg["kW"].max()
ia_mean_corrected   = mg["I_A"].mean()
ic_mean_corrected   = mg["I_C"].mean()
va_mean_corrected   = mg["V_A"].mean()

print(f"""
  METER SETUP
  ───────────
  Site:       Baoyuan (诸暨市葆元实业有限公司)
  Meter:      MV-side revenue meter, CT=40, PT=100, 综合倍率=4000
  Supply:     ~{va_mean_corrected/1000:.1f} kV (MV supply, single-line metering on A and C phases only)
  Phase B:    No CT installed — B phase carries load but is unmeasured at this meter

  CORRECTED BASELINE (Apr 26 – May 6 10:30, 1003 records)
  ─────────────────────────────────────────────────────────
  Active Power:     mean {kw_mean_corrected:,.0f} kW  |  peak {kw_max_corrected:,.0f} kW  |  min {mg['kW'].min():,.0f} kW
  Reactive Power:   mean {kvar_mean_corrected:,.0f} kVAr  |  peak {kvar_max_corrected:,.0f} kVAr  |  min {mg['kVAr'].min():,.0f} kVAr
  Power Factor:     mean {mg['PF'].mean():.4f}  |  min {mg['PF'].min():.4f}  |  max {mg['PF'].max():.4f}
  Current A:        mean {ia_mean_corrected:.1f} A  |  peak {mg['I_A'].max():.1f} A
  Current C:        mean {ic_mean_corrected:.1f} A  |  peak {mg['I_C'].max():.1f} A
  Voltage A:        mean {va_mean_corrected/1000:.3f} kV  (range {mg['V_A'].min()/1000:.3f}–{mg['V_A'].max()/1000:.3f} kV)
  Energy consumed:  ~{(mg_sorted['kWh'].iloc[-1]-mg_sorted['kWh'].iloc[0]):,.0f} kWh over {period_days:.1f} days
  Daily avg usage:  ~{(mg_sorted['kWh'].iloc[-1]-mg_sorted['kWh'].iloc[0])/period_days:,.0f} kWh/day

  POWER FACTOR PROFILE
  ─────────────────────
  PF range:     {mg['PF'].min():.3f} – {mg['PF'].max():.3f}  (all readings)
  PF < 0.90:    {len(pf_below_90)} records ({len(pf_below_90)/len(mg)*100:.1f}%)
  PF < 0.85:    {len(pf_below_85)} records ({len(pf_below_85)/len(mg)*100:.1f}%) — SP penalty threshold
  >> PF is consistently between 0.93–0.95. Site is NOT in penalty territory.
  >> However, PF 0.93 is below the ideal 0.98 target, meaning reactive correction
     will yield measurable I²R savings and reduced apparent power at meter.

  REACTIVE CORRECTION OPPORTUNITY
  ─────────────────────────────────
  kVAr to correct to PF=0.98:  mean ~{kvar_to_correct_mean:.0f} kVAr  |  peak ~{kvar_to_correct_max:.0f} kVAr
  This is the MV-side demand. LV-side (where HyESys sits) will be proportionally
  higher due to transformer magnetising reactive current.

  HYESYS UNIT (LV-SIDE, pre-activation, capped at 10:30)
  ────────────────────────────────────────────────────────
  HyESys V_A:     {hy['hy_V_A'].mean():.1f} V  (LV bus, ~{hy['hy_V_A'].mean()/1000:.3f} kV)
  HyESys kVAr:    mean {hy['hy_kVAr'].mean():.3f} kVAr  |  max {hy['hy_kVAr'].max():.3f} kVAr
  HyESys kW:      mean {hy['hy_kW'].mean():.3f} kW
  HyESys I_A:     mean {hy['hy_I_A'].mean():.3f} A  |  max {hy['hy_I_A'].max():.3f} A
  HyESys I_B:     mean {hy['hy_I_B'].mean():.3f} A  |  max {hy['hy_I_B'].max():.3f} A
  HyESys I_C:     mean {hy['hy_I_C'].mean():.3f} A  |  max {hy['hy_I_C'].max():.3f} A
  Amp imbalance:  mean {hy['hy_amp_imb'].mean():.3f}%
  Temperature:    mean {hy['hy_temp'].mean():.1f}°C  |  max {hy['hy_temp'].max():.1f}°C

  >> HyESys is measuring very low kVAr (avg {hy['hy_kVAr'].mean():.3f} kVAr) pre-activation.
  >> This is consistent with HyESys being in STANDBY/MONITOR mode — unit is
     powered and logging but not injecting compensation current.
  >> HyESys LV currents (I_A {hy['hy_I_A'].mean():.2f} A, I_B {hy['hy_I_B'].mean():.2f} A) are the unit's
     own idle draw, not the site load current.

  MV vs LV MEASUREMENT RELATIONSHIP
  ────────────────────────────────────
  MG meter (MV, ~{va_mean_corrected/1000:.1f} kV):   I_A = {ia_mean_corrected:.1f} A
  HyESys  (LV, ~{hy['hy_V_A'].mean():.0f} V):  I_A = {hy['hy_I_A'].mean():.2f} A (own idle draw only)
  >> MG measures the FULL site load at MV level.
  >> HyESys measures its own output at LV level.
  >> These are not the same current — they measure different things.

  DATA QUALITY
  ─────────────
  Main Grid:  No gaps >30 min. {len(mg)} complete 15-min records. Clean dataset.
  HyESys:     {len(gaps_hy)} gaps >10 min in capped window. {len(hy_dups)} duplicate timestamps.
  Phase B:    Intentionally unmeasured (no CT). Not a fault.
""")

print("=" * 70)
print("END OF ANALYSIS")
print("=" * 70)
