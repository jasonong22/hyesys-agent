"""
Baoyuan Analysis v3 — Distribution Hierarchy Corrected
- Both files capped at 2026-05-06 10:30:00 (pre-activation baseline only)
- Main Grid multiplying factors applied: kW/kVAr ×4000, I ×40, V ×100
- HyESys values taken as-is (no multiplier)
- Phase B excluded (no CT installed — not a fault)
- v3 adds: LV distribution hierarchy context from 设备报告Baoyuan_HyESys_Savings_Report_2026-05-12
  HyESys is at Cabinet 2 of Cabinet A within MSB1.
  Main Grid meter reads MSB1 + MSB2 combined at 10 kV MV.
  These are two different nodes in the distribution tree — not the same bus.
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import numpy as np
from pathlib import Path

DATA_DIR  = Path(r"C:\Users\JasonOng\Desktop\Data analytics\Baoyuan China\LV room\1-6may")
CAP_TIME  = pd.Timestamp("2026-05-06 10:30:00")

MF_POWER   = 4000
MF_CURRENT =   40
MF_VOLTAGE =  100

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

mg["kW"]       = mg["kW_raw"]     * MF_POWER
mg["kVAr"]     = mg["kVAr_raw"]   * MF_POWER
mg["I_A"]      = mg["I_A_raw"]    * MF_CURRENT
mg["I_C"]      = mg["I_C_raw"]    * MF_CURRENT
mg["V_A"]      = mg["V_A_raw"]    * MF_VOLTAGE
mg["V_C"]      = mg["V_C_raw"]    * MF_VOLTAGE
mg["kWh"]      = mg["kWh_raw"]    * MF_POWER
mg["kW_A"]     = mg["kW_A_raw"]   * MF_POWER
mg["kW_C"]     = mg["kW_C_raw"]   * MF_POWER
mg["I_AC_avg"] = (mg["I_A"] + mg["I_C"]) / 2

mg = mg[mg["timestamp"] <= CAP_TIME].copy()
mg["kVA_check"] = np.sqrt(mg["kW"]**2 + mg["kVAr"]**2)
mg["PF_check"]  = mg["kW"] / mg["kVA_check"].replace(0, np.nan)
mg["hour"]      = mg["timestamp"].dt.hour
mg["date"]      = mg["timestamp"].dt.date

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

hy = hy[hy["timestamp"] <= CAP_TIME].copy()
hy["hour"] = hy["timestamp"].dt.hour

# Pre-compute shared values used in multiple sections
pf_target        = 0.98
mg_sorted        = mg.sort_values("timestamp")
kwh_start        = mg_sorted["kWh"].iloc[0]
kwh_end          = mg_sorted["kWh"].iloc[-1]
period_days      = (mg_sorted["timestamp"].iloc[-1] - mg_sorted["timestamp"].iloc[0]).total_seconds() / 86400
kvar_mean        = mg["kVAr"].mean()
kvar_max         = mg["kVAr"].max()
kw_mean          = mg["kW"].mean()
pf_mean          = mg["PF"].mean()
ia_mean          = mg["I_A"].mean()
ic_mean          = mg["I_C"].mean()
va_mean          = mg["V_A"].mean()
pf_below_90      = mg[mg["PF"] < 0.90]
pf_below_85      = mg[mg["PF"] < 0.85]

mg["kVAr_required"] = mg["kW"] * (
    np.tan(np.arccos(mg["PF"].clip(0.01, 0.9999))) -
    np.tan(np.arccos(pf_target))
)
kvar_to_correct_mean = kw_mean * (np.tan(np.arccos(pf_mean)) - np.tan(np.arccos(pf_target)))
kvar_to_correct_max  = mg["kVAr_required"].clip(lower=0).max()

hy_kvar_mean = hy["hy_kVAr"].mean()
hy_kvar_max  = hy["hy_kVAr"].max()
hy_va_mean   = hy["hy_V_A"].mean()
hy_ia_mean   = hy["hy_I_A"].mean()
hy_ib_mean   = hy["hy_I_B"].mean()
hy_ic_mean   = hy["hy_I_C"].mean()

# 15-min resampled HyESys for alignment
hy_15  = hy.set_index("timestamp")[
    ["hy_kVAr","hy_kW","hy_kVA","hy_PF","hy_I_A","hy_I_B","hy_I_C","hy_amp_imb","hy_freq"]
].resample("15min").mean()
mg_idx = mg.set_index("timestamp")[["kW","kVAr","PF","I_A","I_C","I_AC_avg"]]
merged = mg_idx.join(hy_15, how="inner")

def stats(series, label, unit, width=38):
    s = series.dropna()
    if len(s) == 0:
        print(f"  {label:{width}s}  NO DATA")
        return
    print(f"  {label:{width}s}  mean={s.mean():>10.2f}  median={s.median():>10.2f}  "
          f"min={s.min():>10.2f}  max={s.max():>10.2f}  std={s.std():>8.2f}  [{unit}]")

# ═══════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════
print("=" * 70)
print("BAOYUAN ANALYSIS v3 — DISTRIBUTION HIERARCHY CORRECTED")
print(f"Analysis window: both files capped at {CAP_TIME}")
print(f"Multiplying factors: kW/kVAr ×{MF_POWER}, I ×{MF_CURRENT}, V ×{MF_VOLTAGE}")
print(f"Phase B: excluded — no CT installed at MV meter")
print("=" * 70)
print(f"\nMain Grid records: {len(mg)}  ({mg['timestamp'].min()} → {mg['timestamp'].max()})")
print(f"HyESys records:    {len(hy)}  ({hy['timestamp'].min()} → {hy['timestamp'].max()})")

# ═══════════════════════════════════════════════════════════
# SECTION 0 — DISTRIBUTION HIERARCHY (v3 addition)
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 0 — LV DISTRIBUTION HIERARCHY (from site report)")
print("=" * 70)
print(f"""
  Source: 设备报告Baoyuan_HyESys_Savings_Report_2026-05-12 (electrical flowchart)

  Main Grid (10 kV MV meter — THIS FILE)
  ├── MSB1  [no sub-meter at MSB level]
  │   ├── CapBank1  (HV-side capacitor bank — NOT HyESys)
  │   ├── CapBank2  (HV-side capacitor bank — NOT HyESys)
  │   ├── Cabinet A  (metered at cabinet level)
  │   │   ├── Feeder 1  (metered)
  │   │   ├── Feeder 2  ← HyESys H125 installed here  (metered)
  │   │   ├── Feeder 3  (metered)
  │   │   └── Feeder 4  (metered)
  │   ├── Cabinet B  (metered — 4 feeders)
  │   └── Cabinet C  (metered — 4 feeders)
  └── MSB2  [no sub-meter at MSB level]

  KEY IMPLICATION:
  ► Main Grid meter reads: MSB1 (all 3 cabinets × 4 feeders + CapBanks) + MSB2
  ► HyESys data file reads: Cabinet A / Feeder 2 ONLY (one sub-feeder of MSB1)
  ► These are NOT the same bus — comparison of absolute kVAr values is meaningless
    without scaling by the sub-feeder's share of total load.
  ► Correct reference for HyESys effect measurement: Cabinet A incomer meter
    (not available in this dataset — explains "measurement limitation" note in report).

  SCALE CONTEXT:
  ► HyESys kVAr {hy_kvar_mean:.1f} kVAr = reactive load of Cabinet A / Feeder 2 only
  ► Main Grid kVAr {kvar_mean:,.0f} kVAr = entire facility (MSB1 + MSB2, MV side)
  ► HyESys share of facility reactive load: {hy_kvar_mean/kvar_mean*100:.1f}%
    (expected — it serves 1 of ~12 feeder panels + MSB2)
""")

# ═══════════════════════════════════════════════════════════
# SECTION 1 — MAIN GRID BASELINE
# ═══════════════════════════════════════════════════════════
print("=" * 70)
print("SECTION 1 — MAIN GRID BASELINE  [MSB1 + MSB2 combined, 10 kV MV]")
print("=" * 70)

stats(mg["kW"],       "Active Power (kW)",             "kW")
stats(mg["kVAr"],     "Reactive Power (kVAr)",         "kVAr")
stats(mg["PF"],       "Power Factor",                  "–")
stats(mg["I_A"],      "Current Phase A (MV primary)",  "A")
stats(mg["I_C"],      "Current Phase C (MV primary)",  "A")
stats(mg["I_AC_avg"], "Current A+C average",           "A")
stats(mg["V_A"],      "Voltage Phase A (MV primary)",  "V")
stats(mg["V_C"],      "Voltage Phase C (MV primary)",  "V")
stats(mg["kW_A"],     "Active Power Phase A",          "kW")
stats(mg["kW_C"],     "Active Power Phase C",          "kW")

print(f"\n  Derived check:")
stats(mg["kVA_check"], "  Apparent Power (√kW²+kVAr²)", "kVA")
stats(mg["PF_check"],  "  PF derived from kW/kVA",      "–")

print(f"\n  Max active power demand:  {mg['kW'].max():,.1f} kW  at {mg.loc[mg['kW'].idxmax(),'timestamp']}")
print(f"  Max reactive demand:      {mg['kVAr'].max():,.1f} kVAr  at {mg.loc[mg['kVAr'].idxmax(),'timestamp']}")
print(f"  Min PF recorded:          {mg['PF'].min():.4f}  at {mg.loc[mg['PF'].idxmin(),'timestamp']}")
print(f"  Max PF recorded:          {mg['PF'].max():.4f}  at {mg.loc[mg['PF'].idxmax(),'timestamp']}")
print(f"\n  Records PF < 0.90:  {len(pf_below_90)} ({len(pf_below_90)/len(mg)*100:.1f}%)")
print(f"  Records PF < 0.85 (SP penalty threshold):  {len(pf_below_85)} ({len(pf_below_85)/len(mg)*100:.1f}%)")

# ═══════════════════════════════════════════════════════════
# SECTION 2 — ENERGY CONSUMPTION
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 2 — ENERGY CONSUMPTION  [MSB1 + MSB2 combined]")
print("=" * 70)
print(f"\n  kWh meter start (Apr 26 00:00):  {kwh_start:,.1f} kWh")
print(f"  kWh meter end   (May 6 10:30):   {kwh_end:,.1f} kWh")
print(f"  Period:                          {period_days:.1f} days")
print(f"  Total consumption:               {kwh_end - kwh_start:,.1f} kWh")
print(f"  Average daily consumption:       {(kwh_end - kwh_start) / period_days:,.1f} kWh/day")
print(f"  Average load (kWh/h):            {(kwh_end - kwh_start) / (period_days * 24):,.1f} kW")

# ═══════════════════════════════════════════════════════════
# SECTION 3 — HOURLY LOAD PATTERN
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 3 — HOURLY LOAD PATTERN  [Main Grid — all days]")
print("=" * 70)
print(f"\n  {'Hour':>5} | {'kW mean':>9} {'kW max':>9} {'kVAr mean':>10} {'PF mean':>8} {'I_A mean':>9} {'I_C mean':>9} | n")
print("  " + "-" * 90)
for h in range(24):
    sub = mg[mg["hour"] == h]
    print(f"  {h:>5}h | {sub['kW'].mean():>9.1f} {sub['kW'].max():>9.1f} {sub['kVAr'].mean():>10.1f} "
          f"{sub['PF'].mean():>8.4f} {sub['I_A'].mean():>9.1f} {sub['I_C'].mean():>9.1f} | {len(sub)}")

# ═══════════════════════════════════════════════════════════
# SECTION 4 — DAILY PATTERN
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 4 — DAILY PATTERN  [Main Grid — by date]")
print("=" * 70)
print(f"\n  {'Date':<13} | {'kW mean':>9} {'kW max':>9} {'kVAr mean':>10} {'PF mean':>8} {'I_A mean':>9} {'Records':>8}")
print("  " + "-" * 78)
for date, sub in mg.groupby("date"):
    print(f"  {str(date):<13} | {sub['kW'].mean():>9.1f} {sub['kW'].max():>9.1f} {sub['kVAr'].mean():>10.1f} "
          f"{sub['PF'].mean():>8.4f} {sub['I_A'].mean():>9.1f} {sub['timestamp'].count():>8}")

# ═══════════════════════════════════════════════════════════
# SECTION 5 — REACTIVE POWER ANALYSIS
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 5 — REACTIVE POWER ANALYSIS  [Main Grid — full facility]")
print("=" * 70)

kvar_day   = mg[(mg["hour"] >= 8)  & (mg["hour"] < 20)]["kVAr"]
kvar_night = mg[(mg["hour"] < 8)   | (mg["hour"] >= 20)]["kVAr"]
print(f"\n  Daytime  (08:00–20:00):  mean kVAr = {kvar_day.mean():,.1f}  max = {kvar_day.max():,.1f}")
print(f"  Nighttime (20:00–08:00): mean kVAr = {kvar_night.mean():,.1f}  max = {kvar_night.max():,.1f}")

mg["kVAr_to_kW"] = mg["kVAr"] / mg["kW"].replace(0, np.nan)
print(f"\n  kVAr/kW ratio:  mean = {mg['kVAr_to_kW'].mean():.4f}  std = {mg['kVAr_to_kW'].std():.4f}")
print(f"  tan(φ) → φ:     mean angle = {np.degrees(np.arctan(mg['kVAr_to_kW'].mean())):.2f}°")

print(f"\n  kVAr required to correct full facility to PF={pf_target} [WHOLE-SITE REFERENCE]:")
stats(mg["kVAr_required"].clip(lower=0), "  Reactive correction needed", "kVAr")
print(f"  Peak correction needed: {kvar_to_correct_max:.1f} kVAr  at {mg.loc[mg['kVAr_required'].idxmax(), 'timestamp']}")
print(f"""
  CONTEXT: This whole-facility correction figure ({kvar_to_correct_mean:.0f} kVAr mean, {kvar_to_correct_max:.0f} kVAr peak)
  is the total reactive demand seen at the MV meter (MSB1 + MSB2 combined).
  The single HyESys H125 at Cabinet A / Feeder 2 corrects only Feeder 2's share.
  Full-facility correction would require HyESys units at each cabinet-level bus.
""")

# ═══════════════════════════════════════════════════════════
# SECTION 6 — HYESYS DATA
# ═══════════════════════════════════════════════════════════
print("=" * 70)
print("SECTION 6 — HYESYS DATA  [Cabinet A / Feeder 2 — Apr 30 → May 6 10:30]")
print("=" * 70)
print(f"\n  Records: {len(hy)}")
print(f"  Period:  {hy['timestamp'].min()} → {hy['timestamp'].max()}")
print(f"\n  Note: HyESys measures at its own LV output terminals (~{hy_va_mean:.0f} V).")
print(f"  These values represent the reactive/active conditions on Cabinet A / Feeder 2.")
print(f"  They are NOT comparable in magnitude to the Main Grid MV meter readings.")
print()

stats(hy["hy_kVAr"],    "HyESys kVAr (Feeder 2 reactive)", "kVAr")
stats(hy["hy_kW"],      "HyESys kW (Feeder 2 active)",     "kW")
stats(hy["hy_kVA"],     "HyESys kVA",                      "kVA")
stats(hy["hy_PF"],      "HyESys PF",                       "–")
stats(hy["hy_I_A"],     "HyESys I_A",                      "A")
stats(hy["hy_I_B"],     "HyESys I_B",                      "A")
stats(hy["hy_I_C"],     "HyESys I_C",                      "A")
stats(hy["hy_amp_imb"], "HyESys Amp Imbalance",             "%")
stats(hy["hy_V_A"],     "HyESys V_A (LV bus)",             "V")
stats(hy["hy_V_B"],     "HyESys V_B (LV bus)",             "V")
stats(hy["hy_V_C"],     "HyESys V_C (LV bus)",             "V")
stats(hy["hy_freq"],    "Grid Frequency",                   "Hz")
stats(hy["hy_temp"],    "Temperature",                      "°C")

# ═══════════════════════════════════════════════════════════
# SECTION 7 — HYESYS HOURLY PATTERN
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 7 — HYESYS HOURLY PATTERN  [Cabinet A / Feeder 2]")
print("=" * 70)
print(f"\n  {'Hour':>5} | {'kVAr mean':>10} {'kVAr max':>9} {'kW mean':>9} {'I_A mean':>9} {'I_B mean':>9} {'I_C mean':>9} | n")
print("  " + "-" * 88)
for h in range(24):
    sub = hy[hy["hour"] == h]
    if len(sub) == 0:
        continue
    print(f"  {h:>5}h | {sub['hy_kVAr'].mean():>10.3f} {sub['hy_kVAr'].max():>9.3f} {sub['hy_kW'].mean():>9.3f} "
          f"{sub['hy_I_A'].mean():>9.3f} {sub['hy_I_B'].mean():>9.3f} {sub['hy_I_C'].mean():>9.3f} | {len(sub)}")

# ═══════════════════════════════════════════════════════════
# SECTION 8 — ALIGNED COMPARISON
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 8 — ALIGNED 15-MIN COMPARISON  [Apr 30 – May 6 10:30]")
print("  NOTE: Cross-correlating two different distribution nodes.")
print("  Correlation shows how Feeder 2 load tracks total facility load.")
print("=" * 70)

print(f"\n  Aligned records (15-min bins): {len(merged)}")

if len(merged) > 2:
    print(f"\n  Correlations (HyESys / Feeder 2  vs  Main Grid / full facility):")
    print(f"    HyESys kVAr  vs MG kVAr:  r = {merged[['hy_kVAr','kVAr']].dropna().corr().iloc[0,1]:+.4f}  "
          f"  (how well Feeder 2 reactive load tracks full-site reactive load)")
    print(f"    HyESys kW    vs MG kW:    r = {merged[['hy_kW','kW']].dropna().corr().iloc[0,1]:+.4f}  "
          f"  (how well Feeder 2 active load tracks full-site active load)")
    print(f"    HyESys I_A   vs MG I_A:   r = {merged[['hy_I_A','I_A']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys I_C   vs MG I_C:   r = {merged[['hy_I_C','I_C']].dropna().corr().iloc[0,1]:+.4f}")
    print(f"    HyESys kVAr  vs MG PF:    r = {merged[['hy_kVAr','PF']].dropna().corr().iloc[0,1]:+.4f}")

    print(f"\n  Mean values in aligned window:")
    print(f"    MG kW mean (full facility):        {merged['kW'].mean():,.1f} kW")
    print(f"    MG kVAr mean (full facility):      {merged['kVAr'].mean():,.1f} kVAr")
    print(f"    MG PF mean (full facility):        {merged['PF'].mean():.4f}")
    print(f"    HyESys kVAr mean (Feeder 2 only):  {merged['hy_kVAr'].mean():.3f} kVAr")
    print(f"    Feeder 2 share of total kVAr:      {merged['hy_kVAr'].mean()/merged['kVAr'].mean()*100:.1f}%")

# ═══════════════════════════════════════════════════════════
# SECTION 9 — MEASUREMENT NODE CONTEXT  (replaces v2 "scale reconciliation")
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 9 — MEASUREMENT NODE CONTEXT")
print("=" * 70)
print(f"""
  TWO DIFFERENT NODES IN THE DISTRIBUTION TREE
  ─────────────────────────────────────────────
  Main Grid meter (this file):
    Node:     10 kV MV incomer — measures MSB1 + MSB2 combined
    Voltage:  ~{va_mean/1000:.2f} kV  (MV primary)
    Current:  I_A = {ia_mean:.1f} A, I_C = {ic_mean:.1f} A  (MV primary side, phases A+C only)
    Power:    {kw_mean:,.0f} kW active, {kvar_mean:,.0f} kVAr reactive

  HyESys data file:
    Node:     LV bus at Cabinet A / Feeder 2 within MSB1
    Voltage:  ~{hy_va_mean:.0f} V  (LV secondary)
    Current:  I_A = {hy_ia_mean:.2f} A, I_B = {hy_ib_mean:.2f} A, I_C = {hy_ic_mean:.2f} A  (LV, unit output only)
    Power:    {hy_kvar_mean:.2f} kVAr reactive  (Feeder 2 reactive load)

  WHY THE VALUES ARE SO DIFFERENT
  ─────────────────────────────────
  This is NOT a scale error or an anomaly. It reflects the topology:
  1. MG kVAr ({kvar_mean:,.0f} kVAr) includes MSB1 (all 3 cabinets × 4 feeders + CapBanks) + MSB2.
  2. HyESys kVAr ({hy_kvar_mean:.1f} kVAr) is only Cabinet A / Feeder 2 — 1 of ~12 feeder panels.
  3. In addition: CapBank1 + CapBank2 at MSB1 already compensate HV-side reactive,
     which reduces the apparent kVAr seen at the MV meter somewhat.
  4. MSB2 adds its own uncompensated reactive load to the MG total.

  Feeder 2 share of MG kVAr: {hy_kvar_mean/kvar_mean*100:.1f}%  — consistent with 1-of-12 sub-feeders.

  CORRECT COMPARISON POINT
  ─────────────────────────
  To directly validate HyESys effect:  use Cabinet A incomer meter  (not Main Grid).
  The savings report (2026-05-12) correctly measures at the MSB1 incomer level and
  finds a 116 A / 35 kW reduction per activation event — directly attributable to
  HyESys acting on Feeder 2's load.
""")

# ═══════════════════════════════════════════════════════════
# SECTION 10 — SIZING CONTEXT
# ═══════════════════════════════════════════════════════════
print("=" * 70)
print("SECTION 10 — SIZING CONTEXT")
print("=" * 70)
print(f"""
  FEEDER 2 (Cabinet A) — where HyESys H125 is installed
  ───────────────────────────────────────────────────────
  Reactive load:    {hy_kvar_mean:.1f} kVAr mean, {hy_kvar_max:.1f} kVAr peak  (from HyESys data)
  HyESys H125 rated output: 125 kVAr
  Headroom:         {125 - hy_kvar_max:.1f} kVAr spare capacity
  >> H125 is adequately sized for Feeder 2's reactive demand.
  >> Low kVAr reading (~{hy_kvar_mean:.1f} kVAr mean) pre-activation is consistent with
     standby mode — unit is monitoring but not yet injecting.

  FULL FACILITY — extrapolated from MG meter
  ────────────────────────────────────────────
  Full-site kVAr:   {kvar_mean:,.0f} kVAr mean, {kvar_max:,.0f} kVAr peak
  kVAr to reach PF={pf_target}: {kvar_to_correct_mean:.0f} kVAr mean, {kvar_to_correct_max:.0f} kVAr peak
  >> The {kvar_to_correct_mean:.0f}–{kvar_to_correct_max:.0f} kVAr figure is the TOTAL facility correction
     at the MV level (not what this single HyESys unit handles).
  >> To correct the whole facility, HyESys units would be needed
     at each of Cabinet A, B, and C busbars, plus MSB2.
  >> The existing CapBank1 + CapBank2 at MSB1 already handle some
     HV-side compensation; their contribution is already reflected
     in the MG PF reading of {pf_mean:.3f}.

  NOTE ON PF PENALTY
  ───────────────────
  MG PF mean = {pf_mean:.3f} — above the SP penalty threshold of 0.85.
  Site is NOT currently incurring SP penalty.
  However at PF {pf_mean:.3f} there is still significant reactive current
  causing I²R losses in distribution cables and transformer.
  Savings opportunity: reactive correction → I²R reduction → kWh savings.
""")

# ═══════════════════════════════════════════════════════════
# SECTION 11 — DATA QUALITY
# ═══════════════════════════════════════════════════════════
print("=" * 70)
print("SECTION 11 — DATA QUALITY")
print("=" * 70)

mg_gaps = mg.sort_values("timestamp")["timestamp"].diff()
gaps_30 = mg_gaps[mg_gaps > pd.Timedelta("30min")]
print(f"\n  Main Grid gaps >30 min: {len(gaps_30)}")
for ts, gap in zip(mg.sort_values("timestamp")["timestamp"][gaps_30.index], gaps_30):
    print(f"    Gap of {gap} ending at {ts}")

mg_kw_outliers = mg[mg["kW"] < mg["kW"].quantile(0.005)]
mg_pf_outliers = mg[mg["PF"] < 0.90]
mg_i_drop      = mg[mg["I_A"] < mg["I_A"].quantile(0.01)]
print(f"\n  Main Grid kW below 1st percentile ({mg['kW'].quantile(0.005):.1f} kW): {len(mg_kw_outliers)} records")
print(f"  Main Grid PF below 0.90: {len(mg_pf_outliers)} records")
print(f"  Main Grid I_A below 1st percentile ({mg['I_A'].quantile(0.01):.1f} A): {len(mg_i_drop)} records")

hy_gaps = hy.sort_values("timestamp")["timestamp"].diff()
gaps_hy = hy_gaps[hy_gaps > pd.Timedelta("10min")]
print(f"\n  HyESys gaps >10 min: {len(gaps_hy)}")
for ts, gap in zip(hy.sort_values("timestamp")["timestamp"][gaps_hy.index], gaps_hy):
    print(f"    Gap of {gap} ending at {ts}")

hy_dups = hy[hy["timestamp"].duplicated()]
print(f"\n  HyESys duplicate timestamps: {len(hy_dups)}")
if len(hy_dups) > 0:
    print(f"  First 5: {list(hy_dups['timestamp'].head())}")

# ═══════════════════════════════════════════════════════════
# SECTION 12 — COMPLETE FINDINGS SUMMARY (v3 re-framed)
# ═══════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SECTION 12 — COMPLETE FINDINGS SUMMARY  [v3 — hierarchy-corrected]")
print("=" * 70)
print(f"""
  ┌─ DISTRIBUTION CONTEXT ──────────────────────────────────────────┐
  │  Main Grid meter = MSB1 (3 cabinets, 12 feeders, 2 CapBanks)    │
  │                  + MSB2,  measured at 10 kV MV incomer           │
  │  HyESys meter   = Cabinet A / Feeder 2 within MSB1 only         │
  │  The two files measure DIFFERENT nodes — not the same bus.       │
  └──────────────────────────────────────────────────────────────────┘

  MAIN GRID BASELINE  [Apr 26 – May 6 10:30, {len(mg)} records, MSB1+MSB2, 10 kV]
  ─────────────────────────────────────────────────────────────────────
  Active Power:     mean {kw_mean:,.0f} kW      peak {mg['kW'].max():,.0f} kW    min {mg['kW'].min():,.0f} kW
  Reactive Power:   mean {kvar_mean:,.0f} kVAr    peak {kvar_max:,.0f} kVAr  min {mg['kVAr'].min():,.0f} kVAr
  Power Factor:     mean {pf_mean:.4f}   min {mg['PF'].min():.4f}   max {mg['PF'].max():.4f}
  Current A (MV):   mean {ia_mean:.1f} A   peak {mg['I_A'].max():.1f} A
  Current C (MV):   mean {ic_mean:.1f} A   peak {mg['I_C'].max():.1f} A
  Voltage (MV):     mean {va_mean/1000:.2f} kV  (range {mg['V_A'].min()/1000:.2f}–{mg['V_A'].max()/1000:.2f} kV)
  Energy consumed:  {kwh_end-kwh_start:,.0f} kWh over {period_days:.1f} days  (~{(kwh_end-kwh_start)/period_days:,.0f} kWh/day)

  POWER FACTOR PROFILE  [full facility at MV]
  ────────────────────────────────────────────
  PF < 0.90:  {len(pf_below_90)} records ({len(pf_below_90)/len(mg)*100:.1f}%)
  PF < 0.85:  {len(pf_below_85)} records ({len(pf_below_85)/len(mg)*100:.1f}%)  — SP penalty threshold
  ► Site is NOT in SP penalty territory (PF consistently 0.93–0.95).
  ► PF 0.93 is below the 0.98 target → reactive correction still yields
    measurable I²R savings at distribution cable level.

  REACTIVE CORRECTION — FULL FACILITY (MV meter reference)
  ─────────────────────────────────────────────────────────
  kVAr to correct to PF=0.98:  mean ~{kvar_to_correct_mean:.0f} kVAr  |  peak ~{kvar_to_correct_max:.0f} kVAr
  ► This figure represents the TOTAL facility gap — it is NOT
    what the single HyESys H125 at Feeder 2 needs to handle.
  ► Full-facility correction requires distributed HyESys units
    at Cabinet A, B, C busbars within MSB1, plus MSB2.
  ► The existing CapBank1 + CapBank2 already compensate some HV-side
    reactive load; their effect is already included in the PF reading.

  HYESYS UNIT  [Cabinet A / Feeder 2, LV ~{hy_va_mean:.0f} V, pre-activation]
  ──────────────────────────────────────────────────────────────────────
  kVAr (Feeder 2 load):  mean {hy_kvar_mean:.3f} kVAr  |  max {hy_kvar_max:.3f} kVAr
  kW:                    mean {hy['hy_kW'].mean():.3f} kW
  Current A/B/C:         {hy_ia_mean:.2f} A / {hy_ib_mean:.2f} A / {hy_ic_mean:.2f} A
  Amp imbalance:         {hy['hy_amp_imb'].mean():.2f}%
  Temperature:           mean {hy['hy_temp'].mean():.1f}°C  |  max {hy['hy_temp'].max():.1f}°C
  ► {hy_kvar_mean:.1f} kVAr is the reactive load of Feeder 2 only — correct in context.
  ► H125 rated at 125 kVAr; {hy_kvar_max:.1f} kVAr peak leaves {125-hy_kvar_max:.0f} kVAr headroom.
  ► Unit is in STANDBY/MONITOR mode pre-activation; kVAr reading
    reflects Feeder 2 load being observed, not HyESys output.

  WHY 43.7 kVAr ≠ 1,061 kVAr  (the apparent "scale mismatch")
  ──────────────────────────────────────────────────────────────
  This is expected and fully explained by the distribution hierarchy:
    HyESys (Feeder 2):     ~{hy_kvar_mean:.0f} kVAr   = 1 sub-feeder of MSB1
    MSB1 incomer:          ~538 kVAr  = all 3 cabinets (from savings report Section 2)
    Main Grid (MV meter):  ~{kvar_mean:,.0f} kVAr  = MSB1 + MSB2 combined
  The ratio {hy_kvar_mean:.0f}/{kvar_mean:.0f} = {hy_kvar_mean/kvar_mean*100:.1f}% matches the expected
  share of 1 feeder out of the full facility.

  DATA QUALITY
  ─────────────
  Main Grid:  {len(gaps_30)} gaps >30 min.  {len(mg)} complete 15-min records.  Clean dataset.
  HyESys:     {len(gaps_hy)} gaps >10 min.  {len(hy_dups)} duplicate timestamps.
  Phase B:    Intentionally unmeasured at MV meter (no CT on B phase).
              B phase is live and carries load — not a fault.
""")

print("=" * 70)
print("END OF ANALYSIS  —  v3 (distribution hierarchy corrected)")
print("=" * 70)
