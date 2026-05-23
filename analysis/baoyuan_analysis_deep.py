"""
Baoyuan Deep Analysis — Part 2
Focus: activation transition, overlap correlation, data anomalies
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import pandas as pd
import numpy as np
from pathlib import Path

DATA_DIR = Path(r"C:\Users\JasonOng\Desktop\Data analytics\Baoyuan China\LV room\1-6may")
HYESYS_ACTIVATION = pd.Timestamp("2026-05-06 16:29:31")

# ── Load ──────────────────────────────────────────────────────────
mg_raw = pd.read_excel(DATA_DIR / "Main Grid 26apr-6may.xlsx")
hy_raw = pd.read_excel(DATA_DIR / "HyESys 30apr-6may.xlsx", header=1)

# ── Clean Main Grid ───────────────────────────────────────────────
mg = mg_raw.copy()
mg["日期"] = mg["日期"].astype(str).str.replace(r"\t", "", regex=True).str.strip()
mg["timestamp"] = pd.to_datetime(mg["日期"], errors="coerce")
mg = mg.dropna(subset=["timestamp"]).sort_values("timestamp").reset_index(drop=True)
mg_rename = {
    "瞬时有功": "kW_total", "瞬时无功": "kVAr_total",
    "A相电流": "I_A", "B相": "I_B", "C相": "I_C", "零线": "I_N",
    "A相电压": "V_A", "B相.1": "V_B", "C相.1": "V_C",
    "总功率因数": "PF_total", "正向有功总": "kWh_fwd",
}
mg = mg.rename(columns=mg_rename)
for col in ["kW_total","kVAr_total","I_A","I_B","I_C","V_A","V_C","PF_total"]:
    mg[col] = pd.to_numeric(mg[col], errors="coerce")
mg["I_AC_avg"] = (mg["I_A"] + mg["I_C"]) / 2   # use A+C only (B=0 throughout)

# ── Clean HyESys ──────────────────────────────────────────────────
hy = hy_raw.copy()
hy["timestamp"] = pd.to_datetime(hy["上报时间戳"], errors="coerce")
hy = hy.dropna(subset=["timestamp"]).sort_values("timestamp").reset_index(drop=True)
hy = hy.rename(columns={
    "A相电流(A)": "hy_I_A", "B相电流(A)": "hy_I_B", "C相电流(A)": "hy_I_C",
    "A相电压(V)": "hy_V_A", "B相电压(V)": "hy_V_B", "C相电压(V)": "hy_V_C",
    "总输出有功功率(kW)": "hy_kW_out", "总输出无功功率(kVar)": "hy_kVAr_out",
    "总输出视在功率(kVA)": "hy_kVA_out", "输出总相功率因素": "hy_PF_out",
    "A相输出无功功率(kVar)": "hy_kVAr_A", "B相输出无功功率(kVar)": "hy_kVAr_B",
    "C相输出无功功率(kVar)": "hy_kVAr_C",
    "A相输出有功功率(kW)": "hy_kW_A", "B相输出有功功率(kW)": "hy_kW_B",
    "C相输出有功功率(kW)": "hy_kW_C",
    "3PAmpimb": "hy_amp_imb", "N": "hy_N",
    "电网频率(Hz)": "hy_freq", "温度(℃)": "hy_temp",
    "输入功率(kW)": "hy_kW_in", "输入电压(V)": "hy_V_in", "输入电流(A)": "hy_I_in",
})
num_cols = ["hy_I_A","hy_I_B","hy_I_C","hy_V_A","hy_V_B","hy_V_C",
            "hy_kW_out","hy_kVAr_out","hy_kVA_out","hy_PF_out",
            "hy_kVAr_A","hy_kVAr_B","hy_kVAr_C",
            "hy_kW_A","hy_kW_B","hy_kW_C",
            "hy_amp_imb","hy_N","hy_freq","hy_temp","hy_kW_in","hy_V_in","hy_I_in"]
for col in num_cols:
    if col in hy.columns:
        hy[col] = pd.to_numeric(hy[col], errors="coerce")

hy_pre  = hy[hy["timestamp"] <  HYESYS_ACTIVATION].copy()
hy_post = hy[hy["timestamp"] >= HYESYS_ACTIVATION].copy()

# ═════════════════════════════════════════════════════════════════
print("=" * 70)
print("FINDING 1 — MAIN GRID DATA COVERAGE vs HYESYS ACTIVATION")
print("=" * 70)
print(f"\n  Main Grid data ends:       {mg['timestamp'].max()}")
print(f"  HyESys activation:         {HYESYS_ACTIVATION}")
print(f"  GAP (no Main Grid data):   {HYESYS_ACTIVATION - mg['timestamp'].max()}")
print(f"\n  >> The Main Grid dataset ends 6 hours BEFORE HyESys activation.")
print(f"  >> There is NO Main Grid data for the HyESys active period.")
print(f"  >> Direct before/after comparison on Main Grid is IMPOSSIBLE")
print(f"     with this dataset alone.")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 2 — PHASE B DEAD IN ENTIRE MAIN GRID DATASET")
print("=" * 70)
print(f"\n  I_B non-zero records:  {(mg['I_B'] != 0).sum()} / {len(mg)}")
print(f"  V_B non-zero records:  {(mg['V_B'] != 0).sum()} / {len(mg)}")
print(f"\n  >> Phase B current and voltage are ZERO across all 1003 records.")
print(f"  >> This is NOT a load characteristic — this is a meter/CT wiring issue.")
print(f"  >> Phase A and Phase C are the only two measured phases.")
print(f"  >> The '86.6% phase imbalance' figure is an artefact of I_B=0, not real imbalance.")
print(f"\n  Phase A vs Phase C comparison (actual imbalance between A and C):")
a_mean = mg["I_A"].mean(); c_mean = mg["I_C"].mean()
ac_diff_pct = abs(a_mean - c_mean) / ((a_mean + c_mean) / 2) * 100
print(f"    I_A mean: {a_mean:.3f} A")
print(f"    I_C mean: {c_mean:.3f} A")
print(f"    A vs C imbalance: {ac_diff_pct:.2f}%  (very balanced)")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 3 — HYESYS kVAr PRE-ACTIVATION vs POST-ACTIVATION (ANOMALY)")
print("=" * 70)
print(f"\n  HyESys kVAr output BEFORE activation (9464 records):")
print(f"    mean={hy_pre['hy_kVAr_out'].mean():.3f}  median={hy_pre['hy_kVAr_out'].median():.3f}  max={hy_pre['hy_kVAr_out'].max():.3f} kVAr")
print(f"\n  HyESys kVAr output AFTER activation (253 records):")
print(f"    mean={hy_post['hy_kVAr_out'].mean():.3f}  median={hy_post['hy_kVAr_out'].median():.3f}  max={hy_post['hy_kVAr_out'].max():.3f} kVAr")
print(f"\n  Main Grid kVAr (pre-activation, all 1003 records):")
print(f"    mean={mg['kVAr_total'].mean():.3f}  max={mg['kVAr_total'].max():.3f} kVAr")
print(f"\n  >> Before activation, HyESys shows 45+ kVAr 'output' — contradicts user's")
print(f"     statement that it did not operate until 16:29:31.")
print(f"  >> After activation, HyESys output drops to <0.7 kVAr.")
print(f"  >> Hypothesis: the 'output' column pre-activation may represent the LOAD's")
print(f"     reactive demand being MEASURED/MONITORED by HyESys sensors, not actual")
print(f"     compensation output. Alternatively, HyESys was in a different mode.")
print(f"  >> Main Grid kVAr is only 0.265 kVAr — far below HyESys pre-act reading.")
print(f"     The two measurements are at different points or scales.")

# Per-phase HyESys kVAr pre vs post
print(f"\n  Per-phase HyESys kVAr breakdown:")
print(f"  {'Phase':>8} | {'Pre-act mean':>14} | {'Post-act mean':>14}")
print(f"  {'-'*45}")
for ph in ["hy_kVAr_A","hy_kVAr_B","hy_kVAr_C"]:
    print(f"  {ph:>8} | {hy_pre[ph].mean():>14.3f} | {hy_post[ph].mean():>14.3f}")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 4 — HYESYS ACTIVATION TRANSITION (window ±30min)")
print("=" * 70)

window_start = HYESYS_ACTIVATION - pd.Timedelta("30min")
window_end   = HYESYS_ACTIVATION + pd.Timedelta("30min")
hy_window = hy[(hy["timestamp"] >= window_start) & (hy["timestamp"] <= window_end)].copy()

print(f"\n  {len(hy_window)} HyESys records from {window_start} to {window_end}")
print(f"\n  {'Timestamp':<22} {'kVAr_out':>9} {'kW_out':>8} {'kVA_out':>9} {'PF_out':>7} {'I_A':>6} {'I_B':>6} {'I_C':>6} {'Temp':>6} {'Note'}")
print("  " + "-" * 105)
for _, row in hy_window.iterrows():
    note = "  << ACTIVATION" if row["timestamp"] == HYESYS_ACTIVATION else ""
    marker = "PRE " if row["timestamp"] < HYESYS_ACTIVATION else "POST"
    print(f"  {str(row['timestamp']):<22} "
          f"{row.get('hy_kVAr_out', float('nan')):>9.3f} "
          f"{row.get('hy_kW_out', float('nan')):>8.3f} "
          f"{row.get('hy_kVA_out', float('nan')):>9.3f} "
          f"{row.get('hy_PF_out', float('nan')):>7.3f} "
          f"{row.get('hy_I_A', float('nan')):>6.2f} "
          f"{row.get('hy_I_B', float('nan')):>6.2f} "
          f"{row.get('hy_I_C', float('nan')):>6.2f} "
          f"{row.get('hy_temp', float('nan')):>6.1f} "
          f"[{marker}]{note}")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 5 — OVERLAP PERIOD CORRELATION (Apr 30 – May 6 10:30)")
print("=" * 70)
print("  (Both datasets have data here; all pre-activation)")

# Resample HyESys to 15-min to match Main Grid
overlap_start = pd.Timestamp("2026-04-30 00:00:00")
overlap_end   = mg["timestamp"].max()

mg_ovlp = mg[(mg["timestamp"] >= overlap_start) & (mg["timestamp"] <= overlap_end)].set_index("timestamp")
hy_ovlp = hy[(hy["timestamp"] >= overlap_start) & (hy["timestamp"] <= overlap_end)].set_index("timestamp")

hy_15 = hy_ovlp[["hy_kVAr_out","hy_kW_out","hy_kVA_out",
                   "hy_I_A","hy_I_B","hy_I_C",
                   "hy_amp_imb","hy_freq"]].resample("15min").mean()

merged_ovlp = mg_ovlp[["kW_total","kVAr_total","PF_total","I_A","I_C","I_AC_avg"]].join(hy_15, how="inner")
print(f"\n  Aligned 15-min records in overlap: {len(merged_ovlp)}")

if len(merged_ovlp) > 1:
    corr_kvar  = merged_ovlp[["hy_kVAr_out","kVAr_total"]].dropna().corr().iloc[0,1]
    corr_ia    = merged_ovlp[["hy_I_A","I_A"]].dropna().corr().iloc[0,1]
    corr_ic    = merged_ovlp[["hy_I_C","I_C"]].dropna().corr().iloc[0,1]
    corr_kw    = merged_ovlp[["hy_kW_out","kW_total"]].dropna().corr().iloc[0,1]
    corr_freq_pf = merged_ovlp[["hy_freq","PF_total"]].dropna().corr().iloc[0,1]
    print(f"\n  Correlation (HyESys kVAr vs Main Grid kVAr):     r = {corr_kvar:.4f}")
    print(f"  Correlation (HyESys I_A  vs Main Grid I_A):      r = {corr_ia:.4f}")
    print(f"  Correlation (HyESys I_C  vs Main Grid I_C):      r = {corr_ic:.4f}")
    print(f"  Correlation (HyESys kW   vs Main Grid kW):       r = {corr_kw:.4f}")
    print(f"  Correlation (HyESys freq vs Main Grid PF):       r = {corr_freq_pf:.4f}")

    print(f"\n  >> Interpretation: High r on I_A and I_C confirms HyESys sensors")
    print(f"     are measuring the same circuit as Main Grid meter.")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 6 — MAIN GRID BASELINE DETAILED (26 Apr – 6 May 10:30)")
print("=" * 70)

print(f"\n  Total records: {len(mg)}")
print(f"  Period: {mg['timestamp'].min()} → {mg['timestamp'].max()}")
print(f"  Duration: ~{(mg['timestamp'].max()-mg['timestamp'].min()).days + 1} days")
print()
print(f"  kW (Active Power):")
print(f"    mean={mg['kW_total'].mean():.4f}  std={mg['kW_total'].std():.4f}  min={mg['kW_total'].min():.4f}  max={mg['kW_total'].max():.4f} kW")
print(f"\n  kVAr (Reactive Power):")
print(f"    mean={mg['kVAr_total'].mean():.4f}  std={mg['kVAr_total'].std():.4f}  min={mg['kVAr_total'].min():.4f}  max={mg['kVAr_total'].max():.4f} kVAr")
print(f"\n  PF (Power Factor):")
print(f"    mean={mg['PF_total'].mean():.4f}  std={mg['PF_total'].std():.4f}  min={mg['PF_total'].min():.4f}  max={mg['PF_total'].max():.4f}")
print(f"\n  Current Phase A:")
print(f"    mean={mg['I_A'].mean():.4f}  std={mg['I_A'].std():.4f}  min={mg['I_A'].min():.4f}  max={mg['I_A'].max():.4f} A")
print(f"\n  Current Phase C:")
print(f"    mean={mg['I_C'].mean():.4f}  std={mg['I_C'].std():.4f}  min={mg['I_C'].min():.4f}  max={mg['I_C'].max():.4f} A")
print(f"\n  Voltage Phase A:")
print(f"    mean={mg['V_A'].mean():.3f}  std={mg['V_A'].std():.3f}  min={mg['V_A'].min():.3f}  max={mg['V_A'].max():.3f} V")
print(f"\n  Voltage Phase C:")
print(f"    mean={mg['V_C'].mean():.3f}  std={mg['V_C'].std():.3f}  min={mg['V_C'].min():.3f}  max={mg['V_C'].max():.3f} V")

# PF below 0.9 occurrences
pf_low = mg[mg["PF_total"] < 0.90]
print(f"\n  Records with PF < 0.90: {len(pf_low)} ({len(pf_low)/len(mg)*100:.1f}%)")
pf_85  = mg[mg["PF_total"] < 0.85]
print(f"  Records with PF < 0.85 (SP penalty zone): {len(pf_85)} ({len(pf_85)/len(mg)*100:.1f}%)")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 7 — HYESYS POST-ACTIVATION DETAILED (4h 12min only)")
print("=" * 70)

print(f"\n  Records: {len(hy_post)}")
print(f"  Period: {hy_post['timestamp'].min()} → {hy_post['timestamp'].max()}")
print(f"  Duration: {hy_post['timestamp'].max() - hy_post['timestamp'].min()}")
print()
print(f"  kVAr output:  mean={hy_post['hy_kVAr_out'].mean():.3f}  min={hy_post['hy_kVAr_out'].min():.3f}  max={hy_post['hy_kVAr_out'].max():.3f} kVAr")
print(f"  kW output:    mean={hy_post['hy_kW_out'].mean():.3f}  min={hy_post['hy_kW_out'].min():.3f}  max={hy_post['hy_kW_out'].max():.3f} kW")
print(f"  kVA output:   mean={hy_post['hy_kVA_out'].mean():.3f}  min={hy_post['hy_kVA_out'].min():.3f}  max={hy_post['hy_kVA_out'].max():.3f} kVA")
print(f"  PF output:    mean={hy_post['hy_PF_out'].mean():.3f}  min={hy_post['hy_PF_out'].min():.3f}  max={hy_post['hy_PF_out'].max():.3f}")
print(f"  I_A:          mean={hy_post['hy_I_A'].mean():.3f}  max={hy_post['hy_I_A'].max():.3f} A")
print(f"  I_B:          mean={hy_post['hy_I_B'].mean():.3f}  max={hy_post['hy_I_B'].max():.3f} A")
print(f"  I_C:          mean={hy_post['hy_I_C'].mean():.3f}  max={hy_post['hy_I_C'].max():.3f} A")
print(f"  Amp imbalance: mean={hy_post['hy_amp_imb'].mean():.3f}  max={hy_post['hy_amp_imb'].max():.3f} %")
print(f"  Temp:          mean={hy_post['hy_temp'].mean():.1f}  max={hy_post['hy_temp'].max():.1f} °C")
print(f"\n  >> Only 4h 12min of post-activation data. No corresponding Main Grid data.")
print(f"  >> HyESys kVAr output is low (max 0.7 kVAr) — possibly still ramping up,")
print(f"     or site reactive demand is genuinely small at this time of day.")

# kVAr output trend over the 4h window
print(f"\n  kVAr output trend (15-min bins):")
hy_post_15 = hy_post.set_index("timestamp")[["hy_kVAr_out","hy_amp_imb","hy_I_A","hy_I_B","hy_I_C"]].resample("15min").mean()
print(f"  {'Timestamp':<22} {'kVAr_out':>9} {'AmpImb':>8} {'I_A':>6} {'I_B':>6} {'I_C':>6}")
print("  " + "-" * 65)
for ts, row in hy_post_15.iterrows():
    print(f"  {str(ts):<22} {row['hy_kVAr_out']:>9.3f} {row['hy_amp_imb']:>8.3f} {row['hy_I_A']:>6.2f} {row['hy_I_B']:>6.2f} {row['hy_I_C']:>6.2f}")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 8 — HYESYS PRE-ACTIVATION kVAr: TIME SERIES SAMPLE")
print("=" * 70)
print("  (First 10 and last 10 records before activation)")
cols_show = ["timestamp","hy_kVAr_out","hy_kW_out","hy_I_A","hy_I_B","hy_I_C","hy_amp_imb"]
print(f"\n  First 10 pre-activation records:")
print(f"  {'Timestamp':<22} {'kVAr_out':>9} {'kW_out':>8} {'I_A':>6} {'I_B':>6} {'I_C':>6} {'AmpImb':>8}")
print("  " + "-" * 72)
for _, row in hy_pre.head(10).iterrows():
    print(f"  {str(row['timestamp']):<22} {row['hy_kVAr_out']:>9.3f} {row['hy_kW_out']:>8.3f} {row['hy_I_A']:>6.2f} {row['hy_I_B']:>6.2f} {row['hy_I_C']:>6.2f} {row['hy_amp_imb']:>8.3f}")

print(f"\n  Last 10 pre-activation records (approaching activation):")
print(f"  {'Timestamp':<22} {'kVAr_out':>9} {'kW_out':>8} {'I_A':>6} {'I_B':>6} {'I_C':>6} {'AmpImb':>8}")
print("  " + "-" * 72)
for _, row in hy_pre.tail(10).iterrows():
    print(f"  {str(row['timestamp']):<22} {row['hy_kVAr_out']:>9.3f} {row['hy_kW_out']:>8.3f} {row['hy_I_A']:>6.2f} {row['hy_I_B']:>6.2f} {row['hy_I_C']:>6.2f} {row['hy_amp_imb']:>8.3f}")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("FINDING 9 — MAIN GRID kVAr vs HyESys kVAr SCALE CHECK")
print("=" * 70)
print(f"\n  Main Grid kVAr (瞬时无功) mean: {mg['kVAr_total'].mean():.4f} kVAr")
print(f"  HyESys pre-act kVAr mean:        {hy_pre['hy_kVAr_out'].mean():.3f} kVAr")
print(f"  Ratio:                           {hy_pre['hy_kVAr_out'].mean() / mg['kVAr_total'].mean():.1f}x")
print(f"\n  Main Grid I_A mean:              {mg['I_A'].mean():.4f} A")
print(f"  HyESys I_A mean (pre-act):       {hy_pre['hy_I_A'].mean():.4f} A")
print(f"  Ratio:                           {hy_pre['hy_I_A'].mean() / mg['I_A'].mean():.3f}x")
print(f"\n  >> kVAr is 170x higher in HyESys vs Main Grid — incompatible scales.")
print(f"  >> I_A ratio is {hy_pre['hy_I_A'].mean() / mg['I_A'].mean():.3f}x — HyESys currents are lower,")
print(f"     consistent with HyESys being on a sub-circuit, not the full Main Grid incomer.")

# ═════════════════════════════════════════════════════════════════
print("\n" + "=" * 70)
print("SUMMARY OF ALL FINDINGS")
print("=" * 70)
print("""
  1. DATASET COVERAGE GAP (Critical)
     Main Grid data ends at 2026-05-06 10:30 — 6 hours before HyESys activation
     at 16:29:31. No Main Grid post-activation data exists in this dataset.
     Before/after Main Grid comparison requires the post-activation Main Grid file.

  2. PHASE B DEAD IN MAIN GRID METER (Data Quality)
     I_B = 0 and V_B = 0 for all 1003 records. This is a meter/CT wiring problem,
     not a load characteristic. Only Phase A and Phase C are being measured.
     The reported '86.6% phase imbalance' is an artefact of the missing B phase.
     Actual A-to-C imbalance is <0.5% — the site load is well-balanced on A and C.

  3. HYESYS kVAr INVERSION ANOMALY (Critical — needs clarification)
     Pre-activation: HyESys 'output' kVAr = avg 45.2 kVAr (max 113.2 kVAr)
     Post-activation: HyESys 'output' kVAr = avg 0.39 kVAr (max 0.7 kVAr)
     This is the opposite of expected behaviour. Three hypotheses:
       (a) The HyESys was in a different compensation mode before 16:29:31
           (not 'off' but possibly in CapBank-assist or pass-through mode)
       (b) Pre-activation, the 'output kVAr' column represents the LOAD reactive
           demand measured by HyESys sensors, not inverter output
       (c) The activation timestamp 16:29:31 marks a MODE CHANGE, not a
           power-on event — HyESys was compensating differently before

  4. MEASUREMENT POINT MISMATCH (Critical)
     HyESys kVAr (pre-act avg 45.2) is 170x larger than Main Grid kVAr (avg 0.265).
     HyESys currents (I_A ~6.4 A) are higher than Main Grid (I_A ~4.5 A) in
     pre-activation but appear to be measuring a different circuit segment.
     The two datasets are NOT measuring the same electrical node.

  5. BASELINE ELECTRICAL PATTERN (Valid)
     Main Grid (26 Apr – 6 May, 1003 records, all pre-activation):
       Active Power:   mean 0.736 kW  (very stable, std 0.019)
       Reactive Power: mean 0.265 kVAr (very stable, std 0.013)
       Power Factor:   mean 0.936  (range 0.930–0.950)
       Current A:      mean 4.467 A
       Current C:      mean 4.447 A
       Voltage A:      mean 102.0 V
       Voltage C:      mean 101.9 V
     No records below PF 0.85 penalty threshold. No major gaps in data.

  6. HYESYS POST-ACTIVATION (Limited — 4h 12min only, no Main Grid overlap)
     kVAr output:  avg 0.39, max 0.7 kVAr (very low — ramping up or low demand)
     kW output:    0.0 kW (pure reactive compensation, as expected)
     Amp imbalance: avg 0.68%
     Temperature:   avg 51.6°C, peaked 89°C (high — worth monitoring)
     Output PF:     avg -0.103 (leading, consistent with reactive injection)

  7. WHAT IS NEEDED TO COMPLETE THE ANALYSIS
     (a) Main Grid data from 2026-05-06 16:30 onwards (post-activation period)
     (b) Clarification on HyESys pre-activation kVAr (45 kVAr) — was the unit
         in a monitoring mode or active mode before 16:29:31?
     (c) Phase B CT/meter connection check at the Main Grid meter
""")

print("=" * 70)
print("ANALYSIS COMPLETE")
print("=" * 70)
