# HyESys Agent Master Rules Reference
**Version:** 1.0  |  **Created:** 2026-05-24  |  **Maintained by:** AST — Advancer Smart Technology

---

## How to use this document

This is the authoritative reference for all rules and logic in Agent 1 and Agent 2.
When deploying to a new site:

1. Copy `agent1_master.py` → `sites/<site_id>/agent1.py`
2. Copy `agent2_master.py` → `sites/<site_id>/agent2.py`
3. Edit the `SITE_CONFIG` block in each file
4. Override only the methods labelled **OVERRIDE POINT** if needed
5. Changes in one site's files have zero effect on any other site

---

## Agent 1 — Data Quality Validator

### Purpose
Reads raw meter data row by row. Tags each record **CLEAN**, **SUSPECT**, or **REJECTED**.
Only CLEAN and SUSPECT records are written to the database. REJECTED records are dropped.

### Rule execution order

| Rule | Name | Condition | Tag |
|------|------|-----------|-----|
| R1 | Timestamp parse | Timestamp missing or cannot be parsed | REJECTED |
| R2 | Duplicate timestamp | Same site_id + timestamp already seen in this batch | REJECTED |
| R3 | Numeric parse | kW, kVAr, PF, or voltage cannot be cast to a number | REJECTED |
| R4 | Voltage impossible | Voltage ≤ 0 | REJECTED |
| R5 | kW range | kW outside [−kW_MAX, +kW_MAX] | REJECTED |
| R6 | kVAr range | kVAr outside [−kVAr_MAX, +kVAr_MAX] | REJECTED |
| R7 | PF firmware bug | PF outside [−1.0, +1.0] | REJECTED |
| R8 | All-zero row | kW = 0 AND kVAr = 0 AND PF = 0 simultaneously | SUSPECT |
| R9 | PF saturation | abs(PF) == exactly 1.0 | SUSPECT |
| R10 | Voltage out of range | Voltage outside (nominal ± tolerance %) | SUSPECT |
| R11 | Data gap | Time since last record > (expected interval × multiplier) | SUSPECT |
| W1 | PF penalty warning | 0 < abs(PF) < 0.85 (SP penalty threshold) | CLEAN + warning logged |

Rules R1–R7 are **hard stops** — the record is dropped immediately.
Rules R8–R11 **retain** the record as SUSPECT.
W1 passes the record CLEAN but logs a warning.

### Pre-processing step (before any rule)
Multiplying factors (CT ratio, PT ratio, 综合倍率) are applied to raw meter values **before** validation. Site files set `mf_kW`, `mf_kVAr`, `mf_voltage`, `mf_current` in their SITE_CONFIG.

### Site-configurable parameters (Agent 1)

| Parameter | Master default | What it controls |
|-----------|---------------|-----------------|
| `voltage_nominal_V` | 230.0 V | Centre of acceptable voltage range (Rule R10) |
| `voltage_tolerance_pct` | 15% | Width of acceptable voltage band (Rule R10) |
| `kW_max` | 10,000 | Hard reject if kW outside ±this value (Rule R5) |
| `kVAr_max` | 5,000 | Hard reject if kVAr outside ±this value (Rule R6) |
| `mf_kW` / `mf_kVAr` / `mf_voltage` / `mf_current` | 1.0 (no scaling) | Multiplying factors for MV revenue meters |
| `unmeasured_phases` | [] | Phases with no CT/PT — skip checks for those columns |
| `expected_interval_min` | 15 | Expected gap between records in minutes (Rule R11) |
| `gap_multiplier` | 2.0 | Gap threshold = expected × this (Rule R11) |
| `pf_penalty_threshold` | 0.85 | SP penalty floor for W1 warning |

### What is NOT in Agent 1 master (add in site file)
- Activation-state tagging (pre-activation vs post-activation labels)
- Per-column name mapping if site meter uses different headers
- Phase exclusion logic for partially-instrumented meters
- Meter-specific data cleaning (e.g. strip tab characters from timestamps)

---

## Agent 2 — Analysis & Recommendation Engine

### Purpose
Reads validated 15-minute state snapshots from the database.
Detects events, computes kVAr injection decisions, and logs STATE → ACTION → REWARD (SAR) triplets.

---

### Layer 1: State

The `State` struct captures one 15-minute snapshot and exposes derived electrical quantities as properties.

| Property | Equation | Unit |
|----------|----------|------|
| `kVA` | √(P² + Q²) | kVA |
| `phi_rad` | arccos(abs(PF)) | rad |
| `phi_deg` | degrees(phi_rad) | ° |
| `tan_phi` | tan(phi_rad) | — |
| `q_target` | P × tan(arccos(PF_target)) where PF_target = 0.98 | kVAr |
| `q_correction_needed` | kVAr − q_target | kVAr |
| `current_A` | S × 1000 / (3 × V_LN) | A |
| `loss_fraction_vs_unity_pf` | 1 − PF² | — |
| `recoverable_loss_fraction` | 1 − (P/PF_target / S)² | — |

---

### Layer 2: Events (12 event types)

Events are detected on every 15-min snapshot. All four types run every cycle.

#### Threshold events (fixed limits)

| Code | Subtype | Condition | Severity |
|------|---------|-----------|----------|
| E1 | PF_PENALTY_RISK | abs(PF) < 0.85 | CRITICAL |
| E2 | PF_LOW | abs(PF) < PF_target (0.98) | WARNING |
| E3 | PF_LEADING | kVAr < −5 kVAr (configurable) | WARNING |

#### Statistical events (EMA z-score)

EMA α = 2/(N+1) where N = 8 intervals (2 hours).

| Code | Subtype | Condition | Severity |
|------|---------|-----------|----------|
| E4 | PF_ANOMALY_LOW | PF z-score < −3σ | CRITICAL |
| E5 | PF_TREND_DOWN | PF z-score < −2σ | WARNING |
| E6 | DEMAND_SPIKE | kW z-score > +3σ | CRITICAL |
| E7 | DEMAND_ELEVATED | kW z-score > +2σ | WARNING |
| E8 | REACTIVE_SURGE | abs(kVAr) z-score > +3σ | WARNING |

#### Composite events (multiple conditions simultaneously)

| Code | Subtype | Condition | Severity |
|------|---------|-----------|----------|
| E9 | LOW_PF_HIGH_DEMAND | PF < target AND kW ≥ 90% of recent peak | CRITICAL |
| E10 | PF_VOLTAGE_SAG | PF < target AND V < 220 V | WARNING |

#### Scheduled events (time-based)

| Code | Subtype | Condition | Severity |
|------|---------|-----------|----------|
| E11 | PEAK_PERIOD | Hour 08:00–22:00 (Singapore default) | INFO |
| E12 | OFFPEAK_PERIOD | Hour 22:00–08:00 | INFO |

---

### Layer 3: Decision Engine (PI Controller)

Control law runs on every 15-min snapshot:

```
Error signal:     e(t) = PF_target − abs(PF_current)
                  e > 0 → lagging PF, inject capacitive kVAr
                  e < 0 → leading PF, reduce injection
                  e = 0 → at target, hold

Dead-band:        |e(t)| < DEADBAND (0.005) → HOLD

Proportional:     ΔQ_P = K_P × P × (tan φ_current − tan φ_target)
                  K_P = 1.00 (full correction; lower for smoother ramp)

Integral:         I(t) = clamp(I(t−1) + e(t) × Δt, −I_MAX, +I_MAX)
                  ΔQ_I = K_I × I(t)
                  K_I = 0.50    I_MAX = 20 kVAr·hr    Δt = 0.25 hr

Combined:         ΔQ_cmd = ΔQ_P + ΔQ_I
                  Clamp ΔQ_cmd to [−model_kVA, +model_kVA]
```

#### Decision priority

| Priority | Condition | Action |
|----------|-----------|--------|
| 1 | has_solar AND demand == CRITICAL AND recommend_store | ACTION_REDUCE (reserve for storage) |
| 2 | abs(e(t)) < DEADBAND | ACTION_HOLD |
| 3 | ΔQ_cmd > +0.5 kVAr | ACTION_INJECT with magnitude |
| 4 | ΔQ_cmd < −0.5 kVAr | ACTION_REDUCE with magnitude |
| 5 | otherwise | ACTION_HOLD |

#### Demand risk thresholds

| Level | Condition | recommend_store |
|-------|-----------|----------------|
| CRITICAL | demand_pct ≥ 95% | True |
| HIGH | demand_pct ≥ 85% | True |
| MEDIUM | demand_pct ≥ 70% | True if has_solar |
| LOW | demand_pct < 70% | False |

`demand_pct = P_current / P_peak_historical × 100`

---

### Layer 4: Reward

Closes the SAR (State → Action → Reward) loop.

| Component | Equation | Weight |
|-----------|----------|--------|
| PF improvement | r_PF = PF_after − PF_before | 0.60 |
| Loss fraction | r_loss = 1 − (S_after/S_before)² | 0.40 |
| Combined | r_total = 0.60 × r_PF + 0.40 × r_loss | — |
| THD back-calc | THD = √((PF_before/PF_target)² / (1−f) − 1) | — |

**Outcome classification:**
- r_PF ≥ +0.01 → **POSITIVE**
- r_PF ≤ −0.01 → **NEGATIVE**
- otherwise → **NEUTRAL**

### Site-configurable parameters (Agent 2)

| Parameter | Master default | What it controls |
|-----------|---------------|-----------------|
| `has_solar` | True | Enables D1 priority solar/demand shaving |
| `hyesys_model` | H50 | Hardware kVA capacity used in clamping |
| `K_P` | 1.00 | Proportional gain (1.0 = full correction per step) |
| `K_I` | 0.50 | Integral gain |
| `I_MAX` | 20.0 kVAr·hr | Anti-windup clamp |
| `DEADBAND` | 0.005 | PF dead-band threshold |
| `demand_critical_pct` | 95 | CRITICAL demand threshold |
| `demand_high_pct` | 85 | HIGH demand threshold |
| `demand_medium_pct` | 70 | MEDIUM demand threshold |
| `demand_tariff_sgd_per_kw` | 10.0 | Demand charge rate for risk estimation |
| `z_warn` | 2.0 | Statistical event warning z-score |
| `z_crit` | 3.0 | Statistical event critical z-score |
| `ema_n` | 8 | EMA window length (number of intervals) |
| `leading_pf_q_threshold` | −5.0 kVAr | Threshold for E3 PF_LEADING event |
| `voltage_sag_V` | 220 V | Threshold for E10 PF_VOLTAGE_SAG event |
| `high_demand_pct` | 0.90 | Fraction of recent peak for E9 composite event |
| `peak_hours_start` | 8 | Start hour for PEAK_PERIOD scheduled event |
| `peak_hours_end` | 22 | End hour for PEAK_PERIOD scheduled event |
| `history_window` | 16 | State history depth per site |

### What is NOT in Agent 2 master (add in site file)
- Live MQTT command dispatch (`process()` OVERRIDE POINT)
- Non-Singapore tariff peak/off-peak hours
- Harmonic compensation mode (future HyESys V2)
- Multi-unit coordination (when multiple HyESys units share a bus)
- Site-specific alarm escalation (email/SMS/Telegram on CRITICAL events)

---

## Per-site file structure

```
sites/
└── <site_id>/
    ├── agent1.py     ← copy of templates/agent1_master.py + site customisations
    └── agent2.py     ← copy of templates/agent2_master.py + site customisations
```

Each site file starts with a **CHANGELOG** block at the top recording every change made from the master, e.g.:

```python
# ── SITE CHANGELOG (changes from master) ──────────────────────────
# 2026-05-24  initial copy from master v1.0
# 2026-05-24  voltage_nominal_V = 229 V  (LV bus at Cabinet A / Feeder 2)
# 2026-05-24  expected_interval_min = 1  (HyESys logs per minute)
# 2026-05-24  mf_kW / mf_kVAr = 1.0     (HyESys values are already real)
```

---

## Known deployed sites

| Site ID | Solar | HyESys model | Agent 1 notes | Agent 2 notes |
|---------|-------|-------------|--------------|--------------|
| baoyuan | Yes | H125 | LV 229 V, 1-min intervals, no MF needed for HyESys file | K_P may need tuning; Chinese tariff |
| INLET-METER-MAR26 | Yes | H125 | — | avg reactive 55 kVAr |
| MSB-SPPG2-MAR26 | Yes | H50 | solar export (negative kW valid) | — |
| FederalOatMills-MSB1 | No | H50 | — | — |
| FederalOatMills-MSB2 | No | H50 | — | — |
| ShanPoornam | No | H50 | — | — |
| SRN-INCOMING | Yes | H50 | — | — |
| ST-INCOMING | No | H50 | — | — |
