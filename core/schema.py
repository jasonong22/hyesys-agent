"""
Shared data schema — column names, types, and system-wide constants.
Single source of truth for all modules.
"""

# ── Power factor targets ───────────────────────────────────────────
PF_TARGET            = 0.98   # target PF (not 1.0 — law of convergence)
PF_PENALTY_THRESHOLD = 0.85   # SP penalty triggered below this
THD_ASSUMPTION       = 0.15   # 15% THD starting assumption (mixed building)

# ── Voltage limits (Singapore LV, 230V ±15%) ──────────────────────
VOLTAGE_NOMINAL  = 230.0
VOLTAGE_MIN      = 195.5   # 230 × 0.85
VOLTAGE_MAX      = 264.5   # 230 × 1.15

# ── DB / CSV column names ──────────────────────────────────────────
COL_SITE_ID   = "site_id"
COL_TIMESTAMP = "timestamp"
COL_KW        = "kW"
COL_KVAR      = "kVAr"
COL_PF        = "PF"
COL_VOLTAGE   = "voltage_V"
COL_QUALITY   = "quality_tag"
COL_REASON    = "reject_reason"
COL_INGESTED  = "ingested_at"

# ── Quality tags ───────────────────────────────────────────────────
CLEAN    = "CLEAN"
SUSPECT  = "SUSPECT"
REJECTED = "REJECTED"

# ── Agent 2 actions ───────────────────────────────────────────────
ACTION_INJECT = "INJECT_KVAR"
ACTION_HOLD   = "HOLD"
ACTION_REDUCE = "REDUCE"

# ── SAR reward thresholds ─────────────────────────────────────────
REWARD_POSITIVE_PF_DELTA  =  0.01   # PF improvement >= this → positive outcome
REWARD_NEGATIVE_PF_DELTA  = -0.01   # PF worsened by this → negative outcome

# ── HyESys hardware models ────────────────────────────────────────
HYESYS_MODELS = {
    "H30":  {"kVA": 30,  "max_current_A": 43.5,  "vdc_range": "210–850 V",  "dc_voltage_min_V": 231.0,  "dc_voltage_max_V": 269.5, "kWh": 69.3,  "packs": 7,  "weight_kg": 1400, "footprint_m2": 2.1, "price_sgd": 100_000},
    "H50":  {"kVA": 50,  "max_current_A": 72.5,  "vdc_range": "350–850 V",  "dc_voltage_min_V": 363.0,  "dc_voltage_max_V": 423.5, "kWh": 108.9, "packs": 11, "weight_kg": 2200, "footprint_m2": 3.2, "price_sgd": 120_000},
    "H60":  {"kVA": 60,  "max_current_A": 87.0,  "vdc_range": "420–850 V",  "dc_voltage_min_V": 462.0,  "dc_voltage_max_V": 539.0, "kWh": 138.6, "packs": 14, "weight_kg": 2800, "footprint_m2": 4.2, "price_sgd": None},
    "H100": {"kVA": 100, "max_current_A": 145.0, "vdc_range": "680–900 V",  "dc_voltage_min_V": 726.0,  "dc_voltage_max_V": 847.0, "kWh": 217.8, "packs": 22, "weight_kg": 4400, "footprint_m2": 6.3, "price_sgd": None},
    "H125": {"kVA": 125, "max_current_A": 181.0, "vdc_range": "680–900 V",  "dc_voltage_min_V": 726.0,  "dc_voltage_max_V": 847.0, "kWh": 217.8, "packs": 22, "weight_kg": 4400, "footprint_m2": 6.3, "price_sgd": 100_000},
}

# ── Known deployment sites ────────────────────────────────────────
SITE_CONFIG = {
    "INLET-METER-MAR26":    {"solar": True,  "recommended_model": "H125", "notes": "avg reactive 55 kVAr"},
    "MSB-SPPG2-MAR26":      {"solar": True,  "recommended_model": "H50",  "notes": "solar export — negative kW valid"},
    "FederalOatMills-MSB1": {"solar": False, "recommended_model": "H50",  "notes": "no solar — capped at H50"},
    "FederalOatMills-MSB2": {"solar": False, "recommended_model": "H50",  "notes": "no solar — capped at H50"},
    "ShanPoornam":          {"solar": False, "recommended_model": "H50",  "notes": "no solar — capped at H50"},
    "SRN-INCOMING":         {"solar": True,  "recommended_model": "H50",  "notes": ""},
    "ST-INCOMING":          {"solar": False, "recommended_model": "H50",  "notes": "no solar — capped at H50"},
}

# ── Saving priority order ─────────────────────────────────────────
# 1 = highest priority
SAVING_PRIORITY = {
    "solar_storage":      1,
    "reactive_correction": 2,
    "load_balancing":     3,
}
