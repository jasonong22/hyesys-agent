"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 1 — MASTER TEMPLATE                                          ║
║  HyESys Data Quality Validator                                      ║
║  Version: 1.0  |  Created: 2026-05-24                               ║
╚══════════════════════════════════════════════════════════════════════╝

PURPOSE
───────
This is the frozen master reference for Agent 1.
Do NOT modify this file.

To deploy Agent 1 for a new site:
  1. Copy this file to  sites/<site_id>/agent1.py
  2. Edit only the SITE CONFIGURATION block below
  3. Optionally override validate() methods marked "OVERRIDE POINT"
  4. Each site's agent1.py is completely independent

RULE INVENTORY (11 rules, in execution order)
──────────────────────────────────────────────
REJECTED rules — hard stops, record is dropped:
  R1  Unparseable timestamp
  R2  Duplicate timestamp (same site + timestamp seen in this batch)
  R3  Non-numeric field (kW, kVAr, PF, or voltage cannot be cast to float)
  R4  Voltage physically impossible (≤ 0)
  R5  kW out of range  [−kW_MAX, +kW_MAX]
  R6  kVAr out of range  [−kVAr_MAX, +kVAr_MAX]
  R7  PF firmware bug — outside [−1.0, +1.0]

SUSPECT rules — record retained with flag:
  R8  All-zero row  (kW=0 AND kVAr=0 AND PF=0)  → meter dropout
  R9  PF firmware saturation at exactly ±1.0
  R10 Voltage outside nominal range  (nominal ± tolerance %)
  R11 Data gap > expected interval  (configurable per site)

CLEAN pass-through with warning:
  W1  PF below SP penalty threshold  (0 < |PF| < pf_penalty_threshold)
      → CLEAN tag, warning logged

WHAT IS NOT IN THE MASTER (requires site-level customisation):
  • Multiplying factors (CT ratio, PT ratio, combined factor 综合倍率)
  • Per-phase exclusion (e.g. unmeasured B phase)
  • Activation state tagging (pre-activation vs post-activation)
  • Site-specific kW/kVAr operational range
  • MV vs LV voltage nominal
"""

import logging
from dataclasses import dataclass
from core.parser import normalise_timestamp
from core.schema import CLEAN, SUSPECT, REJECTED

log = logging.getLogger("hyesys.agent1")


# ══════════════════════════════════════════════════════════════════════
# SITE CONFIGURATION  — edit this block when copying to a new site
# ══════════════════════════════════════════════════════════════════════

SITE_CONFIG = {
    # ── Identity ──────────────────────────────────────────────────
    "site_id":   "TEMPLATE",           # replace with actual site ID string

    # ── Voltage ───────────────────────────────────────────────────
    # Singapore LV default: 230 V ± 15% → range [195.5 V, 264.5 V]
    # For MV sites: set nominal to actual MV level, e.g. 10_200 V
    "voltage_nominal_V":    230.0,
    "voltage_tolerance_pct": 15.0,     # ±%

    # ── Power range limits (after any multiplying factors applied) ─
    # Hard reject if outside these. Set conservatively wide; tighten per site.
    "kW_max":   10_000,    # kW range: [−kW_max, +kW_max]
    "kVAr_max":  5_000,    # kVAr range: [−kVAr_max, +kVAr_max]

    # ── Multiplying factors ───────────────────────────────────────
    # Applies to raw meter readings before any validation.
    # Set all to 1.0 if meter outputs real engineering values directly.
    # For MV revenue meters: set from the 综合倍率 column in the meter query file.
    "mf_kW":      1.0,   # e.g. 4000 for Baoyuan MV meter
    "mf_kVAr":    1.0,
    "mf_kWh":     1.0,
    "mf_current": 1.0,   # CT ratio
    "mf_voltage": 1.0,   # PT ratio

    # ── Unmeasured phases ─────────────────────────────────────────
    # List phase labels that have no CT/PT — skip range checks for those columns.
    # e.g. ["B"] for Baoyuan MV meter (no CT on B phase)
    "unmeasured_phases": [],

    # ── Timing ────────────────────────────────────────────────────
    # Expected interval between consecutive records (minutes).
    # Used by Rule R11 (data gap detection).
    "expected_interval_min": 15,       # 15 for Main Grid; 1 for HyESys 1-min logs
    "gap_multiplier":         2.0,     # flag SUSPECT if gap > expected × multiplier

    # ── PF penalty threshold ──────────────────────────────────────
    # Singapore Power (SP) penalty is triggered below this.
    "pf_penalty_threshold": 0.85,
}

# Derived voltage limits (computed once from config)
_V_MIN = SITE_CONFIG["voltage_nominal_V"] * (1 - SITE_CONFIG["voltage_tolerance_pct"] / 100)
_V_MAX = SITE_CONFIG["voltage_nominal_V"] * (1 + SITE_CONFIG["voltage_tolerance_pct"] / 100)


# ══════════════════════════════════════════════════════════════════════
# DATA STRUCTURES
# ══════════════════════════════════════════════════════════════════════

@dataclass
class ValidationResult:
    tag:    str           # CLEAN / SUSPECT / REJECTED
    reason: str | None    # None for a clean pass with no warning


# ══════════════════════════════════════════════════════════════════════
# VALIDATOR
# ══════════════════════════════════════════════════════════════════════

class Validator:
    """
    Stateful validator — tracks seen timestamps and last timestamp
    per site_id to detect duplicates and data gaps across a batch.

    Usage:
        v = Validator()
        for raw_row in parse_csv(filepath):
            record, result = v.validate(raw_row)
            if result.tag != REJECTED:
                store.write(record)
    """

    def __init__(self, config: dict = SITE_CONFIG):
        self.cfg    = config
        self._seen:    dict[str, set[str]] = {}
        self._last_ts: dict[str, str]      = {}

        # Derived voltage limits from config
        tol  = self.cfg["voltage_tolerance_pct"] / 100
        nom  = self.cfg["voltage_nominal_V"]
        self._v_min = nom * (1 - tol)
        self._v_max = nom * (1 + tol)

    def validate(self, raw: dict) -> tuple[dict, ValidationResult]:
        """
        Validates a single raw row dict from the site meter.

        Applies multiplying factors first, then all validation rules
        in order R1–R11 followed by the W1 pass-through check.

        Returns (normalised_record, ValidationResult).

        ── OVERRIDE POINT ──────────────────────────────────────────
        Sites may override this method to insert pre-processing steps
        (e.g. phase exclusion, activation-state tagging) before or
        after the standard rules. Call super().validate(raw) to run
        the standard pipeline, then post-process the result.
        """
        site_id = raw.get("site_id", self.cfg.get("site_id", "UNKNOWN"))

        # ── Pre-processing: apply multiplying factors ─────────────
        # OVERRIDE POINT: sites with MV meters override _apply_factors()
        raw = self._apply_factors(raw)

        # ── R1: Timestamp parse ───────────────────────────────────
        ts_raw  = raw.get("timestamp", "")
        ts_norm = normalise_timestamp(ts_raw) if ts_raw else None
        if not ts_norm:
            return self._reject(raw, site_id, None, "R1: unparseable timestamp")

        # ── R2: Duplicate timestamp ───────────────────────────────
        seen = self._seen.setdefault(site_id, set())
        if ts_norm in seen:
            return self._reject(raw, site_id, ts_norm,
                                f"R2: duplicate timestamp: {ts_norm}")
        seen.add(ts_norm)

        # ── R3: Numeric parse ─────────────────────────────────────
        try:
            kw      = float(raw.get("kW",       0) or 0)
            kvar    = float(raw.get("kVAr",     0) or 0)
            pf      = float(raw.get("PF",       0) or 0)
            voltage = float(raw.get("voltage_V",0) or 0)
        except (ValueError, TypeError):
            return self._reject(raw, site_id, ts_norm,
                                "R3: non-numeric field value")

        # ── R4: Voltage physically impossible ─────────────────────
        if voltage <= 0:
            return self._reject(raw, site_id, ts_norm,
                                f"R4: voltage <= 0: {voltage}")

        # ── R5: kW range ──────────────────────────────────────────
        kw_max = self.cfg["kW_max"]
        if not (-kw_max <= kw <= kw_max):
            return self._reject(raw, site_id, ts_norm,
                                f"R5: kW out of range [{-kw_max}, {kw_max}]: {kw}")

        # ── R6: kVAr range ────────────────────────────────────────
        kvar_max = self.cfg["kVAr_max"]
        if not (-kvar_max <= kvar <= kvar_max):
            return self._reject(raw, site_id, ts_norm,
                                f"R6: kVAr out of range [{-kvar_max}, {kvar_max}]: {kvar}")

        # ── R7: PF firmware bug ───────────────────────────────────
        if not (-1.0 <= pf <= 1.0):
            return self._reject(raw, site_id, ts_norm,
                                f"R7: PF firmware bug — outside [-1,1]: {pf}")

        # ── R8: All-zero row (meter dropout) ─────────────────────
        if kw == 0.0 and kvar == 0.0 and pf == 0.0:
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 "R8: all-zero row — possible meter dropout")

        # ── R9: PF firmware saturation ────────────────────────────
        if abs(pf) == 1.0:
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 f"R9: PF firmware saturation at {pf}")

        # ── R10: Voltage outside nominal range ────────────────────
        if not (self._v_min <= voltage <= self._v_max):
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 f"R10: voltage outside {self._v_min:.1f}–{self._v_max:.1f} V: {voltage:.1f} V")

        # ── R11: Data gap ─────────────────────────────────────────
        # OVERRIDE POINT: sites needing strict gap detection can lower gap_multiplier
        gap_result = self._check_gap(site_id, ts_norm)
        if gap_result:
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage, gap_result)

        # ── W1: PF below SP penalty threshold (warn, still CLEAN) ─
        reason = None
        threshold = self.cfg["pf_penalty_threshold"]
        if 0 < abs(pf) < threshold:
            reason = f"W1: PF below SP penalty threshold ({threshold}): {pf:.3f}"
            log.warning("[%s] %s — %s", site_id, ts_norm, reason)

        record = self._build(raw, site_id, ts_norm, kw, kvar, pf, voltage, CLEAN, reason)
        self._last_ts[site_id] = ts_norm
        return record, ValidationResult(CLEAN, reason)

    # ── OVERRIDE POINT: multiplying factors ───────────────────────
    def _apply_factors(self, raw: dict) -> dict:
        """
        Applies meter multiplying factors to raw values before validation.

        Default implementation multiplies kW, kVAr, voltage by the
        factors in SITE_CONFIG. Override in site-specific agents to
        handle column names or additional transformations.
        """
        mf_kw   = self.cfg["mf_kW"]
        mf_kvar = self.cfg["mf_kVAr"]
        mf_v    = self.cfg["mf_voltage"]

        if mf_kw != 1.0 and "kW" in raw:
            raw = dict(raw)
            raw["kW"]        = float(raw.get("kW",        0) or 0) * mf_kw
            raw["kVAr"]      = float(raw.get("kVAr",      0) or 0) * mf_kvar
            raw["voltage_V"] = float(raw.get("voltage_V", 0) or 0) * mf_v
        return raw

    # ── OVERRIDE POINT: gap detection ─────────────────────────────
    def _check_gap(self, site_id: str, ts_norm: str) -> str | None:
        """
        Compares current timestamp against the last seen timestamp.
        Returns a reason string if a gap is detected, else None.

        Override to disable gap detection or change gap logic.
        """
        from datetime import datetime
        last = self._last_ts.get(site_id)
        if last:
            try:
                t_prev = datetime.fromisoformat(last)
                t_curr = datetime.fromisoformat(ts_norm)
                gap_min   = (t_curr - t_prev).total_seconds() / 60
                threshold = (self.cfg["expected_interval_min"]
                             * self.cfg["gap_multiplier"])
                if gap_min > threshold:
                    return (f"R11: data gap {gap_min:.0f} min "
                            f"(expected ≤{threshold:.0f} min)")
            except ValueError:
                pass
        return None

    # ── Internal helpers ───────────────────────────────────────────

    def _build(self, raw, site_id, ts, kw, kvar, pf, voltage, tag, reason):
        return {
            "site_id":       site_id,
            "timestamp":     ts,
            "kW":            kw,
            "kVAr":          kvar,
            "PF":            pf,
            "voltage_V":     voltage,
            "quality_tag":   tag,
            "reject_reason": reason,
            "_solar":        raw.get("_solar", True),
        }

    def _reject(self, raw, site_id, ts, reason):
        log.warning("REJECTED [%s] %s — %s", site_id, ts or "?", reason)
        rec = self._build(raw, site_id, ts or "", 0, 0, 0, 0, REJECTED, reason)
        return rec, ValidationResult(REJECTED, reason)

    def _suspect(self, raw, site_id, ts, kw, kvar, pf, voltage, reason):
        log.warning("SUSPECT  [%s] %s — %s", site_id, ts, reason)
        rec = self._build(raw, site_id, ts, kw, kvar, pf, voltage, SUSPECT, reason)
        self._last_ts[site_id] = ts
        return rec, ValidationResult(SUSPECT, reason)
