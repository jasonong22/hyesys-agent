"""
Agent 1 — Data Quality Validator.
Rule-based, LLM-free, edge-deployable.
Tags each record CLEAN / SUSPECT / REJECTED.
"""

import logging
from dataclasses import dataclass
from core.schema import (
    CLEAN, SUSPECT, REJECTED,
    PF_PENALTY_THRESHOLD, VOLTAGE_MIN, VOLTAGE_MAX,
    COL_SITE_ID, COL_TIMESTAMP, COL_KW, COL_KVAR, COL_PF, COL_VOLTAGE,
)
from core.parser import normalise_timestamp

log = logging.getLogger("hyesys.agent1.validator")


@dataclass
class ValidationResult:
    tag:    str          # CLEAN / SUSPECT / REJECTED
    reason: str | None   # None for CLEAN


class Validator:
    """
    Stateful validator — tracks seen timestamps per site to detect duplicates
    and data gaps across a batch of records.
    """

    def __init__(self):
        self._seen: dict[str, set[str]] = {}   # site_id → set of timestamps
        self._last_ts: dict[str, str]   = {}   # site_id → last timestamp string

    def validate(self, raw: dict) -> tuple[dict, ValidationResult]:
        """
        Validates a single raw row dict (from parser.parse_csv).
        Returns (normalised_record, ValidationResult).
        """
        site_id = raw.get(COL_SITE_ID, "UNKNOWN")

        # ── 1. Timestamp parse ────────────────────────────────────
        ts_raw = raw.get(COL_TIMESTAMP, "")
        ts_norm = normalise_timestamp(ts_raw) if ts_raw else None
        if not ts_norm:
            return self._reject(raw, site_id, ts_norm, "unparseable timestamp")

        # ── 2. Duplicate timestamp ────────────────────────────────
        seen = self._seen.setdefault(site_id, set())
        if ts_norm in seen:
            return self._reject(raw, site_id, ts_norm, f"duplicate timestamp: {ts_norm}")
        seen.add(ts_norm)

        # ── 3. Parse numeric fields ───────────────────────────────
        try:
            kw      = float(raw.get(COL_KW,      0) or 0)
            kvar    = float(raw.get(COL_KVAR,     0) or 0)
            pf      = float(raw.get(COL_PF,       0) or 0)
            voltage = float(raw.get(COL_VOLTAGE,  0) or 0)
        except (ValueError, TypeError):
            return self._reject(raw, site_id, ts_norm, "non-numeric field value")

        # ── 4. Physically impossible values ──────────────────────
        if voltage <= 0:
            return self._reject(raw, site_id, ts_norm, f"voltage <= 0: {voltage}")
        if not (-10_000 <= kw <= 10_000):
            return self._reject(raw, site_id, ts_norm, f"kW out of range: {kw}")
        if not (-5_000 <= kvar <= 5_000):
            return self._reject(raw, site_id, ts_norm, f"kVAr out of range: {kvar}")
        if not (-1.0 <= pf <= 1.0):
            return self._reject(raw, site_id, ts_norm, f"PF firmware bug — outside [-1,1]: {pf}")

        # ── 5. All-zero row (meter dropout) ───────────────────────
        if kw == 0.0 and kvar == 0.0 and pf == 0.0:
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 "all-zero row — possible meter dropout")

        # ── 6. PF firmware saturation at exactly ±1.0 ────────────
        if abs(pf) == 1.0:
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 f"PF firmware saturation at {pf}")

        # ── 7. Voltage outside Singapore LV normal range ─────────
        if not (VOLTAGE_MIN <= voltage <= VOLTAGE_MAX):
            return self._suspect(raw, site_id, ts_norm, kw, kvar, pf, voltage,
                                 f"voltage outside 230V ±15%: {voltage}V")

        # ── 8. PF below SP penalty threshold (flag, don't reject) ─
        reason = None
        if 0 < abs(pf) < PF_PENALTY_THRESHOLD:
            reason = f"PF below SP penalty threshold: {pf:.3f}"
            log.warning("[%s] %s — %s", site_id, ts_norm, reason)

        record = self._build(raw, site_id, ts_norm, kw, kvar, pf, voltage, CLEAN, reason)
        self._last_ts[site_id] = ts_norm
        return record, ValidationResult(CLEAN, reason)

    # ── helpers ───────────────────────────────────────────────────

    def _build(self, raw, site_id, ts, kw, kvar, pf, voltage, tag, reason):
        return {
            COL_SITE_ID:   site_id,
            COL_TIMESTAMP: ts,
            COL_KW:        kw,
            COL_KVAR:      kvar,
            COL_PF:        pf,
            COL_VOLTAGE:   voltage,
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
