"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 2 — MASTER TEMPLATE                                          ║
║  HyESys Analysis & Recommendation Engine                            ║
║  Version: 1.0  |  Created: 2026-05-24                               ║
╚══════════════════════════════════════════════════════════════════════╝

PURPOSE
───────
This is the frozen master reference for Agent 2.
Do NOT modify this file.

To deploy Agent 2 for a new site:
  1. Copy this file to  sites/<site_id>/agent2.py
  2. Edit only the SITE CONFIGURATION block below
  3. Optionally override methods marked "OVERRIDE POINT"
  4. Each site's agent2.py is completely independent

RULE / LOGIC INVENTORY
──────────────────────
STATE LAYER
  S1  State struct: kW, kVAr, PF, voltage_V, solar
  S2  Derived properties: kVA=√(P²+Q²), φ=arccos(|PF|), I=S/(3×V_LN),
      tan_phi, q_target, q_correction_needed, loss_fraction_vs_unity_pf,
      recoverable_loss_fraction, is_leading_pf

EVENT LAYER  (4 event types, run on every 15-min state)
  E1  THRESHOLD / PF_PENALTY_RISK    |PF| < 0.85               CRITICAL
  E2  THRESHOLD / PF_LOW             |PF| < PF_TARGET (0.98)   WARNING
  E3  THRESHOLD / PF_LEADING         Q < −5 kVAr               WARNING
  E4  STATISTICAL / PF_ANOMALY_LOW   PF z-score < −3σ          CRITICAL
  E5  STATISTICAL / PF_TREND_DOWN    PF z-score < −2σ          WARNING
  E6  STATISTICAL / DEMAND_SPIKE     kW z-score > +3σ          CRITICAL
  E7  STATISTICAL / DEMAND_ELEVATED  kW z-score > +2σ          WARNING
  E8  STATISTICAL / REACTIVE_SURGE   |kVAr| z-score > +3σ     WARNING
  E9  COMPOSITE / LOW_PF_HIGH_DEMAND PF<target AND kW≥90% peak CRITICAL
  E10 COMPOSITE / PF_VOLTAGE_SAG     PF<target AND V<220V      WARNING
  E11 SCHEDULED / PEAK_PERIOD        hour 08:00–22:00           INFO
  E12 SCHEDULED / OFFPEAK_PERIOD     hour 22:00–08:00           INFO

DECISION LAYER  (PI controller)
  D1  Priority 1 — Solar storage / demand shaving
        If has_solar AND demand==CRITICAL → ACTION_REDUCE (reserve capacity)
  D2  Dead-band
        |e(t)| < DEADBAND (0.005) → ACTION_HOLD
  D3  Proportional term
        ΔQ_P = K_P × P × (tan φ_current − tan φ_target)   K_P = 1.00
  D4  Integral term (anti-windup)
        I(t) = clamp(I(t−1) + e(t)×Δt, −I_MAX, +I_MAX)
        ΔQ_I = K_I × I(t)    K_I = 0.50   I_MAX = 20 kVAr·hr
  D5  Combined command
        ΔQ_cmd = ΔQ_P + ΔQ_I
        Clamp to model hardware kVA capacity
  D6  Action selection
        ΔQ_cmd > +0.5 → ACTION_INJECT   (capacitive)
        ΔQ_cmd < −0.5 → ACTION_REDUCE   (inductive)
        otherwise     → ACTION_HOLD

REWARD LAYER
  RW1 PF improvement:     r_PF  = PF_after − PF_before
  RW2 Loss fraction:      r_loss = 1 − (S_after/S_before)²
  RW3 Weighted reward:    r_total = 0.60×r_PF + 0.40×r_loss
  RW4 THD back-calc:      THD = √((PF_before/PF_target)²/(1−f) − 1)
  RW5 Outcome:            r_PF ≥ +0.01 → POSITIVE
                          r_PF ≤ −0.01 → NEGATIVE
                          else         → NEUTRAL

WHAT IS NOT IN THE MASTER (requires site-level customisation):
  • Site-specific demand tariff (non-Singapore tariff structure)
  • Peak/off-peak hours for non-Singapore deployments
  • SCADA/MQTT command interface for live deployment
  • Site-specific PI tuning (K_P, K_I, DEADBAND) if default is unsuitable
  • Harmonic compensation mode logic (future HyESys V2 feature)
"""

import math
import logging
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import datetime

from core.schema import (
    ACTION_INJECT, ACTION_HOLD, ACTION_REDUCE,
    PF_TARGET, PF_PENALTY_THRESHOLD,
    HYESYS_MODELS, THD_ASSUMPTION,
    CLEAN, SUSPECT,
    REWARD_POSITIVE_PF_DELTA, REWARD_NEGATIVE_PF_DELTA,
)
from core.store import write_sar

log = logging.getLogger("hyesys.agent2")
SQRT3 = math.sqrt(3)


# ══════════════════════════════════════════════════════════════════════
# SITE CONFIGURATION  — edit this block when copying to a new site
# ══════════════════════════════════════════════════════════════════════

SITE_CONFIG = {
    # ── Identity ──────────────────────────────────────────────────
    "site_id":            "TEMPLATE",
    "has_solar":          True,
    "hyesys_model":       "H50",    # H30 / H50 / H60 / H100 / H125

    # ── Controller tuning ─────────────────────────────────────────
    # K_P = 1.0: full correction in one 15-min step (analytical solution)
    # K_P < 1.0: ramp up — smoother, fewer oscillations
    "K_P":      1.00,
    "K_I":      0.50,    # integral gain (conservative for 15-min resolution)
    "I_MAX":   20.0,     # anti-windup clamp [kVAr·hr]
    "DEADBAND": 0.005,   # |e(t)| < DEADBAND → HOLD (suppresses micro-oscillations)
    "DT_HOURS": 0.25,    # interval in hours (15 min = 0.25 hr)

    # ── Demand risk thresholds ─────────────────────────────────────
    # Adjust if site uses a non-Singapore demand tariff structure
    "demand_critical_pct":   95,   # % of historical peak → CRITICAL
    "demand_high_pct":       85,   # % of historical peak → HIGH
    "demand_medium_pct":     70,   # % of historical peak → MEDIUM
    "demand_tariff_sgd_per_kw": 10.0,  # SGD/kW/month demand charge estimate

    # ── Event thresholds ──────────────────────────────────────────
    "z_warn":         2.0,    # z-score warning threshold (statistical events)
    "z_crit":         3.0,    # z-score critical threshold
    "ema_n":          8,      # EMA window (8 × 15-min = 2 hours)
    "leading_pf_q_threshold": -5.0,   # kVAr — below this → PF_LEADING event
    "voltage_sag_V":  220.0,  # V — below this with low PF → PF_VOLTAGE_SAG
    "high_demand_pct": 0.90,  # fraction of recent peak → HIGH_DEMAND composite event

    # ── Scheduled events peak/off-peak ────────────────────────────
    # Singapore default: peak 08:00–22:00
    "peak_hours_start": 8,
    "peak_hours_end":   22,

    # ── History window ─────────────────────────────────────────────
    "history_window": 16,    # states retained per site (4 hours of 15-min)
}


# ══════════════════════════════════════════════════════════════════════
# STATE
# ══════════════════════════════════════════════════════════════════════

@dataclass
class State:
    """
    15-minute snapshot of site electrical conditions.

    Equations embedded as properties:
      S  = √(P² + Q²)                     apparent power [kVA]
      PF = P / S = cos(φ)                 power factor
      φ  = arccos(|PF|)                   PF angle [rad]
      I  = S / (3 × V_LN)                 3-phase line current [A]
      f  = 1 − (S_after/S_before)²        I²R loss reduction fraction
    """
    site_id:   str
    timestamp: str
    kW:        float
    kVAr:      float    # negative = leading (capacitive)
    PF:        float    # signed: negative = leading
    voltage_V: float    # line-to-neutral RMS [V]
    solar:     bool = False

    @property
    def kVA(self) -> float:
        """S = √(P² + Q²)"""
        return math.sqrt(self.kW ** 2 + self.kVAr ** 2)

    @property
    def phi_rad(self) -> float:
        """φ = arccos(|PF|)"""
        return math.acos(min(abs(self.PF), 1.0))

    @property
    def phi_deg(self) -> float:
        return math.degrees(self.phi_rad)

    @property
    def tan_phi(self) -> float:
        """tan(φ) — used in ΔQ = P × (tan φ_current − tan φ_target)"""
        return math.tan(self.phi_rad)

    @property
    def q_target(self) -> float:
        """Q_target = P × tan(arccos(PF_target))"""
        return self.kW * math.tan(math.acos(PF_TARGET))

    @property
    def q_correction_needed(self) -> float:
        """ΔQ = Q_current − Q_target  (+) = lagging, (−) = leading"""
        return round(self.kVAr - self.q_target, 3)

    @property
    def current_A(self) -> float:
        """I = S / (3 × V_LN)  [3-phase, line-to-neutral]"""
        return (self.kVA * 1000.0) / (3.0 * self.voltage_V) if self.voltage_V > 0 else 0.0

    @property
    def loss_fraction_vs_unity_pf(self) -> float:
        """Excess I²R fraction vs unity PF = 1 − PF²"""
        pf_abs = abs(self.PF)
        return round(1.0 - pf_abs ** 2, 4) if pf_abs > 0 and self.kVA > 0 else 0.0

    @property
    def recoverable_loss_fraction(self) -> float:
        """Fraction of I²R losses recoverable by correcting to PF_TARGET"""
        if self.kW <= 0 or self.kVA <= 0:
            return 0.0
        s_target = self.kW / PF_TARGET
        return round(max(1.0 - (s_target / self.kVA) ** 2, 0.0), 4)

    @property
    def is_leading_pf(self) -> bool:
        return self.kVAr < 0

    @property
    def pf_abs(self) -> float:
        return abs(self.PF)

    @property
    def hour(self) -> int:
        try:
            return datetime.fromisoformat(self.timestamp).hour
        except ValueError:
            return 0

    def __repr__(self):
        return (f"State({self.site_id} @ {self.timestamp} | "
                f"P={self.kW:.1f}kW Q={self.kVAr:.1f}kVAr "
                f"PF={self.PF:.3f} φ={self.phi_deg:.1f}°)")


def build_state(row, site_config: dict = SITE_CONFIG) -> State:
    d = dict(row) if hasattr(row, "keys") else row
    return State(
        site_id   = d.get("site_id",   site_config.get("site_id", "UNKNOWN")),
        timestamp = d.get("timestamp", ""),
        kW        = float(d.get("kW",        0) or 0),
        kVAr      = float(d.get("kVAr",      0) or 0),
        PF        = float(d.get("PF",        0) or 0),
        voltage_V = float(d.get("voltage_V", 0) or 0),
        solar     = site_config.get("has_solar", True),
    )


# ══════════════════════════════════════════════════════════════════════
# EVENTS
# ══════════════════════════════════════════════════════════════════════

@dataclass
class Event:
    type:        str     # THRESHOLD / STATISTICAL / COMPOSITE / SCHEDULED
    subtype:     str
    severity:    str     # INFO / WARNING / CRITICAL
    site_id:     str
    timestamp:   str
    description: str
    data:        dict = field(default_factory=dict)


class EMATracker:
    """
    Exponential moving average + variance tracker.

    EMA update:
      δ   = x − μ_prev
      μ_t = μ_prev + α × δ
      σ²_t = (1−α) × (σ²_prev + α × δ²)
      z   = (x − μ_t) / σ_t
    """
    def __init__(self, alpha: float):
        self.alpha   = alpha
        self.mu      = None
        self.sigma2  = None
        self.n       = 0

    def update(self, x: float) -> tuple[float, float, float]:
        if self.mu is None:
            self.mu, self.sigma2, self.n = x, 0.0, 1
            return x, 0.0, 0.0
        delta       = x - self.mu
        self.mu     = self.mu + self.alpha * delta
        self.sigma2 = (1 - self.alpha) * (self.sigma2 + self.alpha * delta ** 2)
        self.n     += 1
        sigma   = math.sqrt(self.sigma2) if self.sigma2 > 0 else 0.0
        z_score = (x - self.mu) / sigma if sigma > 0 else 0.0
        return round(self.mu, 4), round(sigma, 4), round(z_score, 3)

    @property
    def sigma(self) -> float:
        return math.sqrt(self.sigma2) if self.sigma2 and self.sigma2 > 0 else 0.0


class SiteStats:
    def __init__(self, alpha: float):
        self.pf_tracker   = EMATracker(alpha)
        self.kw_tracker   = EMATracker(alpha)
        self.kvar_tracker = EMATracker(alpha)


def check_threshold_events(state: State, cfg: dict) -> list[Event]:
    """
    E1–E3: Fixed-threshold events.

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override to add site-specific threshold events, e.g.:
      - Current overload (I > rated amps of local feeder)
      - MV voltage deviation thresholds
    """
    events = []
    pf, Q  = state.pf_abs, state.kVAr

    if 0 < pf < PF_PENALTY_THRESHOLD:                                   # E1
        events.append(Event("THRESHOLD", "PF_PENALTY_RISK", "CRITICAL",
            state.site_id, state.timestamp,
            f"PF={pf:.3f} below SP penalty threshold {PF_PENALTY_THRESHOLD}. "
            f"φ={state.phi_deg:.1f}° Q={Q:+.1f} kVAr ΔQ={state.q_correction_needed:+.1f} kVAr",
            {"PF": pf, "threshold": PF_PENALTY_THRESHOLD,
             "phi_deg": state.phi_deg, "Q_needed": state.q_correction_needed}))

    elif 0 < pf < PF_TARGET:                                             # E2
        events.append(Event("THRESHOLD", "PF_LOW", "WARNING",
            state.site_id, state.timestamp,
            f"PF={pf:.3f} below target {PF_TARGET}. "
            f"ΔQ={state.q_correction_needed:+.1f} kVAr needed. "
            f"Recoverable loss={state.recoverable_loss_fraction:.4f}",
            {"PF": pf, "target": PF_TARGET,
             "Q_needed": state.q_correction_needed,
             "recoverable_fraction": state.recoverable_loss_fraction}))

    if Q < cfg["leading_pf_q_threshold"]:                               # E3
        events.append(Event("THRESHOLD", "PF_LEADING", "WARNING",
            state.site_id, state.timestamp,
            f"Leading (capacitive) PF: Q={Q:.1f} kVAr. "
            f"I²R excess={state.loss_fraction_vs_unity_pf*100:.2f}%",
            {"kVAr": Q, "loss_pct": state.loss_fraction_vs_unity_pf * 100}))

    return events


def check_statistical_events(state: State, stats: SiteStats, cfg: dict) -> list[Event]:
    """
    E4–E8: EMA z-score events.

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override to monitor additional signals (e.g. phase imbalance,
    temperature, frequency deviation).
    """
    events = []
    z_warn = cfg["z_warn"]
    z_crit = cfg["z_crit"]

    pf_abs = state.pf_abs
    if pf_abs > 0:
        mu_pf, sigma_pf, z_pf = stats.pf_tracker.update(pf_abs)
        if sigma_pf > 0:
            if z_pf < -z_crit:                                           # E4
                events.append(Event("STATISTICAL", "PF_ANOMALY_LOW", "CRITICAL",
                    state.site_id, state.timestamp,
                    f"PF={pf_abs:.3f} is {abs(z_pf):.1f}σ below EMA. "
                    f"EMA={mu_pf:.3f} ± {sigma_pf:.4f}  z={z_pf:.2f}",
                    {"PF": pf_abs, "EMA_PF": mu_pf, "sigma": sigma_pf, "z": z_pf}))
            elif z_pf < -z_warn:                                         # E5
                events.append(Event("STATISTICAL", "PF_TREND_DOWN", "WARNING",
                    state.site_id, state.timestamp,
                    f"PF trending down: z={z_pf:.2f}σ. PF={pf_abs:.3f} EMA={mu_pf:.3f}",
                    {"PF": pf_abs, "EMA_PF": mu_pf, "sigma": sigma_pf, "z": z_pf}))

    if state.kW > 0:
        mu_kw, sigma_kw, z_kw = stats.kw_tracker.update(state.kW)
        if sigma_kw > 0:
            if z_kw > z_crit:                                            # E6
                events.append(Event("STATISTICAL", "DEMAND_SPIKE", "CRITICAL",
                    state.site_id, state.timestamp,
                    f"Demand spike: P={state.kW:.1f} kW is {z_kw:.1f}σ above EMA. "
                    f"EMA={mu_kw:.1f} ± {sigma_kw:.1f}  z={z_kw:.2f}",
                    {"kW": state.kW, "EMA_kW": mu_kw, "sigma": sigma_kw, "z": z_kw}))
            elif z_kw > z_warn:                                          # E7
                events.append(Event("STATISTICAL", "DEMAND_ELEVATED", "WARNING",
                    state.site_id, state.timestamp,
                    f"Elevated demand: P={state.kW:.1f} kW z={z_kw:.2f}σ EMA={mu_kw:.1f} kW",
                    {"kW": state.kW, "EMA_kW": mu_kw, "sigma": sigma_kw, "z": z_kw}))

    kvar_abs = abs(state.kVAr)
    if kvar_abs > 0:
        mu_q, sigma_q, z_q = stats.kvar_tracker.update(kvar_abs)
        if sigma_q > 0 and z_q > z_crit:                                # E8
            events.append(Event("STATISTICAL", "REACTIVE_SURGE", "WARNING",
                state.site_id, state.timestamp,
                f"Reactive surge: |Q|={kvar_abs:.1f} kVAr is {z_q:.1f}σ above EMA. "
                f"EMA={mu_q:.1f} ± {sigma_q:.1f}",
                {"kVAr_abs": kvar_abs, "EMA_kVAr": mu_q, "sigma": sigma_q, "z": z_q}))

    return events


def check_composite_events(state: State, history: list[State], cfg: dict) -> list[Event]:
    """
    E9–E10: Multi-condition composite events.

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override to add site-specific composite triggers, e.g.:
      - Temperature + high load combination for cooling-sensitive sites
      - MV demand + local LV PF combined risk
    """
    events = []
    pf_low = 0 < state.pf_abs < PF_TARGET

    recent_kw   = [s.kW for s in history[-4:]] if history else []
    recent_peak = max(recent_kw, default=state.kW)
    high_demand = (state.kW > 0 and recent_peak > 0
                   and (state.kW / recent_peak) >= cfg["high_demand_pct"])

    if pf_low and high_demand:                                           # E9
        events.append(Event("COMPOSITE", "LOW_PF_HIGH_DEMAND", "CRITICAL",
            state.site_id, state.timestamp,
            f"Low PF ({state.pf_abs:.3f}) + high demand ({state.kW:.1f} kW = "
            f"{state.kW/recent_peak*100:.0f}% of recent peak {recent_peak:.1f} kW). "
            f"Combined I²R excess ≈ {state.loss_fraction_vs_unity_pf*100:.2f}%",
            {"PF": state.pf_abs, "kW": state.kW,
             "recent_peak_kW": recent_peak,
             "loss_pct": state.loss_fraction_vs_unity_pf * 100}))

    if pf_low and state.voltage_V < cfg["voltage_sag_V"]:               # E10
        events.append(Event("COMPOSITE", "PF_VOLTAGE_SAG", "WARNING",
            state.site_id, state.timestamp,
            f"Low PF ({state.pf_abs:.3f}) with voltage sag ({state.voltage_V:.1f} V "
            f"< {cfg['voltage_sag_V']} V). High reactive current causing resistive drop.",
            {"PF": state.pf_abs, "voltage_V": state.voltage_V}))

    return events


def check_scheduled_events(state: State, cfg: dict) -> list[Event]:
    """
    E11–E12: Time-based storage scheduling events.

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override peak_hours_start / peak_hours_end in SITE_CONFIG for
    non-Singapore tariff structures.
    """
    hour       = state.hour
    peak_start = cfg["peak_hours_start"]
    peak_end   = cfg["peak_hours_end"]
    is_peak    = peak_start <= hour < peak_end

    if is_peak:
        return [Event("SCHEDULED", "PEAK_PERIOD", "INFO",
            state.site_id, state.timestamp,
            f"On-peak period (hour={hour:02d}:00). Demand charges active — "
            f"prioritise storage discharge.",
            {"hour": hour, "period": "peak"})]
    return [Event("SCHEDULED", "OFFPEAK_PERIOD", "INFO",
        state.site_id, state.timestamp,
        f"Off-peak period (hour={hour:02d}:00). Opportunity to charge storage.",
        {"hour": hour, "period": "offpeak"})]


SEVERITY_ORDER = {"CRITICAL": 0, "WARNING": 1, "INFO": 2}


# ══════════════════════════════════════════════════════════════════════
# TOOLS
# ══════════════════════════════════════════════════════════════════════

_PHI_TARGET     = math.acos(PF_TARGET)        # arccos(0.98) ≈ 0.1997 rad
_TAN_PHI_TARGET = math.tan(_PHI_TARGET)       # ≈ 0.2031


def compute_pf_correction(state: State, model: str, cfg: dict = SITE_CONFIG) -> dict:
    """
    Analytical kVAr injection calculation.

    Step 1: Q_target = P × tan(arccos(PF_target)) = P × 0.2031
    Step 2: ΔQ = Q_current − Q_target
    Step 3: Clamp ΔQ to [−model_kVA, +model_kVA]
    Step 4: Q_after = Q − ΔQ_injected
            S_after = √(P² + Q_after²)
            PF_after = P / S_after
    Step 5: f = 1 − (S_after / S_before)²

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override for sites with multiple HyESys units (aggregate capacity)
    or non-standard model sizing.
    """
    P, Q, S, PF = state.kW, state.kVAr, state.kVA, state.pf_abs
    model_kva   = HYESYS_MODELS.get(model, {}).get("kVA", 50)

    if P == 0 or S == 0:
        return {"P_kW": 0, "Q_kVAr_before": 0, "S_kVA_before": 0,
                "PF_before": 0, "delta_Q_required": 0, "delta_Q_injected": 0,
                "Q_kVAr_after": 0, "S_kVA_after": 0, "PF_achievable": 1.0,
                "savings_fraction": 0.0, "kW_saved_est": 0.0,
                "at_capacity": False, "model": model, "model_kVA": model_kva}

    Q_target          = P * _TAN_PHI_TARGET
    delta_Q           = Q - Q_target
    delta_Q_clamped   = max(-model_kva, min(model_kva, delta_Q))
    at_capacity       = abs(delta_Q) > model_kva
    Q_after           = Q - delta_Q_clamped
    S_after           = math.sqrt(P ** 2 + Q_after ** 2)
    PF_after          = (P / S_after) if S_after > 0 else 1.0
    fraction          = max(1.0 - (S_after / S) ** 2, 0.0) if S > 0 else 0.0
    kW_saved_est      = fraction * P * THD_ASSUMPTION

    return {
        "P_kW":             round(P,               2),
        "Q_kVAr_before":    round(Q,               2),
        "S_kVA_before":     round(S,               2),
        "PF_before":        round(PF,              4),
        "phi_deg_before":   round(math.degrees(math.acos(min(PF, 1.0))), 2),
        "Q_target_kVAr":    round(Q_target,        2),
        "delta_Q_required": round(delta_Q,         2),
        "delta_Q_injected": round(delta_Q_clamped, 2),
        "at_capacity":      at_capacity,
        "Q_kVAr_after":     round(Q_after,         2),
        "S_kVA_after":      round(S_after,         2),
        "PF_achievable":    round(PF_after,         4),
        "phi_deg_after":    round(math.degrees(math.acos(min(PF_after, 1.0))), 2),
        "savings_fraction": round(fraction,         4),
        "kW_saved_est":     round(max(kW_saved_est, 0.0), 3),
        "model":            model,
        "model_kVA":        model_kva,
    }


def assess_demand_risk(state: State, peak_kw_history: list[float],
                       cfg: dict = SITE_CONFIG) -> dict:
    """
    Peak demand risk assessment.

    demand_pct = (P_current / P_peak_historical) × 100
    Thresholds: CRITICAL ≥95% / HIGH ≥85% / MEDIUM ≥70% / LOW <70%

    ── OVERRIDE POINT ────────────────────────────────────────────────
    Override demand_critical_pct / demand_high_pct in SITE_CONFIG for
    non-Singapore demand charge structures.
    """
    P = state.kW
    if not peak_kw_history:
        return {"current_kW": P, "historical_peak_kW": P, "historical_avg_kW": P,
                "demand_pct": 100.0, "headroom_kW": 0.0,
                "risk_level": "UNKNOWN", "recommend_store": False,
                "demand_charge_risk_sgd": 0.0}

    P_peak      = max(peak_kw_history)
    P_avg       = sum(peak_kw_history) / len(peak_kw_history)
    demand_pct  = (P / P_peak * 100.0) if P_peak > 0 else 0.0
    P_headroom  = P_peak - P
    P_excess    = max(P - P_peak, 0.0)
    charge_risk = P_excess * cfg["demand_tariff_sgd_per_kw"]

    if demand_pct >= cfg["demand_critical_pct"]:
        risk_level, recommend = "CRITICAL", True
    elif demand_pct >= cfg["demand_high_pct"]:
        risk_level, recommend = "HIGH", True
    elif demand_pct >= cfg["demand_medium_pct"]:
        risk_level, recommend = "MEDIUM", state.solar
    else:
        risk_level, recommend = "LOW", False

    return {
        "current_kW":              round(P,           2),
        "historical_peak_kW":      round(P_peak,      2),
        "historical_avg_kW":       round(P_avg,       2),
        "demand_pct":              round(demand_pct,  1),
        "headroom_kW":             round(P_headroom,  2),
        "risk_level":              risk_level,
        "recommend_store":         recommend,
        "demand_charge_risk_sgd":  round(charge_risk, 2),
    }


# ══════════════════════════════════════════════════════════════════════
# REWARD
# ══════════════════════════════════════════════════════════════════════

W_PF   = 0.60   # weight: PF improvement (SP penalty impact)
W_LOSS = 0.40   # weight: I²R loss fraction (kWh savings)


@dataclass
class Reward:
    r_pf:    float
    r_loss:  float
    r_total: float
    thd_est: float | None
    outcome: str

    @property
    def pf_delta(self) -> float:
        return self.r_pf

    @property
    def loss_fraction(self) -> float:
        return self.r_loss


def compute_reward(state_before: State, state_after: State) -> Reward:
    """
    Step 1: r_PF  = PF_after − PF_before
    Step 2: r_loss = 1 − (S_after / S_before)²
    Step 3: r_total = 0.60 × r_PF + 0.40 × r_loss
    Step 4: THD = √((PF_before/PF_target)² / (1−f) − 1)
    Step 5: POSITIVE if r_PF ≥ +0.01 / NEGATIVE if r_PF ≤ −0.01 / else NEUTRAL
    """
    pf_before, pf_after = abs(state_before.PF), abs(state_after.PF)
    r_pf  = round(pf_after - pf_before, 4)

    S_before, S_after = state_before.kVA, state_after.kVA
    r_loss = round(max(1.0 - (S_after / S_before) ** 2, 0.0), 4) if S_before > 0 else 0.0
    r_total = round(W_PF * r_pf + W_LOSS * r_loss, 4)

    thd_est = _back_calculate_thd(pf_before, r_loss)
    outcome = ("POSITIVE" if r_pf >= REWARD_POSITIVE_PF_DELTA
               else "NEGATIVE" if r_pf <= REWARD_NEGATIVE_PF_DELTA
               else "NEUTRAL")

    return Reward(r_pf=r_pf, r_loss=r_loss, r_total=r_total,
                  thd_est=thd_est, outcome=outcome)


def _back_calculate_thd(pf_before: float, loss_fraction: float = 0.0) -> float | None:
    """THD = √((PF_before/PF_target)² / (1−f) − 1)"""
    try:
        if loss_fraction >= 1.0 or pf_before <= 0:
            return None
        inner = (pf_before / PF_TARGET) ** 2 / (1.0 - loss_fraction) - 1.0
        return round(math.sqrt(inner), 4) if inner >= 0 else None
    except (ValueError, ZeroDivisionError):
        return None


# ══════════════════════════════════════════════════════════════════════
# AGENT 2
# ══════════════════════════════════════════════════════════════════════

class Agent2:
    """
    Processes State snapshots → issues injection decisions → logs SAR.

    Per-site state maintained:
      _history  : rolling window of recent States  (EMA / composite events)
      _peak_kw  : historical peak kW values        (demand risk)
      _integral : PI integrator state per site     [kVAr·hr]
      _stats    : EMA trackers per site            (statistical events)
    """

    def __init__(self, conn, cfg: dict = SITE_CONFIG):
        self.conn        = conn
        self.cfg         = cfg
        alpha            = 2.0 / (cfg["ema_n"] + 1)
        hw               = cfg["history_window"]
        self._history:  dict[str, deque]      = defaultdict(lambda: deque(maxlen=hw))
        self._peak_kw:  dict[str, list]       = defaultdict(list)
        self._integral: dict[str, float]      = defaultdict(float)
        self._stats:    dict[str, SiteStats]  = defaultdict(lambda: SiteStats(alpha))

    def process(self, state: State) -> dict:
        """
        Full pipeline per 15-min snapshot:
          1. Detect events
          2. Compute PF correction
          3. Assess demand risk
          4. Decision engine → (action, kVAr)
          5. Simulate post-action state
          6. Compute reward
          7. Write SAR record
          8. Update history

        ── OVERRIDE POINT ────────────────────────────────────────────
        Override to add live MQTT dispatch, site-specific logging,
        or post-processing of the SAR dict.
        """
        site_id = state.site_id
        history = list(self._history[site_id])

        # 1. Events
        events = self._detect_events(state, history)

        # 2. PF correction tool
        pf_tool = compute_pf_correction(state, self.cfg["hyesys_model"], self.cfg)

        # 3. Demand risk
        self._peak_kw[site_id].append(state.kW)
        demand = assess_demand_risk(state, self._peak_kw[site_id], self.cfg)

        # 4. Decision
        action, action_kvar = self._decide(state, events, pf_tool, demand)

        log.info("[%s] %s → %s kVAr=%+.1f | PF=%.3f→%.3f | demand=%s",
                 site_id, state.timestamp, action, action_kvar or 0.0,
                 abs(state.PF), pf_tool["PF_achievable"], demand["risk_level"])

        # 5. Simulate outcome (historical mode)
        state_after = self._simulate_after(state, action, action_kvar)

        # 6. Reward
        reward = compute_reward(state, state_after)

        # 7. SAR record
        sar = {
            "site_id":         site_id,
            "timestamp":       state.timestamp,
            "state_kW":        state.kW,
            "state_kVAr":      state.kVAr,
            "state_PF":        state.PF,
            "state_voltage_V": state.voltage_V,
            "action":          action,
            "action_kVAr":     action_kvar,
            "reward_pf_delta": reward.r_pf,
            "reward_fraction": reward.r_loss,
            "outcome":         reward.outcome,
        }
        write_sar(self.conn, sar)

        # 8. History
        self._history[site_id].append(state)

        return sar

    def process_batch(self, states: list[State]) -> list[dict]:
        return [self.process(s) for s in states]

    # ── Decision engine ────────────────────────────────────────────

    def _decide(self, state: State, events: list[Event],
                pf_tool: dict, demand: dict) -> tuple[str, float | None]:
        """
        PI controller decision.

        D1 Priority 1 — Solar + CRITICAL demand → REDUCE (reserve for storage)
        D2 Dead-band   — |e| < DEADBAND → HOLD
        D3 Proportional — ΔQ_P = K_P × ΔQ_required
        D4 Integral     — I(t) = clamp(I(t−1) + e×Δt, −I_MAX, +I_MAX)
                           ΔQ_I = K_I × I(t)
        D5 Combine      — ΔQ_cmd = ΔQ_P + ΔQ_I, clamp to model capacity
        D6 Action       — >+0.5 kVAr → INJECT; <−0.5 → REDUCE; else HOLD

        ── OVERRIDE POINT ────────────────────────────────────────────
        Override for custom control logic, e.g. time-of-day injection
        limits, harmonic compensation mode, SCADA interlock checks.
        """
        cfg      = self.cfg
        site_id  = state.site_id
        model_kva = HYESYS_MODELS.get(cfg["hyesys_model"], {}).get("kVA", 50)

        # D1 — Solar storage priority
        if state.solar and demand["risk_level"] == "CRITICAL" and demand["recommend_store"]:
            self._integral[site_id] = 0.0
            return ACTION_REDUCE, None

        # D2 — Dead-band
        pf_abs = state.pf_abs
        e_t    = PF_TARGET - pf_abs
        if abs(e_t) < cfg["DEADBAND"]:
            self._integral[site_id] = 0.0
            return ACTION_HOLD, None

        # D3 — Proportional term
        delta_Q_P = cfg["K_P"] * pf_tool["delta_Q_required"]

        # D4 — Integral term
        i_prev = self._integral[site_id]
        i_new  = max(-cfg["I_MAX"], min(cfg["I_MAX"],
                                        i_prev + e_t * cfg["DT_HOURS"]))
        self._integral[site_id] = i_new
        delta_Q_I = cfg["K_I"] * i_new

        # D5 — Combine + clamp
        delta_Q_cmd     = delta_Q_P + delta_Q_I
        delta_Q_clamped = max(-model_kva, min(model_kva, delta_Q_cmd))

        # D6 — Action
        if abs(delta_Q_clamped) < 0.5:
            return ACTION_HOLD, None
        if delta_Q_clamped > 0:
            return ACTION_INJECT, round(delta_Q_clamped, 2)
        return ACTION_REDUCE, round(abs(delta_Q_clamped), 2)

    # ── State simulator (historical / backtest mode) ───────────────

    def _simulate_after(self, state: State, action: str,
                        action_kvar: float | None) -> State:
        """
        Estimates post-injection state for reward computation.

        INJECT:  Q_after = Q_before − ΔQ       (capacitive, reduces lagging kVAr)
        REDUCE:  Q_after = Q_before − ΔQ (with sign)  or  × 0.5 fallback
        HOLD:    Q_after = Q_before

        In live deployment: replace with actual next 15-min meter reading.

        ── OVERRIDE POINT ────────────────────────────────────────────
        Override when live meter feedback is available, or to model
        site-specific transformer/cable impedance effects.
        """
        if action == ACTION_INJECT and action_kvar is not None:
            kvar_after = state.kVAr - action_kvar
        elif action == ACTION_REDUCE and action_kvar is not None:
            kvar_after = state.kVAr - action_kvar
        elif action == ACTION_REDUCE:
            kvar_after = state.kVAr * 0.5
        else:
            kvar_after = state.kVAr

        kva_after = math.sqrt(state.kW ** 2 + kvar_after ** 2)
        pf_after  = abs(state.kW / kva_after) if kva_after > 0 else 1.0

        return State(site_id=state.site_id, timestamp=state.timestamp,
                     kW=state.kW, kVAr=kvar_after, PF=pf_after,
                     voltage_V=state.voltage_V, solar=state.solar)

    # ── Internal event detection ───────────────────────────────────

    def _detect_events(self, state: State, history: list[State]) -> list[Event]:
        stats  = self._stats[state.site_id]
        events = (
            check_threshold_events(state, self.cfg)
            + check_statistical_events(state, stats, self.cfg)
            + check_composite_events(state, history, self.cfg)
            + check_scheduled_events(state, self.cfg)
        )
        events.sort(key=lambda e: SEVERITY_ORDER.get(e.severity, 9))
        return events
