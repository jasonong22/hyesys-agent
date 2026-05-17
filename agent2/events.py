"""
Event taxonomy for Agent 2.
Four event types: Threshold, Statistical, Composite, Scheduled.

STATISTICAL EVENT EQUATIONS
─────────────────────────────
Exponential Moving Average (EMA):
  μ_t = α × x_t + (1 − α) × μ_{t−1}         α = 2/(N+1)

EMA Variance (Welford-style exponential):
  σ²_t = (1 − α) × (σ²_{t−1} + α × (x_t − μ_{t−1})²)

Z-score (standardised deviation):
  z_t = (x_t − μ_t) / σ_t

Event triggers:
  |z| > Z_WARN  → WARNING
  |z| > Z_CRIT  → CRITICAL
"""

import math
import logging
from dataclasses import dataclass, field
from datetime import datetime
from agent2.state import State
from core.schema import PF_TARGET, PF_PENALTY_THRESHOLD

log = logging.getLogger("hyesys.agent2.events")

# ── Statistical thresholds ────────────────────────────────────────
Z_WARN  = 2.0    # z-score warning threshold
Z_CRIT  = 3.0    # z-score critical threshold
EMA_N   = 8      # EMA window length (number of 15-min intervals = 2 hours)
EMA_ALPHA = 2.0 / (EMA_N + 1)   # α = 2/(N+1) ≈ 0.222

# Peak / off-peak hour boundaries (Singapore)
PEAK_HOURS    = set(range(8, 22))        # 08:00–21:59
OFFPEAK_HOURS = set(range(0, 8)) | {22, 23}


# ── Event dataclass ───────────────────────────────────────────────

@dataclass
class Event:
    type:        str        # THRESHOLD / STATISTICAL / COMPOSITE / SCHEDULED
    subtype:     str
    severity:    str        # INFO / WARNING / CRITICAL
    site_id:     str
    timestamp:   str
    description: str
    data:        dict = field(default_factory=dict)


# ── EMA tracker (per-site, per-variable) ─────────────────────────

class EMATracker:
    """
    Maintains an exponential moving average and variance for a scalar signal.

    EMA update:
      μ_t = α × x_t + (1 − α) × μ_{t−1}

    EMA variance update:
      δ   = x_t − μ_{t−1}
      μ_t = μ_{t−1} + α × δ
      σ²_t = (1 − α) × (σ²_{t−1} + α × δ²)

    Z-score:
      z = (x_t − μ_t) / σ_t     if σ_t > 0, else 0
    """
    def __init__(self, alpha: float = EMA_ALPHA):
        self.alpha   = alpha
        self.mu      = None    # EMA mean
        self.sigma2  = None    # EMA variance
        self.n       = 0       # observations seen

    def update(self, x: float) -> tuple[float, float, float]:
        """
        Update EMA with new observation x.
        Returns (mu, sigma, z_score).
        """
        if self.mu is None:
            self.mu     = x
            self.sigma2 = 0.0
            self.n      = 1
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


# ── Per-site EMA state store ─────────────────────────────────────

class SiteStats:
    """Holds EMA trackers for key signals per site."""
    def __init__(self):
        self.pf_tracker     = EMATracker()
        self.kw_tracker     = EMATracker()
        self.kvar_tracker   = EMATracker()

_site_stats: dict[str, SiteStats] = {}

def _get_stats(site_id: str) -> SiteStats:
    if site_id not in _site_stats:
        _site_stats[site_id] = SiteStats()
    return _site_stats[site_id]


# ── 1. THRESHOLD EVENTS ───────────────────────────────────────────

def check_threshold_events(state: State) -> list[Event]:
    """
    Fires when a measured value crosses a fixed threshold.

    Thresholds:
      PF < 0.85           → CRITICAL (SP penalty zone)
      PF < 0.98           → WARNING  (below target)
      Q < −5 kVAr         → WARNING  (leading PF — capacitive load)
      |PF| == 1.0 exactly → WARNING  (firmware saturation artefact)
    """
    events = []
    pf = state.pf_abs
    Q  = state.kVAr

    if 0 < pf < PF_PENALTY_THRESHOLD:
        events.append(Event(
            type="THRESHOLD", subtype="PF_PENALTY_RISK", severity="CRITICAL",
            site_id=state.site_id, timestamp=state.timestamp,
            description=(
                f"PF={pf:.3f} below SP penalty threshold {PF_PENALTY_THRESHOLD}. "
                f"φ={state.phi_deg:.1f}°  Q={Q:+.1f} kVAr  "
                f"ΔQ_needed={state.q_correction_needed:+.1f} kVAr"
            ),
            data={"PF": pf, "threshold": PF_PENALTY_THRESHOLD,
                  "phi_deg": state.phi_deg, "Q_needed": state.q_correction_needed},
        ))

    elif 0 < pf < PF_TARGET:
        events.append(Event(
            type="THRESHOLD", subtype="PF_LOW", severity="WARNING",
            site_id=state.site_id, timestamp=state.timestamp,
            description=(
                f"PF={pf:.3f} below target {PF_TARGET}. "
                f"ΔQ={state.q_correction_needed:+.1f} kVAr needed. "
                f"Recoverable loss fraction={state.recoverable_loss_fraction:.4f}"
            ),
            data={"PF": pf, "target": PF_TARGET,
                  "Q_needed": state.q_correction_needed,
                  "recoverable_fraction": state.recoverable_loss_fraction},
        ))

    if Q < -5.0:
        events.append(Event(
            type="THRESHOLD", subtype="PF_LEADING", severity="WARNING",
            site_id=state.site_id, timestamp=state.timestamp,
            description=(
                f"Leading (capacitive) PF: Q={Q:.1f} kVAr. "
                f"Leading PF is as harmful as lagging — "
                f"I²R excess = {state.loss_fraction_vs_unity_pf*100:.2f}%"
            ),
            data={"kVAr": Q, "loss_pct": state.loss_fraction_vs_unity_pf * 100},
        ))

    return events


# ── 2. STATISTICAL EVENTS ─────────────────────────────────────────

def check_statistical_events(state: State, history: list[State]) -> list[Event]:
    """
    Fires when a signal deviates significantly from its EMA baseline.

    Z-score method:
      z = (x_t − μ_t) / σ_t

    Monitored signals:
      PF    — sustained downward trend
      kW    — demand spike above EMA baseline
      kVAr  — reactive load surge
    """
    events = []
    stats  = _get_stats(state.site_id)

    # ── PF z-score ────────────────────────────────────────────────
    pf_abs = state.pf_abs
    if pf_abs > 0:
        mu_pf, sigma_pf, z_pf = stats.pf_tracker.update(pf_abs)

        if sigma_pf > 0 and z_pf < -Z_CRIT:
            events.append(Event(
                type="STATISTICAL", subtype="PF_ANOMALY_LOW", severity="CRITICAL",
                site_id=state.site_id, timestamp=state.timestamp,
                description=(
                    f"PF={pf_abs:.3f} is {abs(z_pf):.1f}σ below EMA baseline. "
                    f"EMA_PF={mu_pf:.3f} ± {sigma_pf:.4f}. "
                    f"z = (PF − μ) / σ = ({pf_abs:.3f} − {mu_pf:.3f}) / {sigma_pf:.4f} = {z_pf:.2f}"
                ),
                data={"PF": pf_abs, "EMA_PF": mu_pf, "sigma": sigma_pf, "z": z_pf},
            ))
        elif sigma_pf > 0 and z_pf < -Z_WARN:
            events.append(Event(
                type="STATISTICAL", subtype="PF_TREND_DOWN", severity="WARNING",
                site_id=state.site_id, timestamp=state.timestamp,
                description=(
                    f"PF trending down: z={z_pf:.2f}σ. "
                    f"PF={pf_abs:.3f}  EMA={mu_pf:.3f}  σ={sigma_pf:.4f}"
                ),
                data={"PF": pf_abs, "EMA_PF": mu_pf, "sigma": sigma_pf, "z": z_pf},
            ))

    # ── kW demand z-score ─────────────────────────────────────────
    if state.kW > 0:
        mu_kw, sigma_kw, z_kw = stats.kw_tracker.update(state.kW)

        if sigma_kw > 0 and z_kw > Z_CRIT:
            events.append(Event(
                type="STATISTICAL", subtype="DEMAND_SPIKE", severity="CRITICAL",
                site_id=state.site_id, timestamp=state.timestamp,
                description=(
                    f"Demand spike: P={state.kW:.1f} kW is {z_kw:.1f}σ above EMA. "
                    f"EMA_kW={mu_kw:.1f} ± {sigma_kw:.1f}. "
                    f"z = ({state.kW:.1f} − {mu_kw:.1f}) / {sigma_kw:.1f} = {z_kw:.2f}"
                ),
                data={"kW": state.kW, "EMA_kW": mu_kw, "sigma": sigma_kw, "z": z_kw},
            ))
        elif sigma_kw > 0 and z_kw > Z_WARN:
            events.append(Event(
                type="STATISTICAL", subtype="DEMAND_ELEVATED", severity="WARNING",
                site_id=state.site_id, timestamp=state.timestamp,
                description=(
                    f"Elevated demand: P={state.kW:.1f} kW  z={z_kw:.2f}σ  EMA={mu_kw:.1f} kW"
                ),
                data={"kW": state.kW, "EMA_kW": mu_kw, "sigma": sigma_kw, "z": z_kw},
            ))

    # ── kVAr surge z-score ────────────────────────────────────────
    kvar_abs = abs(state.kVAr)
    if kvar_abs > 0:
        mu_q, sigma_q, z_q = stats.kvar_tracker.update(kvar_abs)

        if sigma_q > 0 and z_q > Z_CRIT:
            events.append(Event(
                type="STATISTICAL", subtype="REACTIVE_SURGE", severity="WARNING",
                site_id=state.site_id, timestamp=state.timestamp,
                description=(
                    f"Reactive surge: |Q|={kvar_abs:.1f} kVAr is {z_q:.1f}σ above EMA. "
                    f"EMA_Q={mu_q:.1f} ± {sigma_q:.1f}"
                ),
                data={"kVAr_abs": kvar_abs, "EMA_kVAr": mu_q, "sigma": sigma_q, "z": z_q},
            ))

    return events


# ── 3. COMPOSITE EVENTS ───────────────────────────────────────────

def check_composite_events(state: State, history: list[State]) -> list[Event]:
    """
    Fires when multiple conditions are simultaneously true.

    Composite logic:
      LOW_PF_HIGH_DEMAND:  PF < target  AND  kW > 90% of recent peak
      PF_VOLTAGE_SAG:      PF < target  AND  V < 220 V
    """
    events = []

    pf_low = 0 < state.pf_abs < PF_TARGET

    # Recent peak from history
    recent_kw = [s.kW for s in history[-4:]] if history else []
    recent_peak = max(recent_kw, default=state.kW)
    high_demand = state.kW > 0 and recent_peak > 0 and (state.kW / recent_peak) >= 0.90

    if pf_low and high_demand:
        events.append(Event(
            type="COMPOSITE", subtype="LOW_PF_HIGH_DEMAND", severity="CRITICAL",
            site_id=state.site_id, timestamp=state.timestamp,
            description=(
                f"Low PF ({state.pf_abs:.3f}) + high demand ({state.kW:.1f} kW = "
                f"{state.kW/recent_peak*100:.0f}% of recent peak {recent_peak:.1f} kW). "
                f"Combined I²R excess ≈ {state.loss_fraction_vs_unity_pf*100:.2f}%"
            ),
            data={"PF": state.pf_abs, "kW": state.kW,
                  "recent_peak_kW": recent_peak, "loss_pct": state.loss_fraction_vs_unity_pf * 100},
        ))

    # Voltage sag under load
    if pf_low and state.voltage_V < 220:
        events.append(Event(
            type="COMPOSITE", subtype="PF_VOLTAGE_SAG", severity="WARNING",
            site_id=state.site_id, timestamp=state.timestamp,
            description=(
                f"Low PF ({state.pf_abs:.3f}) with voltage sag ({state.voltage_V:.1f} V < 220 V). "
                f"High reactive current likely causing resistive voltage drop."
            ),
            data={"PF": state.pf_abs, "voltage_V": state.voltage_V},
        ))

    return events


# ── 4. SCHEDULED EVENTS ───────────────────────────────────────────

def check_scheduled_events(state: State) -> list[Event]:
    """
    Time-based events to guide storage scheduling.

    Singapore tariff:
      On-peak  08:00–22:00 — higher demand charges apply
      Off-peak 22:00–08:00 — opportunity to charge storage
    """
    events = []
    hour = state.hour

    if hour in PEAK_HOURS:
        events.append(Event(
            type="SCHEDULED", subtype="PEAK_PERIOD", severity="INFO",
            site_id=state.site_id, timestamp=state.timestamp,
            description=f"On-peak period (hour={hour:02d}:00). Demand charges active — prioritise storage discharge.",
            data={"hour": hour, "period": "peak"},
        ))
    else:
        events.append(Event(
            type="SCHEDULED", subtype="OFFPEAK_PERIOD", severity="INFO",
            site_id=state.site_id, timestamp=state.timestamp,
            description=f"Off-peak period (hour={hour:02d}:00). Opportunity to charge storage.",
            data={"hour": hour, "period": "offpeak"},
        ))

    return events


# ── Event detector ────────────────────────────────────────────────

SEVERITY_ORDER = {"CRITICAL": 0, "WARNING": 1, "INFO": 2}

def detect_events(state: State, history: list[State]) -> list[Event]:
    """Run all checks and return combined list sorted by severity."""
    events = (
        check_threshold_events(state)
        + check_statistical_events(state, history)
        + check_composite_events(state, history)
        + check_scheduled_events(state)
    )
    events.sort(key=lambda e: SEVERITY_ORDER.get(e.severity, 9))
    for e in events:
        if e.severity != "INFO":
            log.debug("[%s] %s/%s — %s", e.severity, e.type, e.subtype, e.description)
    return events
