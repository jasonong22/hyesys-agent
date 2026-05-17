"""
Agent 2 — Analysis & Recommendation Engine.
Event-driven, 100% local Python, no Claude API at runtime.
Reads 15-min state snapshots from hyesys.db, issues decisions,
and logs STATE → ACTION → REWARD (SAR) triplets.

CONTROL LAW EQUATIONS
─────────────────────
Error signal (PF deviation from target):
  e(t) = PF_target − PF(t)                        [dimensionless, −1 … +1]
  e > 0  → lagging PF, too low     → inject capacitive kVAr
  e < 0  → leading PF, overcorrect → inject inductive kVAr (reduce output)
  e = 0  → at target               → hold

Dead-band suppression (prevents hunting):
  |e(t)| < ε_DEADBAND  → ACTION_HOLD (no injection)
  ε_DEADBAND = 0.005   (0.5 percentage points below target)

Proportional injection (explicit analytical solution):
  ΔQ(t) = P(t) × (tan φ(t) − tan φ_target)       [kVAr]

  This is a proportional controller with gain implicit in the physics:
    gain K_p = P / cos²(φ)  (varies with load)

  Proportional gain applied:
  ΔQ_commanded = K_P × ΔQ_required
    K_P = 1.00  → full correction in one step   (fast but may overshoot)
    K_P = 0.75  → 75% correction per 15-min     (smoother, fewer oscillations)

  K_P is site-configurable; defaults to 1.0 (analytical solution, no ramp-up).

PI integrator (for sustained PF deficit):
  Integral error accumulates when PF stays below target over multiple intervals:
    I(t) = I(t−1) + e(t) × Δt                    [Δt = 0.25 hr per 15-min interval]
    ΔQ_boost = K_I × I(t)                         [additional kVAr, anti-windup clamped]

  Anti-windup: I(t) is clamped to [−I_MAX, +I_MAX] to prevent integrator runaway.
  K_I = 0.5  (integrator gain; tuned conservatively for 15-min resolution)
  I_MAX = 20  (kVAr·hr, equivalent to ~4 hours of 5 kVAr sustained error)

DECISION PRIORITY
──────────────────
  1. Solar storage / demand shaving   (if has_solar AND demand risk CRITICAL/HIGH)
  2. Reactive correction              (PF below target → INJECT or REDUCE)
  3. HOLD                             (within dead-band, PF already at target)
"""

import math
import logging
from collections import defaultdict, deque

from agent2.state import State, build_state
from agent2.tools import compute_pf_correction, assess_demand_risk
from agent2.events import detect_events, Event
from agent2.outcome import compute_reward, Reward
from core.schema import (
    ACTION_INJECT, ACTION_HOLD, ACTION_REDUCE,
    PF_TARGET, PF_PENALTY_THRESHOLD,
    SITE_CONFIG, HYESYS_MODELS,
    CLEAN, SUSPECT,
)
from core.store import write_sar

log = logging.getLogger("hyesys.agent2.agent")

HISTORY_WINDOW  = 16      # states per site for statistical events (~4 hours of 15-min)
DT_HOURS        = 0.25    # 15-min interval in hours (for PI integrator)

# ── Proportional-Integral (PI) controller parameters ──────────────
K_P       = 1.00   # proportional gain (1.0 = full analytical correction)
K_I       = 0.50   # integral gain (anti-windup, conservative for 15-min resolution)
I_MAX     = 20.0   # integrator anti-windup clamp [kVAr·hr]
DEADBAND  = 0.005  # PF dead-band  |e(t)| < 0.005 → HOLD


class Agent2:
    """
    Processes state snapshots and issues injection decisions.

    Maintains per-site state:
      _history    rolling window of recent State objects (EMA / statistical events)
      _peak_kw    list of historical peak kW values (demand risk assessment)
      _integral   PI integrator per site [kVAr·hr]
    """

    def __init__(self, conn, site_models: dict | None = None):
        self.conn         = conn
        self.site_models  = site_models or {}
        self._history:  dict[str, deque] = defaultdict(lambda: deque(maxlen=HISTORY_WINDOW))
        self._peak_kw:  dict[str, list]  = defaultdict(list)
        self._integral: dict[str, float] = defaultdict(float)   # PI integrator state

    def process(self, state: State) -> dict:
        """
        Process one 15-min state snapshot.

        Pipeline:
          1. Detect events (threshold, statistical, composite, scheduled)
          2. Compute PF correction tool output
          3. Assess demand risk
          4. Run decision engine → (action, kVAr_magnitude)
          5. Simulate post-action state (historical mode)
          6. Compute reward
          7. Write SAR record to hyesys.db
          8. Update per-site history

        Returns the SAR dict that was written to the store.
        """
        site_id = state.site_id
        history = list(self._history[site_id])

        # ── 1. Detect events ───────────────────────────────────────
        events = detect_events(state, history)

        # ── 2. Determine model for this site ──────────────────────
        site_cfg  = SITE_CONFIG.get(site_id, {})
        model_key = site_cfg.get("recommended_model", "H50")

        # ── 3. Compute PF correction (analytical ΔQ solution) ─────
        pf_tool = compute_pf_correction(state, model=model_key)

        # ── 4. Assess demand risk ──────────────────────────────────
        self._peak_kw[site_id].append(state.kW)
        demand = assess_demand_risk(state, self._peak_kw[site_id])

        # ── 5. Decision engine ─────────────────────────────────────
        action, action_kvar = self._decide(state, events, pf_tool, demand, site_cfg, site_id)

        log.info(
            "[%s] %s → %s  kVAr=%+.1f | PF=%.3f → %.3f | demand=%s",
            site_id, state.timestamp, action, action_kvar or 0.0,
            abs(state.PF), pf_tool["PF_achievable"], demand["risk_level"],
        )

        # ── 6. Simulate outcome (historical mode — no live feedback) ─
        state_after = self._simulate_after(state, action, action_kvar)
        reward      = compute_reward(state, state_after)

        # ── 7. Build and write SAR record ──────────────────────────
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

        # ── 8. Update per-site history ─────────────────────────────
        self._history[site_id].append(state)

        return sar

    def process_batch(self, states: list[State]) -> list[dict]:
        """Process a list of states in chronological order. Returns all SAR dicts."""
        return [self.process(s) for s in states]

    # ── Decision engine ────────────────────────────────────────────

    def _decide(self, state: State, events: list[Event],
                pf_tool: dict, demand: dict,
                site_cfg: dict, site_id: str) -> tuple[str, float | None]:
        """
        Deterministic PI controller — returns (action, kVAr_magnitude).

        Control law:
          e(t)  = PF_target − PF_current          [error signal]
          |e(t)| < ε_DEADBAND  → HOLD             [suppress micro-oscillations]

          ΔQ_required = pf_tool["delta_Q_required"]   [analytical full-correction kVAr]
          ΔQ_P        = K_P × ΔQ_required             [proportional term]

          Integral update:
            I(t) = clamp(I(t−1) + e(t) × Δt, −I_MAX, +I_MAX)
            ΔQ_I = K_I × I(t)                          [integral boost]

          ΔQ_commanded = ΔQ_P + ΔQ_I
          Clamp to model hardware capacity: ΔQ_clamped = clamp(ΔQ_commanded, −model_kVA, +model_kVA)

        Priority override:
          If has_solar AND demand == CRITICAL → REDUCE (reserve capacity for storage)
        """
        has_solar = site_cfg.get("solar", True)
        model_key = site_cfg.get("recommended_model", "H50")
        model_kva = HYESYS_MODELS.get(model_key, {}).get("kVA", 50)

        # ── Priority 1: Solar storage / demand shaving ─────────────
        # Reserve hardware capacity for peak shaving — more valuable than reactive correction
        if has_solar and demand["risk_level"] == "CRITICAL" and demand["recommend_store"]:
            self._integral[site_id] = 0.0   # reset integrator — mode switch
            return ACTION_REDUCE, None

        # ── Error signal  e(t) = PF_target − PF_current ───────────
        pf_abs = abs(state.PF)
        e_t    = PF_TARGET - pf_abs           # positive → lagging, needs injection
                                              # negative → leading, overcorrected

        # ── Dead-band check  |e(t)| < ε → no action ──────────────
        if abs(e_t) < DEADBAND:
            self._integral[site_id] = 0.0   # reset integrator within dead-band
            return ACTION_HOLD, None

        # ── Proportional term  ΔQ_P = K_P × ΔQ_required ─────────
        # delta_Q_required: P × (tan φ_current − tan φ_target)
        # Positive → lagging load needs capacitive injection
        # Negative → leading load needs inductive injection (REDUCE)
        delta_Q_required = pf_tool["delta_Q_required"]     # signed kVAr
        delta_Q_P        = K_P * delta_Q_required          # proportional correction

        # ── Integral term  I(t) = I(t−1) + e(t) × Δt ─────────────
        # Accumulates error across 15-min intervals to handle persistent PF deficit.
        # Anti-windup: clamp integral to [−I_MAX, +I_MAX]
        i_prev = self._integral[site_id]
        i_new  = max(-I_MAX, min(I_MAX, i_prev + e_t * DT_HOURS))
        self._integral[site_id] = i_new
        delta_Q_I = K_I * i_new                            # integral boost [kVAr]

        # ── Combined command  ΔQ_commanded = ΔQ_P + ΔQ_I ─────────
        delta_Q_cmd = delta_Q_P + delta_Q_I

        # ── Clamp to model hardware capacity ──────────────────────
        delta_Q_clamped = max(-model_kva, min(model_kva, delta_Q_cmd))

        log.debug(
            "[%s] e=%.4f  ΔQ_P=%+.2f  ΔQ_I=%+.2f  ΔQ_cmd=%+.2f → %+.2f kVAr  I=%.3f",
            site_id, e_t, delta_Q_P, delta_Q_I, delta_Q_cmd, delta_Q_clamped, i_new,
        )

        # ── Action selection ───────────────────────────────────────
        if abs(delta_Q_clamped) < 0.5:   # below hardware resolution
            return ACTION_HOLD, None

        if delta_Q_clamped > 0:
            return ACTION_INJECT, round(delta_Q_clamped, 2)
        else:
            return ACTION_REDUCE, round(abs(delta_Q_clamped), 2)

    # ── State simulator ────────────────────────────────────────────

    def _simulate_after(self, state: State, action: str,
                        action_kvar: float | None) -> State:
        """
        Estimates post-action state for reward computation (historical / backtest mode).
        In live deployment this is replaced by the actual next 15-min meter reading.

        Physics:
          Q_after = Q_before − ΔQ_injected

          Injection cancels reactive power:
            Q_before (lagging, positive) − ΔQ_injected (capacitive) → smaller Q_after

          S_after  = √(P² + Q_after²)
          PF_after = P / S_after

          INJECT   →  Q_after = Q_current − ΔQ      [reduces lagging kVAr]
          REDUCE   →  Q_after = Q_current × 0.5     [halves injection — approximate]
          HOLD     →  Q_after = Q_current            [no change]

        Note on sign convention:
          positive kVAr = lagging (inductive load) — most common
          negative kVAr = leading (capacitive load) — e.g., lightly loaded cables
          INJECT ΔQ > 0 → reduces lagging kVAr toward zero
          REDUCE       → used when leading PF or solar storage priority
        """
        if action == ACTION_INJECT and action_kvar is not None:
            # Capacitive injection reduces lagging reactive power
            # ΔQ_injected is the signed kVAr commanded
            kvar_after = state.kVAr - action_kvar
        elif action == ACTION_REDUCE and action_kvar is not None:
            kvar_after = state.kVAr - action_kvar     # inductive injection reduces leading
        elif action == ACTION_REDUCE:
            kvar_after = state.kVAr * 0.5             # fallback: halve reactive
        else:
            kvar_after = state.kVAr                   # HOLD — no change

        # S_after = √(P² + Q_after²)
        # PF_after = P / S_after
        kva_after = math.sqrt(state.kW ** 2 + kvar_after ** 2)
        pf_after  = abs(state.kW / kva_after) if kva_after > 0 else 1.0

        return State(
            site_id   = state.site_id,
            timestamp = state.timestamp,
            kW        = state.kW,
            kVAr      = kvar_after,
            PF        = pf_after,
            voltage_V = state.voltage_V,
            solar     = state.solar,
        )
