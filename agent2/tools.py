"""
Agent 2 decision tools — explicit equations at every step.

PF CORRECTION EQUATIONS
───────────────────────
  S       = √(P² + Q²)                          apparent power [kVA]
  PF      = P / S = cos(φ)                       power factor
  φ       = arccos(PF)                           PF angle [rad]
  Q_target = P × tan(arccos(PF_target))          desired reactive power
  ΔQ      = Q_current − Q_target                 injection needed [kVAr]
           = P × (tan φ_current − tan φ_target)
  S_after = √(P² + (Q − ΔQ_injected)²)
  PF_after = P / S_after

CURRENT REDUCTION EQUATIONS
────────────────────────────
  I = S / (√3 × V_L)                            3-phase line current [A]
  Since V_L is approximately constant:
  I_after / I_before = S_after / S_before
  Savings fraction:
  f = 1 − (I_after / I_before)²
    = 1 − (S_after / S_before)²
    = 1 − (P² + Q_after²) / (P² + Q_before²)

LOSS REDUCTION EQUATIONS
─────────────────────────
  P_loss = 3 × I² × R_phase                     [kW] (3-phase cable)
  ΔP_loss = P_loss_before − P_loss_after
           = 3R × (I_before² − I_after²)
           = 3R × I_before² × f                 R cancels in fraction
  kW_saved = f × P_loss_base

DEMAND RISK EQUATIONS
──────────────────────
  demand_pct = kW_current / kW_peak_15min × 100  [%]
  Demand charge ∝ max(15-min kW) — Singapore tariff
"""

import math
import logging
from agent2.state import State
from core.schema import PF_TARGET, PF_PENALTY_THRESHOLD, HYESYS_MODELS, SITE_CONFIG, THD_ASSUMPTION

log = logging.getLogger("hyesys.agent2.tools")

SQRT3 = math.sqrt(3)
PHI_TARGET = math.acos(PF_TARGET)      # arccos(0.98) ≈ 0.1997 rad ≈ 11.48°
TAN_PHI_TARGET = math.tan(PHI_TARGET)  # tan(arccos(0.98)) ≈ 0.2031


# ── PF CORRECTION ──────────────────────────────────────────────────

def compute_pf_correction(state: State, model: str = "H50") -> dict:
    """
    Calculates kVAr injection to drive PF → PF_TARGET = 0.98.

    Step 1: Q_target = P × tan(arccos(PF_target))
              = P × tan(arccos(0.98))
              = P × 0.2031

    Step 2: ΔQ = Q_current − Q_target
              = P × (tan φ_current − tan φ_target)

    Step 3: Clamp ΔQ to model capacity  [−kVA_model, +kVA_model]

    Step 4: Q_after = Q_current − ΔQ_injected
            S_after = √(P² + Q_after²)
            PF_after = P / S_after

    Step 5: Savings fraction
            f = 1 − (S_after / S_before)²
    """
    P    = state.kW
    Q    = state.kVAr
    S    = state.kVA
    PF   = state.pf_abs

    model_kva  = HYESYS_MODELS.get(model, {}).get("kVA", 50)

    if P == 0 or S == 0:
        return _zero_correction(model_kva)

    # Step 1 — target reactive power
    Q_target = P * TAN_PHI_TARGET                          # [kVAr]

    # Step 2 — required injection (signed)
    delta_Q = Q - Q_target                                 # [kVAr]
    # (+) = lagging load, inject capacitive kVAr
    # (−) = leading load, inject inductive kVAr

    # Step 3 — clamp to model hardware capacity
    delta_Q_clamped = max(-model_kva, min(model_kva, delta_Q))
    at_capacity = abs(delta_Q) > model_kva

    # Step 4 — predicted post-injection state
    Q_after   = Q - delta_Q_clamped                       # [kVAr]
    S_after   = math.sqrt(P ** 2 + Q_after ** 2)          # [kVA]
    PF_after  = (P / S_after) if S_after > 0 else 1.0    # [−]

    # Step 5 — savings fraction
    # f = 1 − (S_after / S_before)²
    # = 1 − (I_after / I_before)²   [since I ∝ S at constant V]
    fraction = 1.0 - (S_after / S) ** 2 if S > 0 else 0.0

    # Estimated loss reduction [kW] using THD assumption for base losses
    P_loss_base  = P * THD_ASSUMPTION              # approximate I²R at nominal THD
    kW_saved_est = fraction * P_loss_base          # [kW]

    result = {
        # Inputs
        "P_kW":             round(P,              2),
        "Q_kVAr_before":    round(Q,              2),
        "S_kVA_before":     round(S,              2),
        "PF_before":        round(PF,             4),
        "phi_deg_before":   round(math.degrees(math.acos(min(PF, 1.0))), 2),

        # Correction
        "Q_target_kVAr":    round(Q_target,       2),
        "delta_Q_required": round(delta_Q,        2),
        "delta_Q_injected": round(delta_Q_clamped,2),
        "at_capacity":      at_capacity,

        # Post-injection prediction
        "Q_kVAr_after":     round(Q_after,        2),
        "S_kVA_after":      round(S_after,        2),
        "PF_achievable":    round(PF_after,        4),
        "phi_deg_after":    round(math.degrees(math.acos(min(PF_after, 1.0))), 2),

        # Savings
        "savings_fraction": round(max(fraction, 0.0), 4),
        "kW_saved_est":     round(max(kW_saved_est, 0.0), 3),

        # Model
        "model":            model,
        "model_kVA":        model_kva,
    }

    log.debug(
        "PF correction [%s]: Q=%+.1f → ΔQ=%+.1f → PF %.3f→%.3f  f=%.4f",
        model, Q, delta_Q_clamped, PF, PF_after, fraction,
    )
    return result


def _zero_correction(model_kva):
    return {
        "P_kW": 0, "Q_kVAr_before": 0, "S_kVA_before": 0,
        "PF_before": 0, "phi_deg_before": 90,
        "Q_target_kVAr": 0, "delta_Q_required": 0, "delta_Q_injected": 0,
        "at_capacity": False,
        "Q_kVAr_after": 0, "S_kVA_after": 0, "PF_achievable": 1.0, "phi_deg_after": 0,
        "savings_fraction": 0.0, "kW_saved_est": 0.0,
        "model": "—", "model_kVA": model_kva,
    }


# ── DEMAND RISK ────────────────────────────────────────────────────

def assess_demand_risk(state: State, peak_kw_history: list[float]) -> dict:
    """
    Evaluates peak demand risk and battery storage recommendation.

    Singapore tariff structure:
      Demand charge ∝ maximum 15-min kW demand in the billing period.

    Risk percentage:
      demand_pct = (P_current / P_peak_historical) × 100  [%]

    Storage recommendation trigger:
      Charge battery during off-peak (store solar or grid);
      discharge during peak to shave demand charge.

    Risk levels:
      demand_pct ≥ 95% → CRITICAL  (near or at historical peak)
      demand_pct ≥ 85% → HIGH      (likely to set new peak)
      demand_pct ≥ 70% → MEDIUM    (elevated, monitor)
      demand_pct <  70% → LOW
    """
    P = state.kW

    if not peak_kw_history:
        return _unknown_demand(P)

    P_peak     = max(peak_kw_history)
    P_avg      = sum(peak_kw_history) / len(peak_kw_history)
    demand_pct = (P / P_peak * 100.0) if P_peak > 0 else 0.0

    # Headroom before new peak is set
    P_headroom = P_peak - P                                # [kW]

    # Estimated demand charge at risk (SGD) — rough proxy
    # Singapore LV tariff: demand charge ≈ SGD 9–12 / kW / month
    DEMAND_TARIFF_SGD_PER_KW = 10.0
    P_excess      = max(P - P_peak, 0.0)
    demand_charge_risk_sgd = P_excess * DEMAND_TARIFF_SGD_PER_KW

    if demand_pct >= 95:
        risk_level      = "CRITICAL"
        recommend_store = True
    elif demand_pct >= 85:
        risk_level      = "HIGH"
        recommend_store = True
    elif demand_pct >= 70:
        risk_level      = "MEDIUM"
        recommend_store = state.solar
    else:
        risk_level      = "LOW"
        recommend_store = False

    return {
        "current_kW":              round(P,                      2),
        "historical_peak_kW":      round(P_peak,                 2),
        "historical_avg_kW":       round(P_avg,                  2),
        "demand_pct":              round(demand_pct,             1),
        "headroom_kW":             round(P_headroom,             2),
        "risk_level":              risk_level,
        "recommend_store":         recommend_store,
        "demand_charge_risk_sgd":  round(demand_charge_risk_sgd, 2),
    }


def _unknown_demand(P):
    return {
        "current_kW": P, "historical_peak_kW": P, "historical_avg_kW": P,
        "demand_pct": 100.0, "headroom_kW": 0.0,
        "risk_level": "UNKNOWN", "recommend_store": False,
        "demand_charge_risk_sgd": 0.0,
    }


# ── LOSS REDUCTION ─────────────────────────────────────────────────

def compute_loss_reduction(state_before: State, state_after: State) -> dict:
    """
    Computes cable I²R loss reduction between two states.

    Physics:
      P_loss = 3 × I² × R_phase                  [kW, 3-phase]
      I = S / (√3 × V_L)                         [A]

    Since R and V_L are approximately constant between states:
      P_loss_before / P_loss_after = I_before² / I_after²
                                   = S_before² / S_after²

    Savings fraction (R cancels):
      f = 1 − (S_after / S_before)²
        = 1 − (P² + Q_after²) / (P² + Q_before²)

    THD back-calculation:
      PF_measured = DPF / √(1 + THD²)
      THD = √( (PF_before / PF_target)² / (1 − f) − 1 )
    """
    S_before = state_before.kVA
    S_after  = state_after.kVA

    # Savings fraction
    if S_before > 0:
        fraction = 1.0 - (S_after / S_before) ** 2
    else:
        fraction = 0.0
    fraction = max(fraction, 0.0)

    # Current reduction [A]
    I_before    = state_before.current_A
    I_after     = state_after.current_A
    delta_I     = I_before - I_after

    # Estimated kW saved
    # P_loss_base = P × THD  (approximate distribution loss at nominal conditions)
    P_loss_base = state_before.kW * THD_ASSUMPTION
    kW_saved    = fraction * P_loss_base

    # THD back-calculation
    # THD = √( (PF_before/PF_target)² / (1 − f) − 1 )
    thd_est = None
    pf_b    = state_before.pf_abs
    if pf_b > 0 and fraction < 1.0:
        inner = (pf_b / PF_TARGET) ** 2 / (1.0 - fraction) - 1.0
        if inner >= 0:
            thd_est = round(math.sqrt(inner) * 100, 1)   # as percentage

    return {
        "S_before_kVA":    round(S_before,  2),
        "S_after_kVA":     round(S_after,   2),
        "I_before_A":      round(I_before,  2),
        "I_after_A":       round(I_after,   2),
        "delta_I_A":       round(delta_I,   2),
        "savings_fraction": round(fraction, 4),
        "kW_saved_est":    round(kW_saved,  3),
        "THD_est_pct":     thd_est,
    }


# ── MODEL RECOMMENDATION ───────────────────────────────────────────

def recommend_model(avg_kvar: float, has_solar: bool) -> str:
    """
    Recommends HyESys model based on average reactive load.
    No-solar sites capped at H50 (SCDF and space constraints).

    Sizing rule: model kVA ≥ avg_kVAr × 1.2  (20% headroom)
    """
    required_kva = avg_kvar * 1.2

    if not has_solar:
        return "H30" if required_kva <= 30 else "H50"

    if required_kva <= 30:   return "H30"
    if required_kva <= 50:   return "H50"
    if required_kva <= 60:   return "H60"
    if required_kva <= 100:  return "H100"
    return "H125"
