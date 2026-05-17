"""
Savings computation — current-reduction fraction method.
Primary validation observable: I_rms at the MSB incomer.
Cable R cancels in the fraction — no impedance modelling required.
"""

import math
from core.schema import PF_TARGET, THD_ASSUMPTION


def compute_savings_fraction(kvar_before: float, kvar_after: float, kw: float) -> float:
    """
    Computes the fraction of distribution losses eliminated.

    fraction = 1 - (I_after / I_before)^2
             = 1 - (kVA_after / kVA_before)^2   [at constant voltage]

    kVA = sqrt(kW^2 + kVAr^2)

    Returns a value in [0, 1]. Negative values (losses increased) clamped to 0.
    """
    kva_before = math.sqrt(kw ** 2 + kvar_before ** 2)
    kva_after  = math.sqrt(kw ** 2 + kvar_after  ** 2)

    if kva_before <= 0:
        return 0.0

    fraction = 1.0 - (kva_after / kva_before) ** 2
    return round(max(fraction, 0.0), 4)


def estimate_kwh_savings(fraction: float, avg_kw: float, hours: float) -> float:
    """
    Estimates kWh savings over a period.

    kW_meter = kW_loads + I^2 * R_losses
    Fraction of losses eliminated × estimated loss component × hours.

    Uses THD_ASSUMPTION to estimate loss component when not directly measured.
    """
    loss_component_kw = avg_kw * THD_ASSUMPTION   # approximate I^2R at 15% THD
    return round(fraction * loss_component_kw * hours, 3)


def back_calculate_thd(pf_before: float, pf_target: float = PF_TARGET,
                       fraction: float = 0.0) -> float | None:
    """
    Back-calculates site THD from measured PF and savings fraction.

    THD = sqrt((PF_before / PF_target)^2 / (1 - fraction) - 1)

    Returns None if inputs are invalid.
    """
    try:
        if fraction >= 1.0 or pf_before <= 0 or pf_target <= 0:
            return None
        inner = (pf_before / pf_target) ** 2 / (1.0 - fraction) - 1.0
        if inner < 0:
            return None
        return round(math.sqrt(inner), 4)
    except (ValueError, ZeroDivisionError):
        return None


def compute_kvar_required(kw: float, pf_current: float,
                           pf_target: float = PF_TARGET) -> float:
    """
    Calculates kVAr correction needed to move from pf_current to pf_target.
    Returns signed kVAr (positive = lagging correction needed).
    """
    if kw == 0 or pf_current <= 0:
        return 0.0
    phi_current = math.acos(min(abs(pf_current), 1.0))
    phi_target  = math.acos(min(pf_target, 1.0))
    return round(kw * (math.tan(phi_current) - math.tan(phi_target)), 2)


def summarise_savings(sar_records: list[dict]) -> dict:
    """
    Computes aggregate savings metrics from a list of SAR records.
    Returns a summary dict suitable for reporting.
    """
    if not sar_records:
        return {}

    fractions  = [r["reward_fraction"] for r in sar_records if r.get("reward_fraction") is not None]
    pf_deltas  = [r["reward_pf_delta"]  for r in sar_records if r.get("reward_pf_delta")  is not None]
    outcomes   = [r["outcome"]           for r in sar_records if r.get("outcome")]

    n = len(sar_records)
    return {
        "total_decisions":        n,
        "avg_loss_fraction":      round(sum(fractions) / len(fractions), 4) if fractions else None,
        "avg_pf_delta":           round(sum(pf_deltas)  / len(pf_deltas),  4) if pf_deltas  else None,
        "positive_outcomes":      outcomes.count("POSITIVE"),
        "neutral_outcomes":       outcomes.count("NEUTRAL"),
        "negative_outcomes":      outcomes.count("NEGATIVE"),
        "positive_outcome_pct":   round(outcomes.count("POSITIVE") / n * 100, 1) if n else 0,
    }
