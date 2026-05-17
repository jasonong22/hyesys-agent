"""
Reward computation — closes the SAR loop.

REWARD EQUATIONS
────────────────
PF improvement:
  r_PF = PF_after − PF_before                           [dimensionless, −1 … +1]

Current-reduction fraction (R cancels — no cable modelling needed):
  f = 1 − (I_after / I_before)²
    = 1 − (S_after / S_before)²                         [dimensionless, 0 … 1]
    = 1 − (P² + Q_after²) / (P² + Q_before²)

  Since I = S / (√3 × V_L) and V_L ≈ constant:
    I_after / I_before = S_after / S_before

Weighted combined reward:
  r_total = W_PF × r_PF + W_LOSS × f

  W_PF   = 0.60  (PF correction dominates — directly affects SP penalty)
  W_LOSS = 0.40  (I²R loss reduction — quantifies kWh savings)

THD back-calculation:
  PF_measured = DPF / √(1 + THD²)

  Rearranging for THD:
    THD = √( (PF_before / PF_target)² / (1 − f) − 1 )

  Interpretation:
    THD = 0   → purely displacement PF — correction sufficient
    THD > 0.2 → harmonic filtering also required (HyESys active filter mode)

OUTCOME CLASSIFICATION
──────────────────────
  r_PF ≥ +0.01  → POSITIVE   (measurable improvement)
  r_PF ≤ −0.01  → NEGATIVE   (PF degraded — action was counterproductive)
  otherwise     → NEUTRAL    (below detection threshold)
"""

import math
import logging
from dataclasses import dataclass, field
from agent2.state import State
from core.schema import PF_TARGET, REWARD_POSITIVE_PF_DELTA, REWARD_NEGATIVE_PF_DELTA

log = logging.getLogger("hyesys.agent2.outcome")

# ── Reward weights ─────────────────────────────────────────────────
W_PF   = 0.60   # weight: PF improvement   (affects tariff penalty directly)
W_LOSS = 0.40   # weight: I²R loss fraction (quantifies energy savings)


# ── Reward dataclass ───────────────────────────────────────────────

@dataclass
class Reward:
    """
    SAR reward struct — records all computed metrics for the SAR log.

    r_pf        PF improvement:        PF_after − PF_before
    r_loss      Loss fraction:         1 − (S_after / S_before)²
    r_total     Weighted combined:     W_PF × r_pf + W_LOSS × r_loss
    thd_est     Back-calculated THD:   √((PF_before/PF_target)² / (1−f) − 1)
    outcome     Classification:        POSITIVE / NEUTRAL / NEGATIVE
    """
    r_pf:        float           # PF_after − PF_before
    r_loss:      float           # 1 − (S_after/S_before)²
    r_total:     float           # weighted combined reward
    thd_est:     float | None    # back-calculated site THD (None if indeterminate)
    outcome:     str             # POSITIVE / NEUTRAL / NEGATIVE

    # Legacy aliases for backward compatibility with SAR store
    @property
    def pf_delta(self) -> float:
        return self.r_pf

    @property
    def loss_fraction(self) -> float:
        return self.r_loss


# ── Reward computation ─────────────────────────────────────────────

def compute_reward(state_before: State, state_after: State) -> Reward:
    """
    Computes the multi-component reward from a before/after state pair.

    Step 1 — PF improvement
    ─────────────────────────
      r_PF = PF_after − PF_before
      > 0 → injection improved power factor
      < 0 → injection degraded power factor (overcorrected or leading)

    Step 2 — Current-reduction fraction
    ─────────────────────────────────────
      f = 1 − (S_after / S_before)²

      Derivation:
        P_loss = 3 × I² × R_phase                [kW, 3-phase Ohmic losses]
        I = S / (√3 × V_L)
        P_loss_ratio = I_after² / I_before²
                     = S_after² / S_before²      [V_L and R cancel]

      f is the fraction of I²R losses eliminated by the injection.
      f = 0.05 means 5% of cable losses removed per 15-min interval.

    Step 3 — Weighted combined reward
    ───────────────────────────────────
      r_total = W_PF × r_PF + W_LOSS × f
        W_PF   = 0.60  (penalises SP threshold violations more heavily)
        W_LOSS = 0.40  (quantifies kWh savings, used for ROI reporting)

    Step 4 — THD back-calculation
    ──────────────────────────────
      THD = √( (PF_before / PF_target)² / (1 − f) − 1 )

      Derived from:
        PF_measured = DPF / √(1 + THD²)
        → THD² = (DPF / PF_measured)² − 1
        → With DPF ≈ PF_target (after correction), PF_measured = PF_before

    Step 5 — Outcome classification
    ──────────────────────────────────
      r_PF ≥ REWARD_POSITIVE_PF_DELTA (+0.01) → POSITIVE
      r_PF ≤ REWARD_NEGATIVE_PF_DELTA (−0.01) → NEGATIVE
      otherwise                                → NEUTRAL
    """
    pf_before = abs(state_before.PF)
    pf_after  = abs(state_after.PF)

    # Step 1 — PF improvement
    r_pf = round(pf_after - pf_before, 4)

    # Step 2 — Current-reduction fraction
    # f = 1 − (S_after / S_before)²
    S_before = state_before.kVA
    S_after  = state_after.kVA

    if S_before > 0:
        r_loss = round(max(1.0 - (S_after / S_before) ** 2, 0.0), 4)
    else:
        r_loss = 0.0

    # Step 3 — Weighted combined reward
    r_total = round(W_PF * r_pf + W_LOSS * r_loss, 4)

    # Step 4 — THD back-calculation
    # THD = √( (PF_before / PF_target)² / (1 − f) − 1 )
    thd_est = back_calculate_thd(pf_before, pf_after=PF_TARGET, loss_fraction=r_loss)

    # Step 5 — Outcome classification
    if r_pf >= REWARD_POSITIVE_PF_DELTA:
        outcome = "POSITIVE"
    elif r_pf <= REWARD_NEGATIVE_PF_DELTA:
        outcome = "NEGATIVE"
    else:
        outcome = "NEUTRAL"

    reward = Reward(
        r_pf    = r_pf,
        r_loss  = r_loss,
        r_total = r_total,
        thd_est = thd_est,
        outcome = outcome,
    )
    log.debug(
        "Reward: r_PF=%+.4f  r_loss=%.4f  r_total=%+.4f  THD_est=%.1f%%  outcome=%s",
        r_pf, r_loss, r_total,
        thd_est * 100 if thd_est is not None else float("nan"),
        outcome,
    )
    return reward


# ── THD back-calculation ───────────────────────────────────────────

def back_calculate_thd(pf_before: float, pf_after: float = PF_TARGET,
                       loss_fraction: float = 0.0) -> float | None:
    """
    Back-calculates site Total Harmonic Distortion from PF and loss fraction.

    Derivation:
      PF_measured = DPF / √(1 + THD²)
      where DPF (displacement PF) ≈ PF_target after correction.

      Rearranging:
        √(1 + THD²) = DPF / PF_measured = PF_target / PF_before
        1 + THD²    = (PF_target / PF_before)²

      With loss fraction adjustment (current reduction shifts the PF ratio):
        THD² = (PF_before / PF_target)² / (1 − f) − 1

      Boundary cases:
        f → 0      : THD = √((PF_before/PF_target)² − 1)  [no loss reduction]
        f → 1      : indeterminate (denominator → 0), return None
        inner < 0  : measurement inconsistency, return None

    Returns:
      THD as a decimal fraction (e.g. 0.15 = 15 %) or None if indeterminate.
    """
    try:
        if loss_fraction >= 1.0 or pf_before <= 0 or pf_after <= 0:
            return None
        inner = (pf_before / pf_after) ** 2 / (1.0 - loss_fraction) - 1.0
        if inner < 0:
            return None
        return round(math.sqrt(inner), 4)
    except (ValueError, ZeroDivisionError):
        return None
