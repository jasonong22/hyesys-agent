"""
State builder — constructs 15-min snapshot structs from hyesys.db records.

Key power equations embedded in State properties:

  Apparent power:   S  = √(P² + Q²)                         [kVA]
  Power factor:     PF = P / S = cos(φ)                     [dimensionless]
  PF angle:         φ  = arccos(PF)                         [rad]
  Line current:     I  = S / (√3 × V_L)                    [A]  3-phase
  Per-phase loss:   P_loss ≈ I² × R  (R cancels in ratios)  [kW]
  Loss fraction:    f  = 1 − (I_after / I_before)²
                       = 1 − (S_after / S_before)²          [dimensionless]
"""

import math
from dataclasses import dataclass
from datetime import datetime
from core.schema import SITE_CONFIG, PF_TARGET

SQRT3 = math.sqrt(3)   # √3 ≈ 1.7321


@dataclass
class State:
    site_id:   str
    timestamp: str
    kW:        float    # Active power P      [kW]
    kVAr:      float    # Reactive power Q    [kVAr]  negative = leading (capacitive)
    PF:        float    # Measured power factor (signed: negative = leading)
    voltage_V: float    # Line-to-neutral RMS voltage V_LN  [V]
    solar:     bool = False

    # ── Apparent power ───────────────────────────────────────────
    @property
    def kVA(self) -> float:
        """
        S = √(P² + Q²)
        Apparent power accounts for both active and reactive components.
        """
        return math.sqrt(self.kW ** 2 + self.kVAr ** 2)

    # ── Power factor angle ───────────────────────────────────────
    @property
    def phi_rad(self) -> float:
        """
        φ = arccos(|PF|)
        PF angle in radians. φ = 0 → unity PF, φ = π/2 → purely reactive.
        """
        return math.acos(min(abs(self.PF), 1.0))

    @property
    def phi_deg(self) -> float:
        """φ in degrees — useful for reporting."""
        return math.degrees(self.phi_rad)

    # ── Q components ─────────────────────────────────────────────
    @property
    def tan_phi(self) -> float:
        """
        tan(φ) = Q / P
        Used in Q-correction formula: ΔQ = P × (tan φ_current − tan φ_target)
        """
        return math.tan(self.phi_rad)

    @property
    def q_target(self) -> float:
        """
        Q_target = P × tan(arccos(PF_target))
        The reactive power at the target PF. Injection drives Q toward Q_target.
        """
        phi_target = math.acos(PF_TARGET)
        return self.kW * math.tan(phi_target)

    @property
    def q_correction_needed(self) -> float:
        """
        ΔQ = Q_current − Q_target
             = P × (tan φ_current − tan φ_target)
        Positive → lagging load needs capacitive injection.
        Negative → leading load needs inductive injection.
        """
        return round(self.kVAr - self.q_target, 3)

    # ── Current ──────────────────────────────────────────────────
    @property
    def current_A(self) -> float:
        """
        I = S / (√3 × V_L)      [3-phase, line-to-line voltage]

        V_L (line-to-line) = V_LN × √3  where V_LN is line-to-neutral.
        So: I = S / (√3 × V_LN × √3) = S / (3 × V_LN)

        If voltage_V is line-to-line:  I = S_kVA × 1000 / (√3 × V_L)
        If voltage_V is line-to-neutral: I = S_kVA × 1000 / (3 × V_LN)

        Here we treat voltage_V as line-to-neutral (230 V typical Singapore LV).
        """
        if self.voltage_V <= 0:
            return 0.0
        # S [kVA] → [VA]; 3 × V_LN for 3-phase line-to-neutral convention
        return (self.kVA * 1000.0) / (3.0 * self.voltage_V)

    @property
    def current_pu(self) -> float:
        """
        Per-unit current relative to nominal (230 V line-to-neutral).
        Used for loss fraction computation where nominal base cancels.
        """
        return self.current_A / (1000.0 / (3.0 * 230.0)) if self.voltage_V > 0 else 0.0

    # ── Losses ───────────────────────────────────────────────────
    @property
    def loss_fraction_vs_unity_pf(self) -> float:
        """
        Additional I²R losses caused by non-unity PF, as a fraction of total I²R.

        At unity PF:  I_unity = P / (√3 × V_L)  →  kVA = kW  →  S_unity = P
        Actual:       I_actual corresponds to S = √(P² + Q²)

        Excess loss fraction = 1 − (S_unity / S_actual)²
                             = 1 − (P / S)²
                             = 1 − PF²

        This is the fraction of I²R losses that HyESys can eliminate by
        correcting PF to unity. PF_target = 0.98 recovers ≈ 1 − 0.98² = 3.96%.
        """
        pf_abs = abs(self.PF)
        if pf_abs <= 0 or self.kVA == 0:
            return 0.0
        return round(1.0 - pf_abs ** 2, 4)

    @property
    def recoverable_loss_fraction(self) -> float:
        """
        Fraction of I²R losses recoverable by correcting to PF_TARGET = 0.98.

        f_recoverable = (I_current² − I_target²) / I_current²
                      = 1 − (S_target / S_current)²
                      = 1 − (P / (P / PF_target))² / (S_current / P)²

        Simplified:
          S_target = P / PF_target
          f = 1 − (S_target / S_current)²
        """
        if self.kW <= 0 or self.kVA <= 0:
            return 0.0
        s_target  = self.kW / PF_TARGET   # kVA at target PF
        fraction  = 1.0 - (s_target / self.kVA) ** 2
        return round(max(fraction, 0.0), 4)

    # ── Derived state flags ───────────────────────────────────────
    @property
    def is_leading_pf(self) -> bool:
        """Leading (capacitive) PF — Q < 0. As harmful as lagging."""
        return self.kVAr < 0

    @property
    def pf_abs(self) -> float:
        """Absolute power factor — removes leading/lagging sign."""
        return abs(self.PF)

    @property
    def hour(self) -> int:
        try:
            return datetime.fromisoformat(self.timestamp).hour
        except ValueError:
            return 0

    def __repr__(self):
        return (
            f"State({self.site_id} @ {self.timestamp} | "
            f"P={self.kW:.1f}kW Q={self.kVAr:.1f}kVAr "
            f"S={self.kVA:.1f}kVA PF={self.PF:.3f} φ={self.phi_deg:.1f}°)"
        )


def build_state(row) -> State:
    d       = dict(row) if hasattr(row, "keys") else row
    site_id = d.get("site_id", "UNKNOWN")
    solar   = SITE_CONFIG.get(site_id, {}).get("solar", True)
    return State(
        site_id   = site_id,
        timestamp = d.get("timestamp", ""),
        kW        = float(d.get("kW",        0) or 0),
        kVAr      = float(d.get("kVAr",      0) or 0),
        PF        = float(d.get("PF",        0) or 0),
        voltage_V = float(d.get("voltage_V", 0) or 0),
        solar     = solar,
    )


def build_states_from_rows(rows) -> list[State]:
    return [build_state(r) for r in rows]
