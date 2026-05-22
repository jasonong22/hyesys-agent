"""
Baoyuan Site — Agent 2: Analysis & Recommendation Engine.

Reads CLEAN CapBank1 and CapBank2 meter records from baoyuan.db,
runs the same PI controller and SAR loop as the base Agent 2,
and writes STATE → ACTION → REWARD triplets back to baoyuan.db.

Run: python sites/baoyuan/agent2.py
Customisation of Baoyuan-specific goals happens in this file progressively.
"""

import logging
import math
import sqlite3
import sys
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

# Add root to path so we can import from core/ and agent2/
ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

from agent2.outcome import compute_reward
from agent2.tools import compute_pf_correction, assess_demand_risk

DB_PATH = Path(__file__).parent / "data" / "baoyuan.db"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baoyuan.agent2")

# ─────────────────────────────────────────────
# BAOYUAN SITE CONFIG
# ─────────────────────────────────────────────
SITE_CONFIG = {
    "BAOYUAN-CAPBANK1": {"solar": False, "recommended_model": "H125"},
    "BAOYUAN-CAPBANK2": {"solar": False, "recommended_model": "H125"},
}

HYESYS_MODELS = {
    "H125": {"kVA": 125},
}

# ─────────────────────────────────────────────
# PI CONTROLLER PARAMETERS (same as base Agent 2)
# ─────────────────────────────────────────────
PF_TARGET  = 0.98
K_P        = 1.00
K_I        = 0.50
I_MAX      = 20.0
DEADBAND   = 0.005
DT_HOURS   = 0.25

ACTION_INJECT = "INJECT"
ACTION_HOLD   = "HOLD"
ACTION_REDUCE = "REDUCE"

HISTORY_WINDOW = 16


# ─────────────────────────────────────────────
# STATE
# ─────────────────────────────────────────────
@dataclass
class State:
    site_id:   str
    timestamp: str
    kW:        float
    kVAr:      float
    PF:        float
    voltage_V: float
    solar:     bool = False

    @property
    def kVA(self) -> float:
        return math.sqrt(self.kW ** 2 + self.kVAr ** 2)


def build_state(record: dict) -> State:
    site_id = record["site_id"]
    cfg     = SITE_CONFIG.get(site_id, {})
    return State(
        site_id   = site_id,
        timestamp = record["timestamp"],
        kW        = float(record.get("kW", 0) or 0),
        kVAr      = float(record.get("kVAr", 0) or 0),
        PF        = float(record.get("PF", 0) or 0),
        voltage_V = float(record.get("voltage_V", 0) or 0),
        solar     = cfg.get("solar", False),
    )


# ─────────────────────────────────────────────
# BAOYUAN AGENT 2
# ─────────────────────────────────────────────
class BaoyuanAgent2:
    """
    Agent 2 for the Baoyuan site.
    Implements the same PI controller and SAR loop as the base Agent 2.
    Baoyuan-specific customisation (goals, reward weights, decision logic)
    is added here progressively in later iterations.
    """

    def __init__(self, conn: sqlite3.Connection):
        self.conn      = conn
        self._history: dict[str, deque]  = defaultdict(lambda: deque(maxlen=HISTORY_WINDOW))
        self._peak_kw: dict[str, list]   = defaultdict(list)
        self._integral: dict[str, float] = defaultdict(float)

    def process(self, state: State) -> dict:
        site_id  = state.site_id
        history  = list(self._history[site_id])
        site_cfg = SITE_CONFIG.get(site_id, {})
        model_key = site_cfg.get("recommended_model", "H125")

        # PF correction tool (analytical ΔQ solution)
        pf_tool = compute_pf_correction(state, model=model_key)

        # Demand risk
        self._peak_kw[site_id].append(state.kW)
        demand = assess_demand_risk(state, self._peak_kw[site_id])

        # Decision
        action, action_kvar = self._decide(state, pf_tool, demand, site_cfg, site_id)

        log.info(
            "[%s] %s → %s  kVAr=%+.1f | PF=%.3f → %.3f | demand=%s",
            site_id, state.timestamp, action, action_kvar or 0.0,
            abs(state.PF), pf_tool["PF_achievable"], demand["risk_level"],
        )

        # Simulate post-action state and compute reward
        state_after = self._simulate_after(state, action, action_kvar)
        reward      = compute_reward(state, state_after)

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
        self._write_sar(sar)
        self._history[site_id].append(state)
        return sar

    def process_batch(self, states: list[State]) -> list[dict]:
        return [self.process(s) for s in states]

    # ── PI Controller ──────────────────────────────────────────────

    def _decide(self, state: State, pf_tool: dict, demand: dict,
                site_cfg: dict, site_id: str) -> tuple[str, float | None]:
        has_solar = site_cfg.get("solar", False)
        model_key = site_cfg.get("recommended_model", "H125")
        model_kva = HYESYS_MODELS.get(model_key, {}).get("kVA", 125)

        # Priority 1: solar storage / demand shaving
        if has_solar and demand["risk_level"] == "CRITICAL" and demand["recommend_store"]:
            self._integral[site_id] = 0.0
            return ACTION_REDUCE, None

        # Error signal
        pf_abs = abs(state.PF)
        e_t    = PF_TARGET - pf_abs

        # Dead-band
        if abs(e_t) < DEADBAND:
            self._integral[site_id] = 0.0
            return ACTION_HOLD, None

        # Proportional term
        delta_Q_P = K_P * pf_tool["delta_Q_required"]

        # Integral term with anti-windup
        i_new = max(-I_MAX, min(I_MAX, self._integral[site_id] + e_t * DT_HOURS))
        self._integral[site_id] = i_new
        delta_Q_I = K_I * i_new

        # Combined command, clamped to hardware capacity
        delta_Q_cmd     = delta_Q_P + delta_Q_I
        delta_Q_clamped = max(-model_kva, min(model_kva, delta_Q_cmd))

        if abs(delta_Q_clamped) < 0.5:
            return ACTION_HOLD, None

        if delta_Q_clamped > 0:
            return ACTION_INJECT, round(delta_Q_clamped, 2)
        return ACTION_REDUCE, round(abs(delta_Q_clamped), 2)

    # ── State Simulator ────────────────────────────────────────────

    def _simulate_after(self, state: State, action: str,
                        action_kvar: float | None) -> State:
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

        return State(
            site_id   = state.site_id,
            timestamp = state.timestamp,
            kW        = state.kW,
            kVAr      = kvar_after,
            PF        = pf_after,
            voltage_V = state.voltage_V,
            solar     = state.solar,
        )

    # ── SAR Writer ─────────────────────────────────────────────────

    def _write_sar(self, sar: dict) -> None:
        now = datetime.now(timezone.utc).isoformat()
        try:
            self.conn.execute(
                """
                INSERT INTO sar_log
                  (site_id, timestamp, state_kW, state_kVAr, state_PF, state_voltage_V,
                   action, action_kVAr, reward_pf_delta, reward_fraction, outcome, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sar["site_id"], sar["timestamp"],
                    sar["state_kW"], sar["state_kVAr"], sar["state_PF"], sar["state_voltage_V"],
                    sar["action"], sar.get("action_kVAr"),
                    sar.get("reward_pf_delta"), sar.get("reward_fraction"),
                    sar.get("outcome"), now,
                ),
            )
            self.conn.commit()
        except sqlite3.Error as e:
            log.error("SAR write failed: %s", e)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def get_clean_records(conn: sqlite3.Connection, site_id: str) -> list[dict]:
    rows = conn.execute(
        "SELECT * FROM meter_records WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') ORDER BY timestamp",
        (site_id,),
    ).fetchall()
    return [dict(zip([d[0] for d in conn.execute("SELECT * FROM meter_records LIMIT 0").description or
                      [(c[1],) for c in conn.execute("PRAGMA table_info(meter_records)").fetchall()]], row))
            for row in rows]


def get_clean_records_v2(conn: sqlite3.Connection, site_id: str) -> list[dict]:
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        "SELECT * FROM meter_records WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') ORDER BY timestamp",
        (site_id,),
    ).fetchall()
    return [dict(r) for r in rows]


def main():
    if not DB_PATH.exists():
        log.error("baoyuan.db not found at %s — run agent1.py first to ingest data.", DB_PATH)
        return

    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row

    agent = BaoyuanAgent2(conn)
    total_sar = 0

    for site_id in SITE_CONFIG:
        records = conn.execute(
            "SELECT * FROM meter_records WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') ORDER BY timestamp",
            (site_id,),
        ).fetchall()

        if not records:
            log.info("[%s] No CLEAN records found — skipping.", site_id)
            continue

        states = [build_state(dict(r)) for r in records]
        results = agent.process_batch(states)
        total_sar += len(results)

        outcomes = {"POSITIVE": 0, "NEUTRAL": 0, "NEGATIVE": 0}
        for r in results:
            outcomes[r["outcome"]] = outcomes.get(r["outcome"], 0) + 1

        log.info(
            "[%s] %d states processed → POSITIVE=%d NEUTRAL=%d NEGATIVE=%d",
            site_id, len(results),
            outcomes["POSITIVE"], outcomes["NEUTRAL"], outcomes["NEGATIVE"],
        )

    log.info("Agent 2 complete. Total SAR records written: %d", total_sar)
    conn.close()


if __name__ == "__main__":
    main()
