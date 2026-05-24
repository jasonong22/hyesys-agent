"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 2 — BAOYUAN SITE                                             ║
║  Analysis & Decision Engine                                          ║
║  Copied from: templates/agent2_master.py  v1.0  (2026-05-24)        ║
╚══════════════════════════════════════════════════════════════════════╝

SITE: Baoyuan Industrial (诸暨市葆元实业有限公司)
      Site IDs: BAOYUAN-CAPBANK1, BAOYUAN-CAPBANK2

── SITE CHANGELOG (changes from master) ─────────────────────────────
2026-05-24  initial copy from master v1.0

STATE STRUCT EXTENSION (vs master):
  • State dataclass extended with Ia, Ib, Ic float fields
  • Added derived properties: I_avg, I_imbalance_pct
  • (Master State struct only holds kW, kVAr, PF, voltage_V, solar)

ACTION MODE OVERRIDE — switchable via agent2_config.json:
  • "current_imbalance_monitor" (DEFAULT) — no injection; monitors phase
    current imbalance only; logs POSITIVE/NEGATIVE outcomes
  • "pf_pi_control" — activates standard master PI controller logic
    (D1–D6) for reactive compensation when H125 is actively injecting

NEW DECISION RULE (not in master):
  • _decide_monitor(): ACTION_MONITOR with imbalance% as action_kVAr
    POSITIVE if imbalance ≤ alert_pct; NEGATIVE otherwise

PARAMETERS NOT IN MASTER:
  • action_mode in agent2_config.json (hot-switchable without restart)
  • BAOYUAN-specific SITE_CONFIG: solar=False, recommended_model=H125
  • HYESYS_MODELS defined locally (not imported from core.schema)
  • ACTION_MONITOR action type added (not in master schema)

REWARD OVERRIDE (monitor mode):
  • reward_pf_delta always 0.0 in monitor mode (no PF correction active)
  • reward_fraction = 1.0 − I_imbalance_pct (current balance metric)

SAR WRITER:
  • _write_sar() is inlined (master delegates to core.store.write_sar)
  • Uses baoyuan.db at sites/baoyuan/data/baoyuan.db

IMPORTS:
  • Imports compute_reward and tools from agent2/ master module
  • Does NOT fully re-implement reward logic (reuses master outcome.py)
──────────────────────────────────────────────────────────────────────

Run: python sites/baoyuan/agent2.py
"""

import json
import logging
import math
import sqlite3
import sys
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

from agent2.outcome import compute_reward
from agent2.tools import compute_pf_correction, assess_demand_risk

DB_PATH     = Path(__file__).parent / "data" / "baoyuan.db"
CONFIG_PATH = Path(__file__).parent / "agent2_config.json"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baoyuan.agent2")

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
def load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text())
    return {
        "action_mode": "current_imbalance_monitor",
        "pf_target": 0.98,
        "pi_control": {"k_p": 1.0, "k_i": 0.5, "i_max": 20.0, "deadband": 0.005, "dt_hours": 0.25},
        "current_imbalance_monitor": {"imbalance_alert_pct": 0.10, "log_observations": True},
    }

# ─────────────────────────────────────────────
# SITE CONFIG
# ─────────────────────────────────────────────
SITE_CONFIG = {
    "BAOYUAN-CAPBANK1": {"solar": False, "recommended_model": "H125"},
    "BAOYUAN-CAPBANK2": {"solar": False, "recommended_model": "H125"},
}

HYESYS_MODELS = {"H125": {"kVA": 125}}

HISTORY_WINDOW = 16

ACTION_INJECT  = "INJECT"
ACTION_HOLD    = "HOLD"
ACTION_REDUCE  = "REDUCE"
ACTION_MONITOR = "MONITOR"

# ─────────────────────────────────────────────
# STATE — includes current readings for Baoyuan
# ─────────────────────────────────────────────
@dataclass
class State:
    site_id:   str
    timestamp: str
    kW:        float
    kVAr:      float
    PF:        float
    voltage_V: float
    Ia:        float = 0.0
    Ib:        float = 0.0
    Ic:        float = 0.0
    solar:     bool  = False

    @property
    def kVA(self) -> float:
        return math.sqrt(self.kW ** 2 + self.kVAr ** 2)

    @property
    def I_avg(self) -> float:
        vals = [v for v in [self.Ia, self.Ib, self.Ic] if v > 0]
        return sum(vals) / len(vals) if vals else 0.0

    @property
    def I_imbalance_pct(self) -> float:
        currents = [self.Ia, self.Ib, self.Ic]
        i_max = max(currents)
        positives = [c for c in currents if c > 0]
        if not positives or i_max == 0:
            return 0.0
        return (i_max - min(positives)) / i_max


def build_state(record: dict) -> State:
    cfg = SITE_CONFIG.get(record["site_id"], {})
    return State(
        site_id   = record["site_id"],
        timestamp = record["timestamp"],
        kW        = float(record.get("kW", 0) or 0),
        kVAr      = float(record.get("kVAr", 0) or 0),
        PF        = float(record.get("PF", 0) or 0),
        voltage_V = float(record.get("voltage_V", 0) or 0),
        Ia        = float(record.get("Ia", 0) or 0),
        Ib        = float(record.get("Ib", 0) or 0),
        Ic        = float(record.get("Ic", 0) or 0),
        solar     = cfg.get("solar", False),
    )


# ─────────────────────────────────────────────
# BAOYUAN AGENT 2
# ─────────────────────────────────────────────
class BaoyuanAgent2:

    def __init__(self, conn: sqlite3.Connection):
        self.conn      = conn
        self.cfg       = load_config()
        self._history:  dict[str, deque]  = defaultdict(lambda: deque(maxlen=HISTORY_WINDOW))
        self._peak_kw:  dict[str, list]   = defaultdict(list)
        self._integral: dict[str, float]  = defaultdict(float)

    def process(self, state: State) -> dict:
        mode     = self.cfg.get("action_mode", "current_imbalance_monitor")
        site_id  = state.site_id
        site_cfg = SITE_CONFIG.get(site_id, {})

        if mode == "current_imbalance_monitor":
            action, action_kvar, outcome = self._decide_monitor(state)
        else:
            action, action_kvar, outcome = self._decide_pi(state, site_cfg, site_id)

        imbal_pct = round(state.I_imbalance_pct * 100, 1)
        log.info(
            "[%s] %s → %s | Ia=%.1f Ib=%.1f Ic=%.1f imbal=%.1f%% | %s",
            site_id, state.timestamp, action,
            state.Ia, state.Ib, state.Ic, imbal_pct, outcome,
        )

        sar = {
            "site_id":         site_id,
            "timestamp":       state.timestamp,
            "state_kW":        state.kW,
            "state_kVAr":      state.kVAr,
            "state_PF":        state.PF,
            "state_voltage_V": state.voltage_V,
            "action":          action,
            "action_kVAr":     action_kvar,
            "reward_pf_delta": 0.0,
            "reward_fraction": round(1.0 - state.I_imbalance_pct, 4),
            "outcome":         outcome,
        }
        self._write_sar(sar)
        self._history[site_id].append(state)
        return sar

    def process_batch(self, states: list[State]) -> list[dict]:
        return [self.process(s) for s in states]

    # ── current_imbalance_monitor mode ────────────────────────────────────
    def _decide_monitor(self, state: State) -> tuple[str, float | None, str]:
        alert_pct = self.cfg.get("current_imbalance_monitor", {}).get("imbalance_alert_pct", 0.10)
        imbal     = state.I_imbalance_pct

        if state.I_avg == 0:
            return ACTION_MONITOR, None, "NEUTRAL"

        outcome = "POSITIVE" if imbal <= alert_pct else "NEGATIVE"
        return ACTION_MONITOR, round(imbal * 100, 2), outcome

    # ── pf_pi_control mode ────────────────────────────────────────────────
    def _decide_pi(self, state: State, site_cfg: dict,
                   site_id: str) -> tuple[str, float | None, str]:
        pi   = self.cfg.get("pi_control", {})
        k_p  = pi.get("k_p", 1.0)
        k_i  = pi.get("k_i", 0.5)
        i_max_v = pi.get("i_max", 20.0)
        db   = pi.get("deadband", 0.005)
        dt   = pi.get("dt_hours", 0.25)
        pf_t = self.cfg.get("pf_target", 0.98)

        model_key = site_cfg.get("recommended_model", "H125")
        model_kva = HYESYS_MODELS.get(model_key, {}).get("kVA", 125)

        self._peak_kw[site_id].append(state.kW)
        pf_tool = compute_pf_correction(state, model=model_key)
        demand  = assess_demand_risk(state, self._peak_kw[site_id])

        if site_cfg.get("solar") and demand["risk_level"] == "CRITICAL" and demand["recommend_store"]:
            self._integral[site_id] = 0.0
            return ACTION_REDUCE, None, "NEUTRAL"

        pf_abs = abs(state.PF)
        e_t    = pf_t - pf_abs

        if abs(e_t) < db:
            self._integral[site_id] = 0.0
            return ACTION_HOLD, None, "POSITIVE"

        delta_Q_P = k_p * pf_tool["delta_Q_required"]
        i_new     = max(-i_max_v, min(i_max_v, self._integral[site_id] + e_t * dt))
        self._integral[site_id] = i_new
        delta_Q_I = k_i * i_new
        cmd       = max(-model_kva, min(model_kva, delta_Q_P + delta_Q_I))

        if abs(cmd) < 0.5:
            return ACTION_HOLD, None, "NEUTRAL"

        state_after = self._simulate_after(state, ACTION_INJECT if cmd > 0 else ACTION_REDUCE,
                                           abs(cmd))
        reward = compute_reward(state, state_after)
        action = ACTION_INJECT if cmd > 0 else ACTION_REDUCE
        return action, round(abs(cmd), 2), reward.outcome

    # ── State Simulator (PI mode only) ────────────────────────────────────
    def _simulate_after(self, state: State, action: str,
                        action_kvar: float | None) -> State:
        if action in (ACTION_INJECT, ACTION_REDUCE) and action_kvar is not None:
            kvar_after = state.kVAr - action_kvar
        elif action == ACTION_REDUCE:
            kvar_after = state.kVAr * 0.5
        else:
            kvar_after = state.kVAr

        kva_after = math.sqrt(state.kW ** 2 + kvar_after ** 2)
        pf_after  = abs(state.kW / kva_after) if kva_after > 0 else 1.0
        return State(
            site_id=state.site_id, timestamp=state.timestamp,
            kW=state.kW, kVAr=kvar_after, PF=pf_after,
            voltage_V=state.voltage_V, solar=state.solar,
        )

    # ── SAR Writer ────────────────────────────────────────────────────────
    def _write_sar(self, sar: dict) -> None:
        now = datetime.now(timezone.utc).isoformat()
        try:
            self.conn.execute(
                """
                INSERT OR IGNORE INTO sar_log
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
def main():
    if not DB_PATH.exists():
        log.error("baoyuan.db not found — run agent1.py first.")
        return

    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    agent = BaoyuanAgent2(conn)
    cfg   = load_config()

    log.info("Agent 2 running in mode: %s", cfg.get("action_mode"))

    total_sar = 0
    for site_id in SITE_CONFIG:
        records = conn.execute(
            "SELECT * FROM meter_records WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') ORDER BY timestamp",
            (site_id,),
        ).fetchall()

        if not records:
            log.info("[%s] No records — skipping.", site_id)
            continue

        states  = [build_state(dict(r)) for r in records]
        results = agent.process_batch(states)
        total_sar += len(results)

        counts = {"POSITIVE": 0, "NEUTRAL": 0, "NEGATIVE": 0}
        for r in results:
            counts[r.get("outcome", "NEUTRAL")] = counts.get(r.get("outcome", "NEUTRAL"), 0) + 1

        log.info(
            "[%s] %d records → POSITIVE=%d NEUTRAL=%d NEGATIVE=%d",
            site_id, len(results),
            counts["POSITIVE"], counts["NEUTRAL"], counts["NEGATIVE"],
        )

    log.info("Agent 2 complete. SAR records written: %d", total_sar)
    conn.close()


if __name__ == "__main__":
    main()
