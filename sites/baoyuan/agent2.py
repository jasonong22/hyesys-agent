"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 2 — BAOYUAN SITE                                             ║
║  Analysis, Decision & Experiment Controller                          ║
║  Copied from: templates/agent2_master.py  v1.0  (2026-05-24)        ║
╚══════════════════════════════════════════════════════════════════════╝

SITE: Baoyuan Industrial (诸暨市葆元实业有限公司)
      Site IDs: BAOYUAN-CAPBANK1, BAOYUAN-CAPBANK2, BAOYUAN-HYESYS

── SITE CHANGELOG (changes from master) ─────────────────────────────
2026-05-24  initial copy from master v1.0
2026-05-25  add kvar_sweep_experiment mode — automated kVAr sweep 0→120 kVAr
            add ExperimentController — state machine with BMS safety rules:
              • singleVoltageAvg > 3.45 V → kW forced to 0, kVAr continues
              • singleVoltageAvg < 3.2 V  → pause experiment, charge -10 kW/phase
              • resume when singleVoltageAvg >= 3.45 V
            add PCS temperature safety cap:
              • if pcs_v3 temp ≥ 83°C → snapshot (|kW|+|kVAr|) total as hard cap;
                no step may exceed that total going forward
            add MQTT publisher for command topic:
              stsc/aems/cabinet/26022703840003/multi/operate/tx
            runs as daemon (loop_forever) in experiment mode
            MQTT command payload confirmed 2026-05-25:
              set_reactive_power: {"cabinetId":..., "index":1, "key":"set_reactive_power",
                "params":{"reactivePowerA":X,"reactivePowerB":X,"reactivePowerC":X},"remote":true}
              set_active_power key assumed for charging (unverified — confirm before BMS test)
            EXPERIMENT STUDY DESIGN:
              Input  (1): HyESys H125 kVAr injection (Feeder 2)
              Outputs (3): CapBank1 current response
                           CapBank2 current response
                           Main Grid meter (provided by Baoyuan after experiment)

ACTION MODES:
  • "kvar_sweep_experiment" — automated sweep; runs as daemon
  • "current_imbalance_monitor" — passive monitoring; runs as batch
  • "pf_pi_control" — PI controller; runs as batch

EXPERIMENT CONFIG (agent2_config.json → "experiment" block):
  id, max_kVAr_total, step_increment_kVAr, step_duration_hours,
  bms_high_threshold_V, bms_low_threshold_V, charge_kW_per_phase,
  tick_interval_seconds
──────────────────────────────────────────────────────────────────────

Run: python sites/baoyuan/agent2.py
Stop: Ctrl+C  (experiment state is saved in DB — safe to resume)
"""

import json
import logging
import math
import sqlite3
import sys
import time
from collections import defaultdict, deque
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

import paho.mqtt.client as mqtt

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

UTC = timezone.utc


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
    "BAOYUAN-HYESYS":   {"solar": False, "recommended_model": "H125"},
}

HYESYS_MODELS  = {"H125": {"kVA": 125}}
HISTORY_WINDOW = 16

ACTION_INJECT  = "INJECT"
ACTION_HOLD    = "HOLD"
ACTION_REDUCE  = "REDUCE"
ACTION_MONITOR = "MONITOR"
ACTION_CHARGE  = "CHARGE"


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
        currents  = [self.Ia, self.Ib, self.Ic]
        i_max     = max(currents)
        positives = [c for c in currents if c > 0]
        if not positives or i_max == 0:
            return 0.0
        return (i_max - min(positives)) / i_max


def build_state(record: dict) -> State:
    cfg = SITE_CONFIG.get(record["site_id"], {})
    return State(
        site_id   = record["site_id"],
        timestamp = record["timestamp"],
        kW        = float(record.get("kW",        0) or 0),
        kVAr      = float(record.get("kVAr",      0) or 0),
        PF        = float(record.get("PF",        0) or 0),
        voltage_V = float(record.get("voltage_V", 0) or 0),
        Ia        = float(record.get("Ia",        0) or 0),
        Ib        = float(record.get("Ib",        0) or 0),
        Ic        = float(record.get("Ic",        0) or 0),
        solar     = cfg.get("solar", False),
    )


# ─────────────────────────────────────────════
# EXPERIMENT CONTROLLER
# ─────────────────────────────────────────════
class ExperimentController:
    """
    Runs the kVAr sweep experiment: 0 → 120 kVAr in +1.0 kVAr steps, 1 hour per step.
    State is persisted in experiment_log (DB) — safe to stop and resume.

    Study design:
      Input  (1): HyESys H125 kVAr injection at Feeder 2
      Outputs (3): CapBank1 phase currents, CapBank2 phase currents,
                   Main Grid meter (imported after experiment via import_maingrid.py)

    BMS safety rules:
      singleVoltageAvg > 3.45 V  →  kW must be 0; kVAr continues normally
      singleVoltageAvg < 3.20 V  →  pause step, charge −10 kW/phase until ≥ 3.45 V then resume
      3.20 ≤ V ≤ 3.45            →  normal operation (kW=0 as default)

    PCS temperature cap:
      If pcs_v3 temperature ≥ 83°C, snapshot |kW_total| + |kVAr_total| as a hard cap.
      No subsequent step may push (kW + kVAr) total above that cap.
    """

    STATUS_RUNNING          = "RUNNING"
    STATUS_COMPLETED        = "COMPLETED"
    STATUS_PAUSED_CHARGING  = "PAUSED_CHARGING"

    def __init__(self, conn: sqlite3.Connection, mqtt_client: mqtt.Client, cfg: dict):
        self.conn        = conn
        self.mqtt        = mqtt_client
        self.cfg         = cfg
        exp              = cfg.get("experiment", {})
        self.exp_id      = exp.get("id", "sweep_001")
        self.max_kvar    = float(exp.get("max_kVAr_total",      129.0))
        self.increment   = float(exp.get("step_increment_kVAr",   0.5))
        self.step_hours  = float(exp.get("step_duration_hours",    1.0))
        self.bms_high    = float(exp.get("bms_high_threshold_V",  3.45))
        self.bms_low     = float(exp.get("bms_low_threshold_V",   3.2))
        self.charge_kw   = float(exp.get("charge_kW_per_phase",  -10.0))
        mqtt_cfg         = cfg.get("mqtt_command", {})
        self.cmd_topic   = mqtt_cfg.get("topic",
                            "stsc/aems/cabinet/26022703840003/multi/operate/tx")
        self.device_id   = mqtt_cfg.get("device_id",  "26022703840003")
        self.cmd_index   = int(mqtt_cfg.get("index",  1))
        # PCS temperature cap — set once when temp first crosses 83°C
        self.temp_cap_threshold:   float       = 83.0
        self.temp_cap_total_kvar:  float | None = None

    # ── tick: called every tick_interval_seconds ──────────────────────
    def tick(self) -> None:
        bms_voltage  = self._latest_bms_voltage()
        hyesys_state = self._latest_hyesys_state()

        # ── PCS temperature cap: snapshot on first breach of 83°C ────
        if (hyesys_state is not None and
                hyesys_state["temp_C"] >= self.temp_cap_threshold and
                self.temp_cap_total_kvar is None):
            cap = round(abs(hyesys_state["kW"]) + abs(hyesys_state["kVAr"]), 2)
            self.temp_cap_total_kvar = cap
            log.warning(
                "[EXP] PCS temp %.1f°C ≥ %.0f°C — capping kW+kVAr total at %.2f "
                "(kW=%.2f  kVAr=%.2f)",
                hyesys_state["temp_C"], self.temp_cap_threshold, cap,
                abs(hyesys_state["kW"]), abs(hyesys_state["kVAr"]),
            )

        step = self._active_step()

        # ── All steps complete (or temp cap reached) ──────────────────
        if step is None:
            next_num  = self._next_step_number()
            next_kvar = round(next_num * self.increment, 3)
            if next_kvar > self.max_kvar:
                log.info("[EXP] Experiment %s COMPLETE — all %d steps done.",
                         self.exp_id, next_num - 1)
                self._issue_command(0.0, 0.0)
                return
            if self.temp_cap_total_kvar is not None and next_kvar > self.temp_cap_total_kvar:
                log.info(
                    "[EXP] Temp cap active (%.2f kVAr). Next step %.2f kVAr would exceed cap — "
                    "experiment stopping.",
                    self.temp_cap_total_kvar, next_kvar,
                )
                self._issue_command(0.0, 0.0)
                return
            step = self._create_step(next_num, next_kvar)
            log.info("[EXP] Starting step %d → %.1f kVAr total (%.2f/phase)",
                     next_num, next_kvar, next_kvar / 3)

        target_kvar_per_phase = step["target_kVAr_per_phase"]

        # ── BMS: low voltage → pause and charge ──────────────────────
        if bms_voltage is not None and bms_voltage < self.bms_low:
            if step["status"] != self.STATUS_PAUSED_CHARGING:
                self._set_step_status(step["id"], self.STATUS_PAUSED_CHARGING)
                self._increment_pause_count(step["id"])
                log.warning(
                    "[EXP] BMS voltage %.3f V < %.2f V — PAUSING step %d, charging at %.0f kW/phase",
                    bms_voltage, self.bms_low, step["step_number"], self.charge_kw,
                )
            self._issue_command(self.charge_kw, target_kvar_per_phase)
            return

        # ── BMS: was paused, now recovered ───────────────────────────
        if step["status"] == self.STATUS_PAUSED_CHARGING:
            if bms_voltage is None or bms_voltage >= self.bms_high:
                self._set_step_status(step["id"], self.STATUS_RUNNING)
                log.info("[EXP] BMS voltage %.3f V ≥ %.2f V — RESUMING step %d",
                         bms_voltage or 0, self.bms_high, step["step_number"])
            else:
                self._issue_command(self.charge_kw, target_kvar_per_phase)
                return

        # ── kW always 0 during injection (unless charging above) ─────
        kw_cmd = 0.0

        # ── Issue active injection command ────────────────────────────
        self._issue_command(kw_cmd, target_kvar_per_phase)

        # ── Check if step duration has elapsed ───────────────────────
        elapsed_h = self._elapsed_hours(step["step_started_at"])
        if elapsed_h >= self.step_hours:
            self._complete_step(step)
            log.info("[EXP] Step %d COMPLETE (%.2f h elapsed). Advancing.",
                     step["step_number"], elapsed_h)

    # ── MQTT command ──────────────────────────────────────────────────
    def _issue_command(self, kw_per_phase: float, kvar_per_phase: float) -> None:
        """
        Publish setpoints to HyESys via MQTT.
        set_reactive_power format confirmed from live broker capture 2026-05-25.
        set_active_power key is assumed for charging — verify before BMS test.
        """
        base = {"cabinetId": self.device_id, "index": self.cmd_index, "remote": True}

        kvar_payload = {
            **base,
            "key":    "set_reactive_power",
            "params": {
                "reactivePowerA": round(kvar_per_phase, 3),
                "reactivePowerB": round(kvar_per_phase, 3),
                "reactivePowerC": round(kvar_per_phase, 3),
            },
        }
        try:
            self.mqtt.publish(self.cmd_topic, json.dumps(kvar_payload), qos=1)
            log.debug("[CMD] kVAr → %.3f/phase", kvar_per_phase)
        except Exception as e:
            log.error("[CMD] kVAr publish failed: %s", e)

        if kw_per_phase != 0.0:
            # ⚠ "set_active_power" key assumed — verify against HyESys docs before use
            kw_payload = {
                **base,
                "key":    "set_active_power",
                "params": {
                    "activePowerA": round(kw_per_phase, 3),
                    "activePowerB": round(kw_per_phase, 3),
                    "activePowerC": round(kw_per_phase, 3),
                },
            }
            try:
                self.mqtt.publish(self.cmd_topic, json.dumps(kw_payload), qos=1)
                log.warning("[CMD] kW (CHARGING) %.1f/phase  ⚠ set_active_power key unverified",
                            kw_per_phase)
            except Exception as e:
                log.error("[CMD] kW publish failed: %s", e)

    # ── DB helpers ────────────────────────────────────────────────────
    def _latest_bms_voltage(self) -> float | None:
        row = self.conn.execute(
            "SELECT single_voltage_avg FROM bms_log ORDER BY timestamp DESC LIMIT 1"
        ).fetchone()
        return float(row[0]) if row and row[0] is not None else None

    def _latest_hyesys_state(self) -> dict | None:
        row = self.conn.execute(
            """SELECT kW, kVAr, temp_C FROM meter_records
               WHERE site_id='BAOYUAN-HYESYS' AND quality_tag IN ('CLEAN','SUSPECT')
               ORDER BY timestamp DESC LIMIT 1"""
        ).fetchone()
        if row and row[2] is not None:
            return {"kW": float(row[0] or 0), "kVAr": float(row[1] or 0), "temp_C": float(row[2])}
        return None

    def _active_step(self) -> dict | None:
        row = self.conn.execute(
            """SELECT id, step_number, target_kVAr_total, target_kVAr_per_phase,
                      step_started_at, status
               FROM experiment_log
               WHERE experiment_id=? AND status IN (?,?)
               ORDER BY step_number DESC LIMIT 1""",
            (self.exp_id, self.STATUS_RUNNING, self.STATUS_PAUSED_CHARGING),
        ).fetchone()
        if row:
            return {
                "id":                    row[0],
                "step_number":           row[1],
                "target_kVAr_total":     row[2],
                "target_kVAr_per_phase": row[3],
                "step_started_at":       row[4],
                "status":                row[5],
            }
        return None

    def _next_step_number(self) -> int:
        row = self.conn.execute(
            "SELECT MAX(step_number) FROM experiment_log WHERE experiment_id=?",
            (self.exp_id,),
        ).fetchone()
        return (row[0] or 0) + 1

    def _create_step(self, step_number: int, kvar_total: float) -> dict:
        now = datetime.now(UTC).isoformat()
        kvar_per_phase = round(kvar_total / 3, 4)
        self.conn.execute(
            """INSERT OR IGNORE INTO experiment_log
               (experiment_id, step_number, target_kVAr_total, target_kVAr_per_phase,
                step_started_at, status)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (self.exp_id, step_number, kvar_total, kvar_per_phase, now, self.STATUS_RUNNING),
        )
        self.conn.commit()
        return {
            "id":                    self.conn.execute("SELECT last_insert_rowid()").fetchone()[0],
            "step_number":           step_number,
            "target_kVAr_total":     kvar_total,
            "target_kVAr_per_phase": kvar_per_phase,
            "step_started_at":       now,
            "status":                self.STATUS_RUNNING,
        }

    def _complete_step(self, step: dict) -> None:
        now   = datetime.now(UTC).isoformat()
        start = step["step_started_at"]

        def avg_col(site, col):
            row = self.conn.execute(
                f"SELECT AVG({col}) FROM meter_records "
                f"WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') "
                f"AND timestamp BETWEEN ? AND ?",
                (site, start, now),
            ).fetchone()
            return round(row[0], 4) if row and row[0] is not None else None

        def max_col(site, col):
            row = self.conn.execute(
                f"SELECT MAX({col}) FROM meter_records "
                f"WHERE site_id=? AND quality_tag IN ('CLEAN','SUSPECT') "
                f"AND timestamp BETWEEN ? AND ?",
                (site, start, now),
            ).fetchone()
            return round(row[0], 4) if row and row[0] is not None else None

        self.conn.execute(
            """UPDATE experiment_log SET
               status='COMPLETED', step_completed_at=?,
               hyesys_kVAr_avg=?, hyesys_kW_avg=?,
               hyesys_Ia_avg=?, hyesys_Ib_avg=?, hyesys_Ic_avg=?, hyesys_PF_avg=?,
               hyesys_temp_max=?,
               capbank1_Ia_avg=?, capbank1_Ib_avg=?, capbank1_Ic_avg=?,
               capbank2_Ia_avg=?, capbank2_Ib_avg=?, capbank2_Ic_avg=?,
               bms_voltage_avg=?
               WHERE id=?""",
            (
                now,
                avg_col("BAOYUAN-HYESYS",   "kVAr"),
                avg_col("BAOYUAN-HYESYS",   "kW"),
                avg_col("BAOYUAN-HYESYS",   "Ia"),
                avg_col("BAOYUAN-HYESYS",   "Ib"),
                avg_col("BAOYUAN-HYESYS",   "Ic"),
                avg_col("BAOYUAN-HYESYS",   "PF"),
                max_col("BAOYUAN-HYESYS",   "temp_C"),
                avg_col("BAOYUAN-CAPBANK1", "Ia"),
                avg_col("BAOYUAN-CAPBANK1", "Ib"),
                avg_col("BAOYUAN-CAPBANK1", "Ic"),
                avg_col("BAOYUAN-CAPBANK2", "Ia"),
                avg_col("BAOYUAN-CAPBANK2", "Ib"),
                avg_col("BAOYUAN-CAPBANK2", "Ic"),
                self._avg_bms_voltage(start, now),
                step["id"],
            ),
        )
        self.conn.commit()

    def _avg_bms_voltage(self, start: str, end: str) -> float | None:
        row = self.conn.execute(
            "SELECT AVG(single_voltage_avg) FROM bms_log WHERE timestamp BETWEEN ? AND ?",
            (start, end),
        ).fetchone()
        return round(row[0], 4) if row and row[0] is not None else None

    def _set_step_status(self, row_id: int, status: str) -> None:
        self.conn.execute(
            "UPDATE experiment_log SET status=? WHERE id=?", (status, row_id)
        )
        self.conn.commit()

    def _increment_pause_count(self, row_id: int) -> None:
        self.conn.execute(
            "UPDATE experiment_log SET pause_count = pause_count + 1 WHERE id=?", (row_id,)
        )
        self.conn.commit()

    @staticmethod
    def _elapsed_hours(started_at: str) -> float:
        try:
            start = datetime.fromisoformat(started_at)
            if start.tzinfo is None:
                start = start.replace(tzinfo=UTC)
            return (datetime.now(UTC) - start).total_seconds() / 3600
        except Exception:
            return 0.0


# ─────────────────────────────────────────────
# BAOYUAN AGENT 2 (batch modes)
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

    def _decide_monitor(self, state: State) -> tuple[str, float | None, str]:
        alert_pct = self.cfg.get("current_imbalance_monitor", {}).get("imbalance_alert_pct", 0.10)
        imbal     = state.I_imbalance_pct

        if state.I_avg == 0:
            return ACTION_MONITOR, None, "NEUTRAL"

        outcome = "POSITIVE" if imbal <= alert_pct else "NEGATIVE"
        return ACTION_MONITOR, round(imbal * 100, 2), outcome

    def _decide_pi(self, state: State, site_cfg: dict,
                   site_id: str) -> tuple[str, float | None, str]:
        pi        = self.cfg.get("pi_control", {})
        k_p       = pi.get("k_p",       1.0)
        k_i       = pi.get("k_i",       0.5)
        i_max_v   = pi.get("i_max",    20.0)
        db        = pi.get("deadband", 0.005)
        dt        = pi.get("dt_hours", 0.25)
        pf_t      = self.cfg.get("pf_target", 0.98)
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

        state_after = self._simulate_after(state, ACTION_INJECT if cmd > 0 else ACTION_REDUCE, abs(cmd))
        reward      = compute_reward(state, state_after)
        action      = ACTION_INJECT if cmd > 0 else ACTION_REDUCE
        return action, round(abs(cmd), 2), reward.outcome

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

    def _write_sar(self, sar: dict) -> None:
        now = datetime.now(UTC).isoformat()
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
# MQTT PUBLISHER (experiment mode only)
# ─────────────────────────────────────────────
def _build_mqtt_publisher(cfg: dict) -> mqtt.Client:
    mc  = cfg.get("mqtt_command", {})
    pub = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    pub.username_pw_set(
        mc.get("username", "TYKJadmin"),
        mc.get("password", "TYKJ2018."),
    )

    def on_connect(client, userdata, flags, rc, properties):
        if rc == 0:
            log.info("[MQTT-PUB] Connected to broker for command publishing.")
        else:
            log.error("[MQTT-PUB] Connection failed: %s", rc)

    pub.on_connect = on_connect
    pub.connect(
        mc.get("broker_host", "loragw.advastech.com"),
        mc.get("broker_port", 1883),
        60,
    )
    pub.loop_start()
    return pub


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if not DB_PATH.exists():
        log.error("baoyuan.db not found — run agent1.py first.")
        return

    cfg  = load_config()
    mode = cfg.get("action_mode", "current_imbalance_monitor")
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    log.info("Agent 2 running in mode: %s", mode)

    # ── Experiment daemon mode ────────────────────────────────────────
    if mode == "kvar_sweep_experiment":
        exp      = cfg.get("experiment", {})
        tick_sec = int(exp.get("tick_interval_seconds", 60))
        pub      = _build_mqtt_publisher(cfg)
        ctrl     = ExperimentController(conn, pub, cfg)

        log.info("[EXP] Experiment controller starting — ID=%s  max=%.0f kVAr  step=%.1f kVAr  hold=%sh",
                 ctrl.exp_id, ctrl.max_kvar, ctrl.increment, ctrl.step_hours)
        log.info("[EXP] Total steps: %d  Estimated duration: %.1f hours",
                 int(ctrl.max_kvar / ctrl.increment),
                 int(ctrl.max_kvar / ctrl.increment) * ctrl.step_hours)

        try:
            while True:
                ctrl.tick()
                time.sleep(tick_sec)
        except KeyboardInterrupt:
            log.info("[EXP] Stopped. Experiment state saved in DB — safe to resume.")
        finally:
            ctrl._issue_command(0.0, 0.0)  # zero out on exit
            pub.loop_stop()
            pub.disconnect()
            conn.close()
        return

    # ── Batch modes (current_imbalance_monitor / pf_pi_control) ──────
    agent     = BaoyuanAgent2(conn)
    total_sar = 0

    for site_id in ["BAOYUAN-CAPBANK1", "BAOYUAN-CAPBANK2"]:
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
            o = r.get("outcome", "NEUTRAL")
            counts[o] = counts.get(o, 0) + 1

        log.info("[%s] %d records → POSITIVE=%d NEUTRAL=%d NEGATIVE=%d",
                 site_id, len(results), counts["POSITIVE"], counts["NEUTRAL"], counts["NEGATIVE"])

    log.info("Agent 2 complete. SAR records written: %d", total_sar)
    conn.close()


if __name__ == "__main__":
    main()
