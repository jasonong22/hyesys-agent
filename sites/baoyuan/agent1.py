"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 1 — BAOYUAN SITE                                             ║
║  Data Quality & Ingestion — CapBank1, CapBank2, HyESys H125         ║
║  Copied from: templates/agent1_master.py  v1.0  (2026-05-24)        ║
╚══════════════════════════════════════════════════════════════════════╝

SITE: Baoyuan Industrial (诸暨市葆元实业有限公司)
      Monitors: CapBank1 (device 0086040215999997)
                CapBank2 (device 0086040215999996)
                HyESys H125 (device 26022703840003)
      HyESys H125 installed at Cabinet A / Feeder 2

── SITE CHANGELOG (changes from master) ─────────────────────────────
2026-05-24  initial copy from master v1.0
2026-05-25  add HyESys H125 MQTT subscription (stsc/aems/message/26022703840003)
            add bms_log, experiment_log, maingrid_history DB tables
            add parse_hyesys_message() — dispatches on 'type' field
              pcs_v3: full electrical readings (kW/kVAr/V/A/PF/temp/per-phase)
              bms: battery management (singleVoltageAvg is critical safety field)
            extend meter_records with temp_C, kVAr_A/B/C, kW_A/B/C, amp_imbalance_pct

DATA SOURCE CHANGES (vs master):
  • Live MQTT ingestion (not batch CSV) — paho-mqtt client added
  • CapBank topics: hyesys/data/dev/<device_id>  (current-only meters)
  • HyESys topic:  stsc/aems/message/26022703840003  (full electrical + BMS)

METER TYPE OVERRIDE — CapBank current-only instruments:
  • kW, kVAr, PF, voltage_V always 0 for CapBanks — normal, not a fault
  • Validation uses current-only rules for CapBank site_ids

PAYLOAD FORMAT NOTE (stsc/aems/message/<id>):
  All data is nested under raw["data"]; timestamp is reportTimeTs in milliseconds.
  Field names confirmed from live MQTT capture 2026-05-25.
──────────────────────────────────────────────────────────────────────

Run: python sites/baoyuan/agent1.py
Stop: Ctrl+C
"""

import json
import logging
import math
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path

import paho.mqtt.client as mqtt

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
_CONFIG_PATH = Path(__file__).parent / "agent1_config.json"

def _load_config() -> dict:
    if _CONFIG_PATH.exists():
        return json.loads(_CONFIG_PATH.read_text())
    return {"zero_current_threshold_A": 1.0, "imbalance_threshold_pct": 0.10}

AGENT1_CFG = _load_config()

BROKER_HOST   = "loragw.advastech.com"
BROKER_PORT   = 1883
KEEPALIVE_SEC = 60
MQTT_USERNAME = "TYKJadmin"
MQTT_PASSWORD = "TYKJ2018."

TOPICS = [
    ("hyesys/data/dev/0086040215999997",       0),  # CapBank1
    ("hyesys/data/dev/0086040215999996",       0),  # CapBank2
    ("stsc/aems/message/26022703840003",        0),  # HyESys H125
]

DEVICE_TO_SITE: dict[str, str] = {
    "0086040215999997": "BAOYUAN-CAPBANK1",
    "0086040215999996": "BAOYUAN-CAPBANK2",
    "26022703840003":   "BAOYUAN-HYESYS",
}

DB_PATH = Path(__file__).parent / "data" / "baoyuan.db"

# pcs_v3 confirmed field names (verified from live MQTT capture 2026-05-25)
# All data is nested under raw["data"]; timestamp is reportTimeTs in milliseconds.
# pcs_v3: voltageA/B/C, currentA/B/C, hz, activePowerTotal, reactivePowerTotal,
#         apparentPowerTotal, powerFactorTotal, reactivePowerA/B/C, activePowerA/B/C,
#         temperature, inputPower, inputVoltage, inputCurrent
# bms:    singleVoltageAvg, soc, soh, voltage, current, tempMain, singleVoltageMax/Min

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("baoyuan.agent1")


# ─────────────────────────────────────────────
# DATABASE
# ─────────────────────────────────────────────
def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")

    conn.execute("""
        CREATE TABLE IF NOT EXISTS meter_records (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id          TEXT    NOT NULL,
            timestamp        TEXT    NOT NULL,
            kW               REAL    NOT NULL,
            kVAr             REAL    NOT NULL,
            PF               REAL    NOT NULL,
            voltage_V        REAL    NOT NULL,
            kVA              REAL,
            Ia               REAL,
            Ib               REAL,
            Ic               REAL,
            frequency_Hz     REAL,
            temp_C           REAL,
            kVAr_A           REAL,
            kVAr_B           REAL,
            kVAr_C           REAL,
            kW_A             REAL,
            kW_B             REAL,
            kW_C             REAL,
            amp_imbalance_pct REAL,
            quality_tag      TEXT    NOT NULL,
            reject_reason    TEXT,
            ingested_at      TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_site_ts
        ON meter_records (site_id, timestamp)
    """)

    # Migrate existing DB: add new columns if they don't exist
    existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(meter_records)")}
    for col, coltype in [
        ("temp_C",            "REAL"),
        ("kVAr_A",            "REAL"),
        ("kVAr_B",            "REAL"),
        ("kVAr_C",            "REAL"),
        ("kW_A",              "REAL"),
        ("kW_B",              "REAL"),
        ("kW_C",              "REAL"),
        ("amp_imbalance_pct", "REAL"),
    ]:
        if col not in existing_cols:
            try:
                conn.execute(f"ALTER TABLE meter_records ADD COLUMN {col} {coltype}")
            except sqlite3.OperationalError:
                pass

    conn.execute("""
        CREATE TABLE IF NOT EXISTS sar_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id         TEXT NOT NULL,
            timestamp       TEXT NOT NULL,
            state_kW        REAL NOT NULL,
            state_kVAr      REAL NOT NULL,
            state_PF        REAL NOT NULL,
            state_voltage_V REAL NOT NULL,
            action          TEXT NOT NULL,
            action_kVAr     REAL,
            reward_pf_delta REAL,
            reward_fraction REAL,
            outcome         TEXT,
            created_at      TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_sar_site
        ON sar_log (site_id, timestamp)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS bms_log (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp           TEXT    NOT NULL,
            single_voltage_avg  REAL,
            raw_json            TEXT,
            ingested_at         TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_bms_ts ON bms_log (timestamp)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS experiment_log (
            id                    INTEGER PRIMARY KEY AUTOINCREMENT,
            experiment_id         TEXT    NOT NULL,
            step_number           INTEGER NOT NULL,
            target_kVAr_total     REAL    NOT NULL,
            target_kVAr_per_phase REAL    NOT NULL,
            step_started_at       TEXT    NOT NULL,
            step_completed_at     TEXT,
            status                TEXT    NOT NULL,
            hyesys_kVAr_avg       REAL,
            hyesys_kW_avg         REAL,
            hyesys_Ia_avg         REAL,
            hyesys_Ib_avg         REAL,
            hyesys_Ic_avg         REAL,
            hyesys_PF_avg         REAL,
            hyesys_temp_max       REAL,
            capbank1_Ia_avg       REAL,
            capbank1_Ib_avg       REAL,
            capbank1_Ic_avg       REAL,
            capbank2_Ia_avg       REAL,
            capbank2_Ib_avg       REAL,
            capbank2_Ic_avg       REAL,
            bms_voltage_avg       REAL,
            pause_count           INTEGER DEFAULT 0,
            notes                 TEXT
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_exp_step
        ON experiment_log (experiment_id, step_number)
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS maingrid_history (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp   TEXT    NOT NULL UNIQUE,
            kW          REAL,
            kVAr        REAL,
            PF          REAL,
            I_A         REAL,
            I_C         REAL,
            V_A         REAL,
            V_C         REAL,
            kWh         REAL,
            kW_A        REAL,
            kW_C        REAL,
            imported_at TEXT    NOT NULL
        )
    """)

    conn.commit()
    log.info("Database ready: %s", db_path)
    return conn


# ─────────────────────────────────────────────
# CAPBANK PAYLOAD PARSER
# ─────────────────────────────────────────────
def parse_capbank_payload(raw: dict, topic: str) -> dict | None:
    reported = raw.get("reported", {})
    if not reported:
        log.warning("REJECTED [%s] — missing 'reported' key", topic)
        return None

    data_key = next(iter(reported), None)
    if not data_key:
        return None

    m         = reported[data_key]
    device_id = data_key.replace("0_5_", "")
    site_id   = DEVICE_TO_SITE.get(device_id, device_id)

    ua = float(m.get("Ua", 0) or 0)
    ub = float(m.get("Ub", 0) or 0)
    uc = float(m.get("Uc", 0) or 0)
    phase_vs = [v for v in [ua, ub, uc] if v > 0]

    if phase_vs:
        voltage_V = sum(phase_vs) / len(phase_vs)
    else:
        uab = float(m.get("Uab", 0) or 0)
        ubc = float(m.get("Ubc", 0) or 0)
        uca = float(m.get("Uca", 0) or 0)
        line_vs = [v / math.sqrt(3) for v in [uab, ubc, uca] if v > 0]
        voltage_V = sum(line_vs) / len(line_vs) if line_vs else 0.0

    sendtime = raw.get("sendtime") or raw.get("timestamp", 0)
    try:
        ts = datetime.fromtimestamp(int(sendtime), tz=timezone.utc).isoformat()
    except (OSError, OverflowError, ValueError, TypeError):
        ts = datetime.now(timezone.utc).isoformat()

    return {
        "site_id":      site_id,
        "timestamp":    ts,
        "kW":           float(m.get("P",  0) or 0),
        "kVAr":         float(m.get("Q",  0) or 0),
        "PF":           float(m.get("PF", 0) or 0),
        "voltage_V":    round(voltage_V, 2),
        "kVA":          float(m.get("S",  0) or 0),
        "Ia":           float(m.get("Ia", 0) or 0),
        "Ib":           float(m.get("Ib", 0) or 0),
        "Ic":           float(m.get("Ic", 0) or 0),
        "frequency_Hz": float(m.get("Fr", 0) or 0),
    }


# ─────────────────────────────────────────────
# HYESYS PAYLOAD PARSER (stsc/aems/message/...)
# ─────────────────────────────────────────────
def parse_hyesys_message(raw: dict, topic: str) -> list[dict]:
    """
    Parse STSC AEMS messages from the HyESys H125 device.
    Each message has a 'type' field; we handle 'bms' and 'pcs_v3' only.
    Returns a list of parsed records (empty list if type is ignored).

    All data fields are nested under raw["data"].
    Timestamp is raw["reportTimeTs"] in milliseconds.
    Field names confirmed from live MQTT capture 2026-05-25.
    """
    ts_ms = raw.get("reportTimeTs", 0)
    try:
        ts = datetime.fromtimestamp(int(ts_ms) / 1000, tz=timezone.utc).isoformat()
    except (OSError, OverflowError, ValueError, TypeError):
        ts = datetime.now(timezone.utc).isoformat()

    msg_type = str(raw.get("type", "")).lower()
    d = raw.get("data", {})

    if msg_type == "bms":
        return [{"_record_type": "bms", "timestamp": ts, "raw": raw}]

    if msg_type == "pcs_v3":
        va = float(d.get("voltageA", 0) or 0)
        vb = float(d.get("voltageB", 0) or 0)
        vc = float(d.get("voltageC", 0) or 0)
        voltages  = [v for v in [va, vb, vc] if v > 0]
        voltage_V = sum(voltages) / len(voltages) if voltages else 0.0

        ia = float(d.get("currentA", 0) or 0)
        ib = float(d.get("currentB", 0) or 0)
        ic = float(d.get("currentC", 0) or 0)
        i_max = max(ia, ib, ic)
        i_min = min(c for c in [ia, ib, ic] if c > 0) if any(c > 0 for c in [ia, ib, ic]) else 0.0
        amp_imbalance = round((i_max - i_min) / i_max, 4) if i_max > 0 else 0.0

        rec = {
            "_record_type": "pcs_v3",
            "site_id":           "BAOYUAN-HYESYS",
            "timestamp":         ts,
            "kW":                float(d.get("activePowerTotal",   0) or 0),
            "kVAr":              float(d.get("reactivePowerTotal", 0) or 0),
            "PF":                float(d.get("powerFactorTotal",   0) or 0),
            "voltage_V":         round(voltage_V, 2),
            "kVA":               float(d.get("apparentPowerTotal", 0) or 0),
            "Ia":                ia,
            "Ib":                ib,
            "Ic":                ic,
            "frequency_Hz":      float(d.get("hz",          0) or 0),
            "temp_C":            float(d.get("temperature", 0) or 0),
            "kVAr_A":            float(d.get("reactivePowerA", 0) or 0),
            "kVAr_B":            float(d.get("reactivePowerB", 0) or 0),
            "kVAr_C":            float(d.get("reactivePowerC", 0) or 0),
            "kW_A":              float(d.get("activePowerA",  0) or 0),
            "kW_B":              float(d.get("activePowerB",  0) or 0),
            "kW_C":              float(d.get("activePowerC",  0) or 0),
            "amp_imbalance_pct": amp_imbalance,
        }
        return [rec]

    return []  # ignore all other types


# ─────────────────────────────────────────────
# AGENT 1 — VALIDATION RULES
# ─────────────────────────────────────────────
def validate_capbank(rec: dict) -> tuple[str, str | None]:
    zero_thresh  = AGENT1_CFG.get("zero_current_threshold_A", 1.0)
    imbal_thresh = AGENT1_CFG.get("imbalance_threshold_pct", 0.10)

    ia, ib, ic = rec["Ia"], rec["Ib"], rec["Ic"]
    currents   = [ia, ib, ic]

    if not any(c > zero_thresh for c in currents):
        return "SUSPECT", "all-zero current — meter dropout or cap bank offline"

    i_max = max(currents)
    i_min = min(c for c in currents if c > 0)
    if i_max > 0 and (i_max - i_min) / i_max > imbal_thresh:
        return "SUSPECT", f"phase imbalance >10% (Ia={ia:.1f} Ib={ib:.1f} Ic={ic:.1f})"

    return "CLEAN", None


def validate_hyesys(rec: dict) -> tuple[str, str | None]:
    kvar = abs(rec.get("kVAr", 0))
    temp = rec.get("temp_C", 0)
    ia   = rec.get("Ia", 0)
    ib   = rec.get("Ib", 0)
    ic   = rec.get("Ic", 0)

    if kvar > 135:
        return "SUSPECT", f"kVAr {kvar:.1f} exceeds H125 rated output (125 kVAr)"
    if temp > 90:
        return "SUSPECT", f"temperature {temp:.1f}°C above 90°C warning threshold"
    if max(ia, ib, ic) < 1.0 and kvar < 1.0:
        return "SUSPECT", "all currents and kVAr near zero — unit may be offline"

    return "CLEAN", None


# ─────────────────────────────────────────────
# DB WRITERS
# ─────────────────────────────────────────────
def write_record(conn: sqlite3.Connection, rec: dict, tag: str, reason: str | None) -> None:
    now = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO meter_records
              (site_id, timestamp, kW, kVAr, PF, voltage_V,
               kVA, Ia, Ib, Ic, frequency_Hz,
               temp_C, kVAr_A, kVAr_B, kVAr_C, kW_A, kW_B, kW_C, amp_imbalance_pct,
               quality_tag, reject_reason, ingested_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                rec["site_id"],      rec["timestamp"],
                rec["kW"],           rec["kVAr"],          rec["PF"],   rec["voltage_V"],
                rec.get("kVA"),      rec.get("Ia"),         rec.get("Ib"), rec.get("Ic"),
                rec.get("frequency_Hz"),
                rec.get("temp_C"),   rec.get("kVAr_A"),    rec.get("kVAr_B"), rec.get("kVAr_C"),
                rec.get("kW_A"),     rec.get("kW_B"),      rec.get("kW_C"),
                rec.get("amp_imbalance_pct"),
                tag, reason, now,
            ),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed: %s", e)


def write_bms_record(conn: sqlite3.Connection, ts: str, raw: dict) -> None:
    now  = datetime.now(timezone.utc).isoformat()
    d    = raw.get("data", {})
    svav = float(d.get("singleVoltageAvg", 0) or 0)
    try:
        conn.execute(
            "INSERT INTO bms_log (timestamp, single_voltage_avg, raw_json, ingested_at) VALUES (?, ?, ?, ?)",
            (ts, svav, json.dumps(raw), now),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("BMS write failed: %s", e)


# ─────────────────────────────────────────────
# MQTT CALLBACKS
# ─────────────────────────────────────────────
_db_conn: sqlite3.Connection | None = None
msg_count = 0

CAPBANK_TOPICS = {
    "hyesys/data/dev/0086040215999997",
    "hyesys/data/dev/0086040215999996",
}
HYESYS_TOPIC = "stsc/aems/message/26022703840003"


def on_connect(client, userdata, flags, reason_code, properties):
    if reason_code == 0:
        log.info("Connected to %s:%s", BROKER_HOST, BROKER_PORT)
        client.subscribe(TOPICS)
        log.info("Subscribed: %s", [t for t, _ in TOPICS])
    else:
        log.error("Connection refused — code: %s", reason_code)


def on_disconnect(client, userdata, flags, reason_code, properties):
    if reason_code != 0:
        log.warning("Unexpected disconnect (%s) — reconnecting...", reason_code)


def on_message(client, userdata, msg):
    global _db_conn, msg_count
    msg_count += 1
    topic   = msg.topic
    raw_str = msg.payload.decode("utf-8", errors="replace")

    try:
        raw = json.loads(raw_str)
    except json.JSONDecodeError as e:
        log.warning("REJECTED [%s] — invalid JSON: %s", topic, e)
        return

    # ── CapBank topics ────────────────────────────────────────────────
    if topic in CAPBANK_TOPICS:
        record = parse_capbank_payload(raw, topic)
        if record is None:
            return
        tag, reason = validate_capbank(record)
        log.info(
            "[%s] #%d %s | site=%-20s Ia=%6.1f Ib=%6.1f Ic=%6.1f%s",
            tag, msg_count, record["timestamp"], record["site_id"],
            record["Ia"], record["Ib"], record["Ic"],
            f" | {reason}" if reason else "",
        )
        if tag != "REJECTED":
            write_record(_db_conn, record, tag, reason)

    # ── HyESys H125 topic ─────────────────────────────────────────────
    elif topic == HYESYS_TOPIC:
        records = parse_hyesys_message(raw, topic)
        for rec in records:
            rtype = rec.get("_record_type")

            if rtype == "bms":
                svav = rec["raw"].get("data", {}).get("singleVoltageAvg", "?")
                log.info(
                    "[BMS] #%d %s | singleVoltageAvg=%s V",
                    msg_count, rec["timestamp"], svav,
                )
                write_bms_record(_db_conn, rec["timestamp"], rec["raw"])

            elif rtype == "pcs_v3":
                tag, reason = validate_hyesys(rec)
                log.info(
                    "[%s] #%d %s | HYESYS kW=%7.3f kVAr=%7.3f PF=%6.3f "
                    "Ia=%6.1f Ib=%6.1f Ic=%6.1f T=%.1f°C%s",
                    tag, msg_count, rec["timestamp"],
                    rec["kW"], rec["kVAr"], rec["PF"],
                    rec["Ia"], rec["Ib"], rec["Ic"], rec.get("temp_C", 0),
                    f" | {reason}" if reason else "",
                )
                if tag != "REJECTED":
                    write_record(_db_conn, rec, tag, reason)


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    global _db_conn
    _db_conn = init_db(DB_PATH)

    client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2)
    client.username_pw_set(MQTT_USERNAME, MQTT_PASSWORD)
    client.on_connect    = on_connect
    client.on_disconnect = on_disconnect
    client.on_message    = on_message

    log.info("Baoyuan Agent 1 starting — connecting to %s:%s", BROKER_HOST, BROKER_PORT)
    client.connect(BROKER_HOST, BROKER_PORT, KEEPALIVE_SEC)

    try:
        client.loop_forever()
    except KeyboardInterrupt:
        log.info("Stopped. Total messages: %d", msg_count)
    finally:
        client.disconnect()
        if _db_conn:
            _db_conn.close()


if __name__ == "__main__":
    main()
