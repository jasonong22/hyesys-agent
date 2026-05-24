"""
╔══════════════════════════════════════════════════════════════════════╗
║  AGENT 1 — BAOYUAN SITE                                             ║
║  Data Quality & Ingestion — CapBank1 + CapBank2 MQTT meters         ║
║  Copied from: templates/agent1_master.py  v1.0  (2026-05-24)        ║
╚══════════════════════════════════════════════════════════════════════╝

SITE: Baoyuan Industrial (诸暨市葆元实业有限公司)
      Monitors: CapBank1 (device 0086040215999997)
                CapBank2 (device 0086040215999996)
      HyESys H125 installed at Cabinet A / Feeder 2

── SITE CHANGELOG (changes from master) ─────────────────────────────
2026-05-24  initial copy from master v1.0

DATA SOURCE CHANGES (vs master):
  • Live MQTT ingestion (not batch CSV) — paho-mqtt client added
  • Subscribes to topics hyesys/data/dev/<device_id>
  • Custom parse_payload() to unpack nested MQTT JSON structure
  • Timestamp sourced from raw["sendtime"] (Unix seconds → ISO UTC)
  • Voltage: phase voltages Ua/Ub/Uc preferred; line Uab/Ubc/Uca fallback

METER TYPE OVERRIDE — CapBank current-only instruments:
  • Baoyuan CapBank meters measure Ia/Ib/Ic ONLY
  • kW, kVAr, PF, voltage_V are always 0 — this is normal; not an anomaly
  • Master R3 (non-numeric), R4 (voltage≤0), R5 (kW range), R8 (all-zero),
    R9 (PF saturation), R10 (voltage range) would all misfire here
  • Replaced entire validate() with current-only logic:
      NEW rule: SUSPECT if all phase currents ≤ zero_current_threshold_A
      NEW rule: SUSPECT if phase current imbalance > imbalance_threshold_pct

PARAMETERS NOT IN MASTER:
  • zero_current_threshold_A = 1.0  (in agent1_config.json)
  • imbalance_threshold_pct  = 0.10 (in agent1_config.json)
  • MQTT broker: loragw.advastech.com:1883
  • DEVICE_TO_SITE mapping dict

DB SCHEMA ADDITIONS (vs master):
  • meter_records table extended with Ia, Ib, Ic, frequency_Hz columns
  • sar_log table added in agent1.py (master puts this in agent2)
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
        import json as _json
        return _json.loads(_CONFIG_PATH.read_text())
    return {"zero_current_threshold_A": 1.0, "imbalance_threshold_pct": 0.10}

AGENT1_CFG = _load_config()

BROKER_HOST   = "loragw.advastech.com"
BROKER_PORT   = 1883
KEEPALIVE_SEC = 60
MQTT_USERNAME = "TYKJadmin"
MQTT_PASSWORD = "TYKJ2018."

TOPICS = [
    ("hyesys/data/dev/0086040215999997", 0),  # CapBank1
    ("hyesys/data/dev/0086040215999996", 0),  # CapBank2
]

# Maps device ID (from MQTT topic suffix and data key) → site_id
DEVICE_TO_SITE: dict[str, str] = {
    "0086040215999997": "BAOYUAN-CAPBANK1",
    "0086040215999996": "BAOYUAN-CAPBANK2",
}

DB_PATH = Path(__file__).parent / "data" / "baoyuan.db"

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
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
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id       TEXT NOT NULL,
            timestamp     TEXT NOT NULL,
            kW            REAL NOT NULL,
            kVAr          REAL NOT NULL,
            PF            REAL NOT NULL,
            voltage_V     REAL NOT NULL,
            kVA           REAL,
            Ia            REAL,
            Ib            REAL,
            Ic            REAL,
            frequency_Hz  REAL,
            quality_tag   TEXT NOT NULL,
            reject_reason TEXT,
            ingested_at   TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_site_ts
        ON meter_records (site_id, timestamp)
    """)
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
    conn.commit()
    log.info("Database ready: %s", db_path)
    return conn


# ─────────────────────────────────────────────
# PAYLOAD PARSER
# ─────────────────────────────────────────────
def parse_payload(raw: dict, topic: str) -> dict | None:
    """
    Parse the nested HyESys meter payload format.

    Structure: raw["reported"]["0_5_<device_id>"] = meter fields dict.
    Voltage: prefer phase voltages (Ua/Ub/Uc); fall back to line voltages / √3.
    Timestamp: converted from raw["sendtime"] (Unix seconds) to ISO 8601 UTC.
    """
    reported = raw.get("reported", {})
    if not reported:
        log.warning("REJECTED [%s] — missing 'reported' key", topic)
        return None

    data_key = next(iter(reported), None)
    if not data_key:
        log.warning("REJECTED [%s] — empty 'reported' object", topic)
        return None

    m = reported[data_key]

    # Device ID: strip the '0_5_' routing prefix
    device_id = data_key.replace("0_5_", "")
    site_id   = DEVICE_TO_SITE.get(device_id, device_id)

    # Voltage — phase voltages preferred, line voltages as fallback
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

    # Timestamp from sendtime (Unix seconds → ISO 8601 UTC)
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
        "state":        str(m.get("state", "")),
    }


# ─────────────────────────────────────────────
# AGENT 1 — VALIDATION RULES
# ─────────────────────────────────────────────
def validate(rec: dict) -> tuple[str, str | None]:
    """
    Tag each record CLEAN / SUSPECT / REJECTED.

    Baoyuan CapBank meters are current-only instruments — Ia/Ib/Ic are the only
    meaningful fields. Voltage, kW, kVAr, PF will always be 0; this is normal
    and must not trigger SUSPECT or REJECTED.
    """
    zero_thresh    = AGENT1_CFG.get("zero_current_threshold_A", 1.0)
    imbal_thresh   = AGENT1_CFG.get("imbalance_threshold_pct", 0.10)

    ia, ib, ic = rec["Ia"], rec["Ib"], rec["Ic"]
    currents    = [ia, ib, ic]
    any_current = any(c > zero_thresh for c in currents)

    # ── SUSPECT: no current at all → meter dropout or cap bank offline ──
    if not any_current:
        return "SUSPECT", "all-zero current — meter dropout or cap bank offline"

    # ── SUSPECT: phase current imbalance above threshold ─────────────────
    i_max = max(currents)
    i_min = min(c for c in currents if c > 0)
    if i_max > 0 and (i_max - i_min) / i_max > imbal_thresh:
        return "SUSPECT", (
            f"phase current imbalance >10% "
            f"(Ia={ia:.1f} Ib={ib:.1f} Ic={ic:.1f})"
        )

    return "CLEAN", None


# ─────────────────────────────────────────────
# DB WRITER
# ─────────────────────────────────────────────
def write_record(conn: sqlite3.Connection, rec: dict, tag: str, reason: str | None) -> None:
    now = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO meter_records
              (site_id, timestamp, kW, kVAr, PF, voltage_V,
               kVA, Ia, Ib, Ic, frequency_Hz,
               quality_tag, reject_reason, ingested_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                rec["site_id"], rec["timestamp"],
                rec["kW"], rec["kVAr"], rec["PF"], rec["voltage_V"],
                rec.get("kVA"), rec.get("Ia"), rec.get("Ib"), rec.get("Ic"),
                rec.get("frequency_Hz"),
                tag, reason, now,
            ),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed: %s", e)


# ─────────────────────────────────────────────
# MQTT CALLBACKS
# ─────────────────────────────────────────────
_db_conn: sqlite3.Connection | None = None
msg_count = 0


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
    topic  = msg.topic
    raw_str = msg.payload.decode("utf-8", errors="replace")

    try:
        raw = json.loads(raw_str)
    except json.JSONDecodeError as e:
        log.warning("REJECTED [%s] — invalid JSON: %s", topic, e)
        return

    record = parse_payload(raw, topic)
    if record is None:
        return

    tag, reason = validate(record)
    log.info(
        "[%s] #%d %s | site=%-20s kW=%7.3f kVAr=%7.3f PF=%.3f V=%5.1f  Ia=%6.1f Ib=%6.1f Ic=%6.1f%s",
        tag, msg_count, record["timestamp"], record["site_id"],
        record["kW"], record["kVAr"], record["PF"], record["voltage_V"],
        record["Ia"], record["Ib"], record["Ic"],
        f"  | {reason}" if reason else "",
    )

    if tag == "REJECTED":
        return

    write_record(_db_conn, record, tag, reason)


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
