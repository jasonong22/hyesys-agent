"""
HyESys MQTT Ingestion Script
Subscribes to AC meter topics (hyesys/#) and BMS cabinet topics (stsc/aems/message/+).
Routes incoming JSON payloads by 'type' field:
  - AC meter data  → Agent 1 validation → meter_records table
  - BMS data       → BMS validation     → bms_records table
Writes CLEAN and SUSPECT records to hyesys.db for Agent 2 processing.
"""

import json
import logging
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path

import paho.mqtt.client as mqtt

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
BROKER_HOST   = "loragw.advastech.com"
BROKER_PORT   = 1883
KEEPALIVE_SEC = 60

MQTT_USERNAME = "TYKJadmin"
MQTT_PASSWORD = "TYKJ2018."

TOPICS = [
    ("hyesys/#",               0),   # AC meter data
    ("stsc/aems/message/+",    0),   # BMS cabinet data
]

DB_PATH = Path(__file__).parent / "data" / "hyesys.db"

# Maps BMS cabinetId → site_id (add entries as new cabinets are deployed)
CABINET_TO_SITE: dict[str, str] = {
    "e45f01FFFEe9380f": "INLET-METER-MAR26",
}

# ─────────────────────────────────────────────
# EXPECTED SCHEMAS
# ─────────────────────────────────────────────
METER_REQUIRED = ["site_id", "timestamp", "kW", "kVAr", "PF", "voltage_V"]

BMS_REQUIRED   = ["type", "cabinetId", "reportTimeTs", "soc", "soh", "tempMain"]

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("hyesys.ingest")

# ─────────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────────
def init_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)

    # AC meter records (Agent 1 output)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS meter_records (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id       TEXT    NOT NULL,
            timestamp     TEXT    NOT NULL,
            kW            REAL    NOT NULL,
            kVAr          REAL    NOT NULL,
            PF            REAL    NOT NULL,
            voltage_V     REAL    NOT NULL,
            quality_tag   TEXT    NOT NULL,
            reject_reason TEXT,
            ingested_at   TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_site_ts
        ON meter_records (site_id, timestamp)
    """)

    # BMS records (battery health data)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS bms_records (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id           TEXT    NOT NULL,
            cabinet_id        TEXT    NOT NULL,
            timestamp         TEXT    NOT NULL,          -- ISO 8601 (converted from reportTimeTs)
            soc               REAL    NOT NULL,          -- State of Charge  0–100 %
            soh               REAL    NOT NULL,          -- State of Health   0–100 %
            temp_main         REAL    NOT NULL,          -- °C
            single_volt_diff  REAL,                      -- max cell-voltage spread (V)
            raw_payload       TEXT,                      -- full JSON for debugging
            quality_tag       TEXT    NOT NULL,          -- CLEAN / SUSPECT / REJECTED
            reject_reason     TEXT,
            ingested_at       TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_bms_cabinet_ts
        ON bms_records (cabinet_id, timestamp)
    """)

    conn.commit()
    log.info("Database ready: %s", db_path)
    return conn

# ─────────────────────────────────────────────
# AGENT 1 — METER VALIDATION RULES
# ─────────────────────────────────────────────
def validate_meter(payload: dict) -> tuple[str, str | None]:
    kW      = payload["kW"]
    kVAr    = payload["kVAr"]
    pf      = payload["PF"]
    voltage = payload["voltage_V"]

    if voltage <= 0:
        return "REJECTED", "voltage <= 0"
    if not (-10000 <= kW <= 10000):
        return "REJECTED", f"kW out of plausible range: {kW}"
    if not (-2000 <= kVAr <= 2000):
        return "REJECTED", f"kVAr out of plausible range: {kVAr}"
    if not (-1.0 <= pf <= 1.0):
        return "REJECTED", f"PF outside [-1, 1]: {pf}"

    if kW == 0 and kVAr == 0 and pf == 0:
        return "SUSPECT", "all-zero row — possible meter dropout"
    if abs(pf) == 1.0:
        return "SUSPECT", f"PF firmware saturation at {pf}"
    if not (195.5 <= voltage <= 264.5):
        return "SUSPECT", f"voltage outside 230V ±15% range: {voltage}V"

    return "CLEAN", None

# ─────────────────────────────────────────────
# AGENT 1 — BMS VALIDATION RULES
# ─────────────────────────────────────────────
def validate_bms(payload: dict) -> tuple[str, str | None]:
    soc  = payload["soc"]
    soh  = payload["soh"]
    temp = payload["tempMain"]
    svd  = payload.get("singleVoltageDiff")   # optional field

    # REJECTED: physically impossible
    if not (0 <= soc <= 100):
        return "REJECTED", f"SOC out of range: {soc}%"
    if not (0 <= soh <= 100):
        return "REJECTED", f"SOH out of range: {soh}%"
    if not (-20 <= temp <= 80):
        return "REJECTED", f"temperature out of range: {temp}°C"

    # SUSPECT: degraded or at-risk conditions
    if soc < 10:
        return "SUSPECT", f"SOC critically low: {soc}%"
    if temp > 50:
        return "SUSPECT", f"battery temperature high: {temp}°C"
    if soh < 80:
        return "SUSPECT", f"SOH below 80%: {soh}%"
    if svd is not None and svd > 0.1:
        return "SUSPECT", f"high cell-voltage spread: {svd}V"

    return "CLEAN", None

# ─────────────────────────────────────────────
# HANDLERS
# ─────────────────────────────────────────────
def handle_meter(conn: sqlite3.Connection, topic: str, payload: dict) -> None:
    missing = [f for f in METER_REQUIRED if f not in payload]
    if missing:
        log.warning("METER REJECTED [%s] — missing fields: %s", topic, missing)
        return

    tag, reason = validate_meter(payload)
    log.info(
        "METER [%s] %s | site=%s kW=%.2f kVAr=%.2f PF=%.3f V=%.1f%s",
        tag, payload["timestamp"], payload["site_id"],
        payload["kW"], payload["kVAr"], payload["PF"], payload["voltage_V"],
        f" | {reason}" if reason else "",
    )

    if tag == "REJECTED":
        return

    ingested_at = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO meter_records
              (site_id, timestamp, kW, kVAr, PF, voltage_V, quality_tag, reject_reason, ingested_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                payload["site_id"], payload["timestamp"],
                payload["kW"], payload["kVAr"], payload["PF"], payload["voltage_V"],
                tag, reason, ingested_at,
            ),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed (meter): %s", e)


def handle_bms(conn: sqlite3.Connection, topic: str, payload: dict, raw: str) -> None:
    missing = [f for f in BMS_REQUIRED if f not in payload]
    if missing:
        log.warning("BMS REJECTED [%s] — missing fields: %s", topic, missing)
        return

    cabinet_id = payload["cabinetId"]
    site_id    = CABINET_TO_SITE.get(cabinet_id, cabinet_id)  # fall back to cabinetId if unmapped

    # Convert Unix ms timestamp → ISO 8601
    ts_ms = payload["reportTimeTs"]
    try:
        ts_iso = datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc).isoformat()
    except (OSError, OverflowError, ValueError):
        log.warning("BMS REJECTED [%s] — invalid reportTimeTs: %s", topic, ts_ms)
        return

    tag, reason = validate_bms(payload)
    log.info(
        "BMS [%s] %s | site=%s SOC=%.1f%% SOH=%.1f%% T=%.1f°C%s",
        tag, ts_iso, site_id,
        payload["soc"], payload["soh"], payload["tempMain"],
        f" | {reason}" if reason else "",
    )

    if tag == "REJECTED":
        return

    svd = payload.get("singleVoltageDiff")
    ingested_at = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO bms_records
              (site_id, cabinet_id, timestamp, soc, soh, temp_main,
               single_volt_diff, raw_payload, quality_tag, reject_reason, ingested_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                site_id, cabinet_id, ts_iso,
                payload["soc"], payload["soh"], payload["tempMain"],
                svd, raw,
                tag, reason, ingested_at,
            ),
        )
        conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed (bms): %s", e)

# ─────────────────────────────────────────────
# MQTT CALLBACKS
# ─────────────────────────────────────────────
_db_conn: sqlite3.Connection | None = None

def on_connect(client, userdata, flags, reason_code, properties):
    if reason_code == 0:
        log.info("Connected to broker %s:%s", BROKER_HOST, BROKER_PORT)
        client.subscribe(TOPICS)
        log.info("Subscribed to topics: %s", [t for t, _ in TOPICS])
    else:
        log.error("Connection refused — reason code: %s", reason_code)

def on_disconnect(client, userdata, flags, reason_code, properties):
    if reason_code != 0:
        log.warning("Unexpected disconnect (code %s) — will auto-reconnect", reason_code)

def on_message(client, userdata, msg):
    global _db_conn
    topic = msg.topic
    raw   = msg.payload.decode("utf-8", errors="replace")

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as e:
        log.warning("REJECTED [%s] — invalid JSON: %s | raw: %.120s", topic, e, raw)
        return

    msg_type = payload.get("type", "").lower()

    if msg_type == "bms" or topic.startswith("stsc/aems/message/"):
        handle_bms(_db_conn, topic, payload, raw)
    else:
        handle_meter(_db_conn, topic, payload)

def on_log(client, userdata, level, buf):
    if level == mqtt.MQTT_LOG_ERR:
        log.debug("MQTT internal: %s", buf)

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
    client.on_log        = on_log

    log.info("Connecting to %s:%s …", BROKER_HOST, BROKER_PORT)
    client.connect(BROKER_HOST, BROKER_PORT, KEEPALIVE_SEC)

    try:
        client.loop_forever()
    except KeyboardInterrupt:
        log.info("Shutting down.")
    finally:
        client.disconnect()
        _db_conn.close()

if __name__ == "__main__":
    main()
