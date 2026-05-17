"""
HyESys MQTT Ingestion Script
Subscribes to HyESys meter topics, validates incoming JSON payloads (Agent 1 rules),
and writes CLEAN records to hyesys.db for Agent 2 processing.
"""

import json
import logging
import sqlite3
import time
from datetime import datetime, timezone
from pathlib import Path

import paho.mqtt.client as mqtt

# ─────────────────────────────────────────────
# CONFIGURATION — fill in before running
# ─────────────────────────────────────────────
BROKER_HOST     = "192.168.1.100"       # MQTT broker IP or hostname
BROKER_PORT     = 1883                  # Default: 1883 (unencrypted), 8883 (TLS)
KEEPALIVE_SEC   = 60

# Leave as None if broker has no authentication
MQTT_USERNAME   = None                  # e.g. "hyesys_user"
MQTT_PASSWORD   = None                  # e.g. "s3cr3t"

# Topics to subscribe to — use # wildcard to catch all site subtopics
# e.g. "hyesys/+/state" matches hyesys/INLET-METER-MAR26/state
TOPICS = [
    ("hyesys/#", 0),                    # QoS 0 — best effort
]

# Path to SQLite database
DB_PATH = Path(__file__).parent / "data" / "hyesys.db"

# ─────────────────────────────────────────────
# EXPECTED JSON PAYLOAD SCHEMA
# ─────────────────────────────────────────────
# {
#   "site_id":   "INLET-METER-MAR26",
#   "timestamp": "2026-05-16T10:00:00+08:00",  # ISO 8601
#   "kW":        12.5,
#   "kVAr":      3.2,
#   "PF":        0.97,
#   "voltage_V": 230.1
# }
REQUIRED_FIELDS = ["site_id", "timestamp", "kW", "kVAr", "PF", "voltage_V"]

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
    conn.execute("""
        CREATE TABLE IF NOT EXISTS meter_records (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id     TEXT    NOT NULL,
            timestamp   TEXT    NOT NULL,
            kW          REAL    NOT NULL,
            kVAr        REAL    NOT NULL,
            PF          REAL    NOT NULL,
            voltage_V   REAL    NOT NULL,
            quality_tag TEXT    NOT NULL,   -- CLEAN / SUSPECT / REJECTED
            reject_reason TEXT,             -- populated for SUSPECT and REJECTED
            ingested_at TEXT    NOT NULL
        )
    """)
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_site_ts
        ON meter_records (site_id, timestamp)
    """)
    conn.commit()
    log.info("Database ready: %s", db_path)
    return conn

# ─────────────────────────────────────────────
# AGENT 1 — VALIDATION RULES
# ─────────────────────────────────────────────
def validate(payload: dict) -> tuple[str, str | None]:
    """
    Returns (tag, reason) where tag is CLEAN / SUSPECT / REJECTED.
    Reason is None for CLEAN records.
    """
    kW      = payload["kW"]
    kVAr    = payload["kVAr"]
    pf      = payload["PF"]
    voltage = payload["voltage_V"]

    # REJECTED: physically impossible values
    if voltage <= 0:
        return "REJECTED", "voltage <= 0"
    if not (-10000 <= kW <= 10000):
        return "REJECTED", f"kW out of plausible range: {kW}"
    if not (-2000 <= kVAr <= 2000):
        return "REJECTED", f"kVAr out of plausible range: {kVAr}"
    if not (-1.0 <= pf <= 1.0):
        return "REJECTED", f"PF firmware bug — value outside [-1, 1]: {pf}"

    # SUSPECT: zero-row (meter offline or comms dropout)
    if kW == 0 and kVAr == 0 and pf == 0:
        return "SUSPECT", "all-zero row — possible meter dropout"

    # SUSPECT: PF at exactly ±1.0 (firmware saturation artefact)
    if abs(pf) == 1.0:
        return "SUSPECT", f"PF firmware saturation at {pf}"

    # SUSPECT: voltage out of normal Singapore LV range (230 V ±15%)
    if not (195.5 <= voltage <= 264.5):
        return "SUSPECT", f"voltage outside 230V ±15% range: {voltage}V"

    return "CLEAN", None

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

    # ── Parse JSON ──
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as e:
        log.warning("REJECTED [%s] — invalid JSON: %s | raw: %.120s", topic, e, raw)
        return

    # ── Check required fields ──
    missing = [f for f in REQUIRED_FIELDS if f not in payload]
    if missing:
        log.warning("REJECTED [%s] — missing fields: %s", topic, missing)
        return

    # ── Agent 1 validation ──
    tag, reason = validate(payload)

    log.info(
        "[%s] %s | site=%s kW=%.2f kVAr=%.2f PF=%.3f V=%.1f%s",
        tag, payload["timestamp"], payload["site_id"],
        payload["kW"], payload["kVAr"], payload["PF"], payload["voltage_V"],
        f" | {reason}" if reason else "",
    )

    # ── Write CLEAN records to DB (skip REJECTED) ──
    if tag == "REJECTED":
        return

    ingested_at = datetime.now(timezone.utc).isoformat()
    try:
        _db_conn.execute(
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
        _db_conn.commit()
    except sqlite3.Error as e:
        log.error("DB write failed: %s", e)

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

    if MQTT_USERNAME:
        client.username_pw_set(MQTT_USERNAME, MQTT_PASSWORD)

    client.on_connect    = on_connect
    client.on_disconnect = on_disconnect
    client.on_message    = on_message
    client.on_log        = on_log

    log.info("Connecting to %s:%s …", BROKER_HOST, BROKER_PORT)
    client.connect(BROKER_HOST, BROKER_PORT, KEEPALIVE_SEC)

    try:
        client.loop_forever()           # blocks; handles reconnects automatically
    except KeyboardInterrupt:
        log.info("Shutting down.")
    finally:
        client.disconnect()
        _db_conn.close()

if __name__ == "__main__":
    main()
