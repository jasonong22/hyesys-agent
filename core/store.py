"""
SQLite read/write for hyesys.db.
Two tables: meter_records (validated meter data) and sar_log (Agent 2 decisions).
"""

import sqlite3
import logging
from datetime import datetime, timezone
from pathlib import Path

from core.schema import CLEAN, SUSPECT

log = logging.getLogger("hyesys.store")

DB_PATH = Path(__file__).parent.parent / "data" / "hyesys.db"


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db(conn: sqlite3.Connection):
    conn.executescript("""
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
        );

        CREATE UNIQUE INDEX IF NOT EXISTS idx_site_ts
            ON meter_records (site_id, timestamp);

        CREATE TABLE IF NOT EXISTS sar_log (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id         TEXT    NOT NULL,
            timestamp       TEXT    NOT NULL,
            state_kW        REAL    NOT NULL,
            state_kVAr      REAL    NOT NULL,
            state_PF        REAL    NOT NULL,
            state_voltage_V REAL    NOT NULL,
            action          TEXT    NOT NULL,
            action_kVAr     REAL,
            reward_pf_delta REAL,
            reward_fraction REAL,
            outcome         TEXT,
            created_at      TEXT    NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_sar_site
            ON sar_log (site_id, timestamp);
    """)
    conn.commit()
    log.info("Database initialised: %s", DB_PATH)


# ── meter_records ──────────────────────────────────────────────────

def write_record(conn: sqlite3.Connection, record: dict) -> bool:
    """Insert a validated record. Returns True if inserted, False if duplicate."""
    now = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
            """
            INSERT OR IGNORE INTO meter_records
              (site_id, timestamp, kW, kVAr, PF, voltage_V, quality_tag, reject_reason, ingested_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record["site_id"], record["timestamp"],
                record["kW"], record["kVAr"], record["PF"], record["voltage_V"],
                record["quality_tag"], record.get("reject_reason"), now,
            ),
        )
        conn.commit()
        return conn.execute("SELECT changes()").fetchone()[0] > 0
    except sqlite3.Error as e:
        log.error("write_record failed: %s", e)
        return False


def read_clean_records(conn: sqlite3.Connection, site_id: str | None = None,
                       limit: int | None = None) -> list[sqlite3.Row]:
    """Read CLEAN (and SUSPECT) records, optionally filtered by site."""
    tags = (CLEAN, SUSPECT)
    if site_id:
        sql = "SELECT * FROM meter_records WHERE site_id=? AND quality_tag IN (?,?) ORDER BY timestamp"
        params = [site_id, *tags]
    else:
        sql = "SELECT * FROM meter_records WHERE quality_tag IN (?,?) ORDER BY site_id, timestamp"
        params = list(tags)
    if limit:
        sql += f" LIMIT {int(limit)}"
    return conn.execute(sql, params).fetchall()


def get_latest_record(conn: sqlite3.Connection, site_id: str) -> sqlite3.Row | None:
    return conn.execute(
        "SELECT * FROM meter_records WHERE site_id=? AND quality_tag=? ORDER BY timestamp DESC LIMIT 1",
        (site_id, CLEAN),
    ).fetchone()


def get_sites(conn: sqlite3.Connection) -> list[str]:
    rows = conn.execute("SELECT DISTINCT site_id FROM meter_records ORDER BY site_id").fetchall()
    return [r[0] for r in rows]


def get_record_count(conn: sqlite3.Connection, site_id: str | None = None) -> dict:
    if site_id:
        rows = conn.execute(
            "SELECT quality_tag, COUNT(*) FROM meter_records WHERE site_id=? GROUP BY quality_tag",
            (site_id,),
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT quality_tag, COUNT(*) FROM meter_records GROUP BY quality_tag"
        ).fetchall()
    return {r[0]: r[1] for r in rows}


# ── sar_log ────────────────────────────────────────────────────────

def write_sar(conn: sqlite3.Connection, sar: dict) -> bool:
    now = datetime.now(timezone.utc).isoformat()
    try:
        conn.execute(
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
        conn.commit()
        return True
    except sqlite3.Error as e:
        log.error("write_sar failed: %s", e)
        return False


def read_sar(conn: sqlite3.Connection, site_id: str | None = None) -> list[sqlite3.Row]:
    if site_id:
        return conn.execute(
            "SELECT * FROM sar_log WHERE site_id=? ORDER BY timestamp", (site_id,)
        ).fetchall()
    return conn.execute("SELECT * FROM sar_log ORDER BY site_id, timestamp").fetchall()


def get_sar_summary(conn: sqlite3.Connection, site_id: str | None = None) -> dict:
    """Returns outcome counts for the SAR log."""
    if site_id:
        rows = conn.execute(
            "SELECT outcome, COUNT(*) FROM sar_log WHERE site_id=? GROUP BY outcome", (site_id,)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT outcome, COUNT(*) FROM sar_log GROUP BY outcome"
        ).fetchall()
    return {r[0]: r[1] for r in rows}
