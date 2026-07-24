"""SQLite layer — the single source of truth.

WAL mode, one connection per request (SQLite is fine with this at household
scale). Schema is created/updated idempotently on boot. The Excel workbook is
GENERATED from this DB (see app/export/steuer.py) — it is never the live store.
"""
from __future__ import annotations

import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Iterator, Optional

from . import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS receipts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kind TEXT NOT NULL CHECK(kind IN ('out','in')),
  serial INTEGER NOT NULL,
  date TEXT,
  merchant TEXT,
  article TEXT,
  business_personal TEXT,
  client_name TEXT,
  treatment TEXT,
  total REAL DEFAULT 0,
  currency TEXT DEFAULT '',
  eur REAL,
  fx_rate REAL,
  fx_date TEXT,
  receipt_code TEXT DEFAULT '',
  note TEXT DEFAULT '',
  location TEXT DEFAULT '',
  map_url TEXT DEFAULT '',
  status TEXT DEFAULT 'OK',
  photo TEXT DEFAULT '',
  generated_card INTEGER DEFAULT 0,
  captured_at TEXT DEFAULT '',
  saved_at TEXT DEFAULT '',
  operation_id TEXT,
  operation_hash TEXT,
  UNIQUE(kind, serial)
);
CREATE TABLE IF NOT EXISTS trips (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  purpose TEXT NOT NULL,
  start TEXT,
  end TEXT,
  nights INTEGER DEFAULT 0,
  meal_reduction REAL DEFAULT 0,
  overnight_cost REAL DEFAULT 0,
  incidentals REAL DEFAULT 0,
  note TEXT DEFAULT '',
  created_at TEXT
);
CREATE TABLE IF NOT EXISTS legs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  trip_id INTEGER NOT NULL REFERENCES trips(id) ON DELETE CASCADE,
  seq INTEGER NOT NULL,
  date TEXT,
  from_place TEXT,
  to_place TEXT,
  km REAL DEFAULT 0,
  overnight INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS treatments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  price REAL DEFAULT 0,
  pos INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS places (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  address TEXT DEFAULT '',
  pos INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS open_trip (
  id INTEGER PRIMARY KEY CHECK(id = 1),
  payload TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS open_trip_legs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  trip_token TEXT NOT NULL,
  operation_id TEXT NOT NULL UNIQUE,
  operation_hash TEXT NOT NULL,
  seq INTEGER NOT NULL,
  date TEXT,
  from_place TEXT,
  to_place TEXT,
  km REAL DEFAULT 0,
  overnight INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS meta (
  key TEXT PRIMARY KEY,
  value TEXT
);
CREATE INDEX IF NOT EXISTS idx_receipts_kind_date ON receipts(kind, date);
CREATE INDEX IF NOT EXISTS idx_legs_trip ON legs(trip_id, seq);
"""


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(config.DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


@contextmanager
def get_db() -> Iterator[sqlite3.Connection]:
    conn = connect()
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_schema(db: sqlite3.Connection) -> None:
    """Create/refresh the schema on an arbitrary connection. Shared by the app
    boot path (init_db) and the v1 migration tool (app.migrate), which operates
    on an explicitly given target database."""
    db.executescript(SCHEMA)
    # Forward-only, idempotent schema upgrade for databases created before
    # receipt operation IDs were introduced.
    receipt_cols = {
        row["name"] for row in db.execute("PRAGMA table_info(receipts)").fetchall()
    }
    if "operation_id" not in receipt_cols:
        db.execute("ALTER TABLE receipts ADD COLUMN operation_id TEXT")
    if "operation_hash" not in receipt_cols:
        db.execute("ALTER TABLE receipts ADD COLUMN operation_hash TEXT")
    db.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_receipts_operation_id "
        "ON receipts(operation_id) WHERE operation_id IS NOT NULL"
    )
    # Seed tax parameters once; afterwards the DB values rule.
    for k, v in config.TAX_DEFAULTS.items():
        db.execute(
            "INSERT OR IGNORE INTO meta(key, value) VALUES (?, ?)",
            (f"tax.{k}", str(v)),
        )
    # Default place anchor (home) if no places yet.
    n = db.execute("SELECT COUNT(*) c FROM places").fetchone()["c"]
    if n == 0:
        db.execute("INSERT INTO places(name, address, pos) VALUES ('Zuhause', '', 0)")


def init_db() -> None:
    with get_db() as db:
        init_schema(db)


# --- meta helpers -----------------------------------------------------------

def get_meta(db: sqlite3.Connection, key: str, default: str = "") -> str:
    row = db.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else default


def set_meta(db: sqlite3.Connection, key: str, value: Any) -> None:
    db.execute(
        "INSERT INTO meta(key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, str(value)),
    )


def tax_rates(db: sqlite3.Connection) -> dict:
    return {
        k: float(get_meta(db, f"tax.{k}", str(v)))
        for k, v in config.TAX_DEFAULTS.items()
    }


# --- receipts ---------------------------------------------------------------

def next_serial(db: sqlite3.Connection, kind: str) -> int:
    row = db.execute(
        "SELECT COALESCE(MAX(serial), 0) + 1 AS s FROM receipts WHERE kind = ?", (kind,)
    ).fetchone()
    return int(row["s"])


def insert_receipt(db: sqlite3.Connection, r: dict) -> int:
    cols = ", ".join(r.keys())
    marks = ", ".join("?" for _ in r)
    cur = db.execute(f"INSERT INTO receipts ({cols}) VALUES ({marks})", list(r.values()))
    return int(cur.lastrowid)


def receipt_by_operation(db: sqlite3.Connection, operation_id: str) -> sqlite3.Row | None:
    return db.execute(
        "SELECT * FROM receipts WHERE operation_id = ?", (operation_id,)
    ).fetchone()


# --- open trip (durable across restarts, single slot like v1) ----------------

def read_open_trip(db: sqlite3.Connection) -> Optional[dict]:
    row = db.execute("SELECT payload FROM open_trip WHERE id = 1").fetchone()
    if not row:
        return None
    try:
        trip = json.loads(row["payload"])
    except (ValueError, TypeError):
        return None
    token = str(trip.get("tripToken") or "")
    if token:
        legs = db.execute(
            "SELECT date, from_place, to_place, km, overnight, operation_id "
            "FROM open_trip_legs WHERE trip_token = ? ORDER BY seq, id", (token,)
        ).fetchall()
        if legs:
            trip["legs"] = [
                {"date": leg["date"], "from": leg["from_place"],
                 "to": leg["to_place"], "km": leg["km"],
                 "overnight": bool(leg["overnight"]),
                 "operationId": leg["operation_id"]}
                for leg in legs
            ]
    return trip


def write_open_trip(db: sqlite3.Connection, trip: dict) -> None:
    db.execute(
        "INSERT INTO open_trip(id, payload) VALUES (1, ?) "
        "ON CONFLICT(id) DO UPDATE SET payload = excluded.payload",
        (json.dumps(trip, ensure_ascii=False),),
    )


def clear_open_trip(db: sqlite3.Connection) -> None:
    db.execute("DELETE FROM open_trip_legs")
    db.execute("DELETE FROM open_trip WHERE id = 1")


class TripLegConflict(Exception):
    """The client attempted to append to a stale or conflicting trip leg."""


def append_open_trip_leg(conn: sqlite3.Connection, trip_token: str,
                         operation_id: str, operation_hash: str, leg: dict):
    """Append one leg under a write lock; retries return the existing leg."""
    conn.execute("BEGIN IMMEDIATE")
    row = conn.execute("SELECT payload FROM open_trip WHERE id = 1").fetchone()
    if not row:
        raise TripLegConflict("No open trip. Start one first.")
    try:
        payload = json.loads(row["payload"])
        active_token = str(payload.get("tripToken") or "")
    except (TypeError, ValueError):
        payload = {}
        active_token = ""
    if not active_token:
        active_token = str(uuid.uuid4())
        legacy_legs = payload.pop("legs", []) or []
        payload["tripToken"] = active_token
        conn.execute("UPDATE open_trip SET payload = ? WHERE id = 1",
                     (json.dumps(payload, ensure_ascii=False),))
        for seq, old_leg in enumerate(legacy_legs, start=1):
            old_op = str(old_leg.get("operationId") or
                         f"legacy-{active_token}-{seq}")
            conn.execute(
                "INSERT OR IGNORE INTO open_trip_legs "
                "(trip_token, operation_id, operation_hash, seq, date, from_place, "
                "to_place, km, overnight) VALUES (?,?,?,?,?,?,?,?,?)",
                (active_token, old_op, old_op, seq, old_leg.get("date", ""),
                 old_leg.get("from", ""), old_leg.get("to", ""),
                 float(old_leg.get("km", 0)), 1 if old_leg.get("overnight") else 0),
            )
    trip_token = trip_token or active_token
    if trip_token != active_token:
        raise TripLegConflict("This trip is stale. Refresh and try again.")
    existing = conn.execute(
        "SELECT * FROM open_trip_legs WHERE operation_id = ?", (operation_id,)
    ).fetchone()
    if existing:
        if existing["operation_hash"] != operation_hash:
            raise TripLegConflict("operationId was already used for different leg data.")
        return ({"date": existing["date"], "from": existing["from_place"],
                 "to": existing["to_place"], "km": existing["km"],
                 "overnight": bool(existing["overnight"]),
                 "operationId": operation_id}, True)
    seq = conn.execute(
        "SELECT COALESCE(MAX(seq), 0) + 1 FROM open_trip_legs "
        "WHERE trip_token = ?", (trip_token,)
    ).fetchone()[0]
    conn.execute(
        "INSERT INTO open_trip_legs "
        "(trip_token, operation_id, operation_hash, seq, date, from_place, "
        "to_place, km, overnight) VALUES (?,?,?,?,?,?,?,?,?)",
        (trip_token, operation_id, operation_hash, seq, leg["date"],
         leg["from"], leg["to"], leg["km"], 1 if leg["overnight"] else 0),
    )
    result = dict(leg)
    result["operationId"] = operation_id
    return result, False


# --- finished trips ----------------------------------------------------------

def insert_trip(db: sqlite3.Connection, trip: dict, legs: list) -> int:
    cur = db.execute(
        "INSERT INTO trips (purpose, start, end, nights, meal_reduction, "
        "overnight_cost, incidentals, note, created_at) VALUES (?,?,?,?,?,?,?,?,?)",
        (
            trip.get("purpose", ""), trip.get("start", ""), trip.get("end", ""),
            int(trip.get("nights", 0)), float(trip.get("meal_reduction", 0)),
            float(trip.get("overnight_cost", 0)), float(trip.get("incidentals", 0)),
            trip.get("note", ""), datetime.now().isoformat(timespec="seconds"),
        ),
    )
    trip_id = int(cur.lastrowid)
    for i, leg in enumerate(legs, start=1):
        db.execute(
            "INSERT INTO legs (trip_id, seq, date, from_place, to_place, km, overnight) "
            "VALUES (?,?,?,?,?,?,?)",
            (
                trip_id, i, leg.get("date", ""), leg.get("from", ""),
                leg.get("to", ""), float(leg.get("km", 0)),
                1 if leg.get("overnight") else 0,
            ),
        )
    return trip_id


def touch_last_activity(db: sqlite3.Connection) -> None:
    set_meta(db, "last_activity", datetime.now().isoformat(timespec="seconds"))
