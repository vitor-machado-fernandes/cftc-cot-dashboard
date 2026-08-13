"""SQLite storage for normalized CONAB crop progress rows."""

from __future__ import annotations

from pathlib import Path
import sqlite3

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
DB_PATH = BASE_DIR / "data" / "conab_progress.db"
TABLE = "conab_progress"


def _connect(db_path: str | Path = DB_PATH) -> sqlite3.Connection:
    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {TABLE} (
            bulletin_date TEXT NOT NULL,
            crop TEXT NOT NULL,
            state TEXT NOT NULL,
            season TEXT NOT NULL,
            planting_pct REAL,
            harvest_pct REAL,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (bulletin_date, crop, state, season)
        )
        """
    )
    return conn


def load_data(db_path: str | Path = DB_PATH) -> pd.DataFrame:
    """Load all stored CONAB progress rows."""

    with _connect(db_path) as conn:
        df = pd.read_sql_query(f"SELECT * FROM {TABLE}", conn)
    if df.empty:
        return pd.DataFrame(
            columns=["bulletin_date", "crop", "state", "season", "planting_pct", "harvest_pct"]
        )
    df["bulletin_date"] = pd.to_datetime(df["bulletin_date"])
    return df.drop(columns=["updated_at"], errors="ignore")


def latest_bulletin_date(db_path: str | Path = DB_PATH) -> pd.Timestamp | None:
    """Return the newest stored bulletin date."""

    with _connect(db_path) as conn:
        value = conn.execute(f"SELECT MAX(bulletin_date) FROM {TABLE}").fetchone()[0]
    return pd.to_datetime(value) if value else None


def upsert_data(df: pd.DataFrame, db_path: str | Path = DB_PATH) -> int:
    """Upsert normalized rows and return the number of new unique keys added."""

    if df.empty:
        return 0

    rows = df.copy()
    rows["bulletin_date"] = pd.to_datetime(rows["bulletin_date"]).dt.strftime("%Y-%m-%d")
    rows = rows.drop_duplicates(["bulletin_date", "crop", "state", "season"], keep="last")

    with _connect(db_path) as conn:
        before = conn.execute(f"SELECT COUNT(*) FROM {TABLE}").fetchone()[0]
        conn.executemany(
            f"""
            INSERT INTO {TABLE} (
                bulletin_date, crop, state, season, planting_pct, harvest_pct, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(bulletin_date, crop, state, season) DO UPDATE SET
                planting_pct=excluded.planting_pct,
                harvest_pct=excluded.harvest_pct,
                updated_at=CURRENT_TIMESTAMP
            """,
            rows[["bulletin_date", "crop", "state", "season", "planting_pct", "harvest_pct"]].itertuples(index=False, name=None),
        )
        conn.commit()
        after = conn.execute(f"SELECT COUNT(*) FROM {TABLE}").fetchone()[0]
    return int(after - before)
