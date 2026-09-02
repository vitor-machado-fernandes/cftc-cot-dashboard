"""Parquet storage for normalized CONAB crop progress rows."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / "data" / "processed" / "conab_progress.parquet"
COLUMNS = ["bulletin_date", "crop", "state", "season", "planting_pct", "harvest_pct"]
KEY_COLUMNS = ["bulletin_date", "crop", "state", "season"]


def _empty_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return _empty_frame()

    rows = df.copy()
    for column in COLUMNS:
        if column not in rows.columns:
            rows[column] = pd.NA
    rows = rows[COLUMNS]
    rows["bulletin_date"] = pd.to_datetime(rows["bulletin_date"])
    rows["planting_pct"] = pd.to_numeric(rows["planting_pct"], errors="coerce")
    rows["harvest_pct"] = pd.to_numeric(rows["harvest_pct"], errors="coerce")
    return rows.drop_duplicates(KEY_COLUMNS, keep="last").sort_values(KEY_COLUMNS).reset_index(drop=True)


def load_data(data_path: str | Path = DATA_PATH) -> pd.DataFrame:
    """Load all stored CONAB progress rows."""

    path = Path(data_path)
    if not path.exists():
        return _empty_frame()

    return _normalize(pd.read_parquet(path))


def latest_bulletin_date(data_path: str | Path = DATA_PATH) -> pd.Timestamp | None:
    """Return the newest stored bulletin date."""

    df = load_data(data_path)
    if df.empty:
        return None
    return pd.to_datetime(df["bulletin_date"]).max()


def upsert_data(df: pd.DataFrame, data_path: str | Path = DATA_PATH) -> int:
    """Upsert normalized rows and return the number of new unique keys added."""

    if df.empty:
        return 0

    path = Path(data_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    existing = load_data(path)
    before_keys = set(existing[KEY_COLUMNS].itertuples(index=False, name=None))
    combined = _normalize(pd.concat([existing, df], ignore_index=True))
    after_keys = set(combined[KEY_COLUMNS].itertuples(index=False, name=None))
    combined.to_parquet(path, index=False)
    return len(after_keys - before_keys)
