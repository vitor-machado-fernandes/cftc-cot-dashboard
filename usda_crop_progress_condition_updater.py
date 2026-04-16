from __future__ import annotations

import os
from pathlib import Path

import pandas as pd
import requests

API_URL = "https://quickstats.nass.usda.gov/api/api_GET/"
DATA_FILENAME = "usda_crop_progress_condition.parquet"
DATE_COL = "week_ending"

COMMODITIES = {
    "Corn": {"commodity_desc": "CORN"},
    "Soybeans": {"commodity_desc": "SOYBEANS"},
    "Cotton": {"commodity_desc": "COTTON"},
    "Winter Wheat": {"commodity_desc": "WHEAT", "class_desc": "WINTER"},
}

SENTINEL_COMMODITY = "CORN"


def _get_api_key(api_key: str | None = None) -> str | None:
    if api_key:
        return api_key

    env_key = (
        os.getenv("USDA_QUICKSTATS_API_KEY")
        or os.getenv("QUICKSTATS_API_KEY")
        or os.getenv("NASS_API_KEY")
    )
    if env_key:
        return env_key

    return None


def _clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("(D)", "", regex=False)
        .str.replace("(Z)", "0", regex=False)
        .str.strip()
        .replace({"": None, "nan": None, "None": None}),
        errors="coerce",
    )


def _fetch_quickstats_rows(api_key: str, params: dict) -> pd.DataFrame:
    response = requests.get(
        API_URL,
        params={**params, "key": api_key, "format": "JSON"},
        timeout=120,
    )
    response.raise_for_status()
    payload = response.json()

    if "error" in payload:
        raise RuntimeError(payload["error"])

    rows = payload.get("data", [])
    return pd.DataFrame(rows)


def _fetch_latest_remote_date(api_key: str) -> pd.Timestamp | None:
    df = _fetch_quickstats_rows(
        api_key,
        {
            "source_desc": "SURVEY",
            "sector_desc": "CROPS",
            "group_desc": "FIELD CROPS",
            "commodity_desc": SENTINEL_COMMODITY,
            "statisticcat_desc": "PROGRESS",
            "agg_level_desc": "NATIONAL",
            "freq_desc": "WEEKLY",
            "year__GE": 2020,
        },
    )
    if df.empty or DATE_COL not in df.columns:
        return None

    latest = pd.to_datetime(df[DATE_COL], errors="coerce")
    if latest.isna().all():
        return None

    return latest.max().normalize()


def _fetch_commodity_history(api_key: str, crop: str, crop_params: dict, start_year: int = 2010) -> pd.DataFrame:
    return fetch_crop_progress_condition_history(
        api_key=api_key,
        crop=crop,
        agg_level_desc="NATIONAL",
        state_name=None,
        start_year=start_year,
    )


def fetch_crop_progress_condition_history(
    api_key: str,
    crop: str,
    agg_level_desc: str = "NATIONAL",
    state_name: str | None = None,
    start_year: int = 2010,
) -> pd.DataFrame:
    crop_params = COMMODITIES[crop]
    params = {
        "source_desc": "SURVEY",
        "sector_desc": "CROPS",
        "group_desc": "FIELD CROPS",
        "agg_level_desc": agg_level_desc,
        "freq_desc": "WEEKLY",
        "year__GE": start_year,
        **crop_params,
    }
    if state_name:
        params["state_name"] = state_name

    df = _fetch_quickstats_rows(
        api_key,
        params,
    )

    if df.empty:
        return df

    df = df[df["statisticcat_desc"].isin(["PROGRESS", "CONDITION"])].copy()
    if df.empty:
        return df

    df["crop"] = crop
    df["report_date"] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df["value"] = _clean_numeric(df["Value"])
    df["week_of_year"] = df["report_date"].dt.isocalendar().week.astype("Int64")
    df["year"] = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    df["class_desc"] = df["class_desc"].fillna("").str.strip()
    df["short_desc"] = df["short_desc"].fillna("").str.strip()
    df["unit_desc"] = df["unit_desc"].fillna("").str.strip()
    df["reference_period_desc"] = df["reference_period_desc"].fillna("").str.strip()
    df["agg_level_desc"] = df["agg_level_desc"].fillna("").str.strip()
    df["state_name"] = df["state_name"].fillna("").str.strip()
    df["location_label"] = df["state_name"].where(df["state_name"].ne(""), "National")

    keep_cols = [
        "crop",
        "commodity_desc",
        "statisticcat_desc",
        "class_desc",
        "unit_desc",
        "short_desc",
        "reference_period_desc",
        "agg_level_desc",
        "state_name",
        "location_label",
        "report_date",
        "week_of_year",
        "year",
        "value",
        "Value",
        "load_time",
    ]
    return df[keep_cols].dropna(subset=["report_date"]).sort_values(
        ["crop", "statisticcat_desc", "report_date", "class_desc"]
    )


def _load_local_latest_date(path: Path) -> pd.Timestamp | None:
    if not path.exists():
        return None

    df = pd.read_parquet(path, columns=["report_date"])
    if df.empty:
        return None

    latest = pd.to_datetime(df["report_date"], errors="coerce")
    if latest.isna().all():
        return None

    return latest.max().normalize()


def refresh_crop_progress_condition_data(
    data_dir: str | Path = ".",
    api_key: str | None = None,
    force: bool = False,
) -> dict:
    data_dir = Path(data_dir)
    path = data_dir / DATA_FILENAME
    resolved_key = _get_api_key(api_key)

    if not resolved_key:
        return {
            "did_update": False,
            "skipped": True,
            "reason": "missing_api_key",
            "path": str(path),
            "local_latest": _load_local_latest_date(path),
            "remote_latest": None,
            "messages": [],
        }

    remote_latest = _fetch_latest_remote_date(resolved_key)
    local_latest = _load_local_latest_date(path)

    if not force and local_latest is not None and remote_latest is not None and local_latest >= remote_latest:
        return {
            "did_update": False,
            "skipped": False,
            "reason": None,
            "path": str(path),
            "local_latest": local_latest,
            "remote_latest": remote_latest,
            "messages": [],
        }

    frames: list[pd.DataFrame] = []
    messages: list[str] = []

    for crop, crop_params in COMMODITIES.items():
        crop_df = _fetch_commodity_history(resolved_key, crop, crop_params)
        if crop_df.empty:
            messages.append(f"{crop}: no weekly progress/condition rows returned.")
            continue

        frames.append(crop_df)
        crop_min = crop_df["report_date"].min().date()
        crop_max = crop_df["report_date"].max().date()
        messages.append(f"{crop}: {len(crop_df):,} rows ({crop_min} -> {crop_max})")

    if not frames:
        raise RuntimeError("USDA Quick Stats returned no crop progress/condition rows for the configured crops.")

    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(
        subset=["crop", "agg_level_desc", "state_name", "statisticcat_desc", "class_desc", "report_date", "short_desc"],
        keep="last",
    ).sort_values(["crop", "statisticcat_desc", "report_date", "class_desc"])

    out.to_parquet(path, index=False)

    return {
        "did_update": True,
        "skipped": False,
        "reason": None,
        "path": str(path),
        "local_latest": local_latest,
        "remote_latest": remote_latest,
        "messages": messages,
    }
