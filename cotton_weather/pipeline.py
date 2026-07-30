from __future__ import annotations

from datetime import date, datetime, timedelta
import json

import pandas as pd

from cotton_weather.config import (
    BACKFILL_PLAN_FILE,
    DEFAULT_HISTORY_DAYS,
    DEFAULT_BACKFILL_CHUNK_DAYS,
    DEFAULT_REPROCESS_DAYS,
    METADATA_FILE,
    PROCESSED_DATA_FILE,
    PROCESSED_DIR,
    RAW_DIR,
    SUPPORTED_VARIABLES,
)
from cotton_weather.data import load_locations
from cotton_weather.prism import PrismDownloadError, ensure_prism_asset, sample_points


def _date_range(start_date: date, end_date: date) -> list[date]:
    day_count = (end_date - start_date).days + 1
    return [start_date + timedelta(days=offset) for offset in range(day_count)]


def _chunk_ranges(start_date: date, end_date: date, chunk_days: int) -> list[tuple[date, date]]:
    ranges: list[tuple[date, date]] = []
    cursor = start_date
    while cursor <= end_date:
        chunk_end = min(cursor + timedelta(days=chunk_days - 1), end_date)
        ranges.append((cursor, chunk_end))
        cursor = chunk_end + timedelta(days=1)
    return ranges


def _load_existing_dataset() -> pd.DataFrame:
    if not PROCESSED_DATA_FILE.exists():
        return pd.DataFrame()
    dataset = pd.read_parquet(PROCESSED_DATA_FILE)
    dataset["date"] = pd.to_datetime(dataset["date"])
    return dataset


def _determine_start_date(
    existing: pd.DataFrame,
    history_days: int,
    reprocess_days: int,
    end_date: date,
) -> date:
    if existing.empty:
        return end_date - timedelta(days=history_days - 1)
    latest_existing = existing["date"].max().date()
    return min(
        latest_existing - timedelta(days=reprocess_days - 1),
        end_date,
    )


def update_weather_dataset(
    history_days: int = DEFAULT_HISTORY_DAYS,
    reprocess_days: int = DEFAULT_REPROCESS_DAYS,
    start_date: date | None = None,
    end_date: date | None = None,
    metadata_context: dict | None = None,
) -> dict:
    locations = load_locations()
    existing = _load_existing_dataset()

    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)

    resolved_end = end_date or (date.today() - timedelta(days=1))
    resolved_start = start_date or _determine_start_date(
        existing=existing,
        history_days=history_days,
        reprocess_days=reprocess_days,
        end_date=resolved_end,
    )

    if resolved_start > resolved_end:
        raise ValueError("start_date cannot be after end_date")

    daily_frames: list[pd.DataFrame] = []
    missing_assets: list[str] = []
    source_urls: dict[str, str] = {}

    for day in _date_range(resolved_start, resolved_end):
        merged = locations.copy()
        merged["date"] = pd.to_datetime(day)

        for variable in SUPPORTED_VARIABLES:
            try:
                asset = ensure_prism_asset(variable=variable, date_value=day, raw_dir=RAW_DIR)
                source_urls[f"{variable}_{day.isoformat()}"] = asset.source_url
                sampled = sample_points(asset, locations)
            except PrismDownloadError:
                missing_assets.append(f"{variable}:{day.isoformat()}")
                sampled = pd.DataFrame(
                    {
                        "location_id": locations["location_id"].values,
                        "date": pd.to_datetime(day),
                        variable: [None] * len(locations),
                    }
                )
            merged = merged.merge(sampled, on=["location_id", "date"], how="left")

        merged = merged.rename(
            columns={
                "ppt": "ppt_mm",
                "tmin": "tmin_c",
                "tmax": "tmax_c",
            }
        )
        merged["tmean_c"] = (merged["tmin_c"] + merged["tmax_c"]) / 2.0
        daily_frames.append(merged)

    updated = pd.concat(daily_frames, ignore_index=True) if daily_frames else existing.copy()

    if not existing.empty:
        updated_dates = set(updated["date"].dt.date)
        retained = existing.loc[
            ~existing["date"].dt.date.isin(updated_dates)
        ].copy()
        final = pd.concat([retained, updated], ignore_index=True)
    else:
        final = updated

    final = final.sort_values(["date", "state", "location_name"]).reset_index(drop=True)
    final.to_parquet(PROCESSED_DATA_FILE, index=False)

    metadata = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "requested_start": resolved_start.isoformat(),
        "requested_end": resolved_end.isoformat(),
        "rows_written": int(len(final)),
        "dates_processed": len(_date_range(resolved_start, resolved_end)),
        "missing_assets": missing_assets,
        "source_urls": source_urls,
        "dataset_start": final["date"].min().date().isoformat() if not final.empty else None,
        "dataset_end": final["date"].max().date().isoformat() if not final.empty else None,
        "variables": list(SUPPORTED_VARIABLES),
        "location_count": int(final["location_id"].nunique()) if not final.empty else 0,
    }
    if metadata_context:
        metadata["context"] = metadata_context
    METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return metadata


def refresh_recent_weather(
    history_days: int = DEFAULT_HISTORY_DAYS,
    reprocess_days: int = DEFAULT_REPROCESS_DAYS,
) -> dict:
    return update_weather_dataset(
        history_days=history_days,
        reprocess_days=reprocess_days,
        metadata_context={
            "job_type": "recent_refresh",
            "history_days": history_days,
            "reprocess_days": reprocess_days,
        },
    )


def create_backfill_plan(
    start_date: date,
    end_date: date,
    chunk_days: int = DEFAULT_BACKFILL_CHUNK_DAYS,
) -> dict:
    if start_date > end_date:
        raise ValueError("start_date cannot be after end_date")
    if chunk_days <= 0:
        raise ValueError("chunk_days must be positive")

    chunks = [
        {
            "index": index,
            "start_date": chunk_start.isoformat(),
            "end_date": chunk_end.isoformat(),
            "status": "pending",
        }
        for index, (chunk_start, chunk_end) in enumerate(
            _chunk_ranges(start_date, end_date, chunk_days),
            start=1,
        )
    ]
    plan = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "chunk_days": chunk_days,
        "chunk_count": len(chunks),
        "completed_chunks": 0,
        "chunks": chunks,
    }
    BACKFILL_PLAN_FILE.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    return plan


def load_backfill_plan() -> dict:
    if not BACKFILL_PLAN_FILE.exists():
        return {}
    return json.loads(BACKFILL_PLAN_FILE.read_text(encoding="utf-8"))


def run_backfill_chunk(
    start_date: date,
    end_date: date,
    chunk_index: int,
    total_chunks: int,
) -> dict:
    return update_weather_dataset(
        start_date=start_date,
        end_date=end_date,
        metadata_context={
            "job_type": "historical_backfill_chunk",
            "chunk_index": chunk_index,
            "total_chunks": total_chunks,
        },
    )


def update_backfill_plan_chunk(
    chunk_index: int,
    status: str,
    result: dict | None = None,
) -> dict:
    plan = load_backfill_plan()
    if not plan:
        raise ValueError("No backfill plan found.")

    for chunk in plan["chunks"]:
        if chunk["index"] == chunk_index:
            chunk["status"] = status
            chunk["updated_at"] = datetime.now().isoformat(timespec="seconds")
            if result:
                chunk["result"] = {
                    "requested_start": result["requested_start"],
                    "requested_end": result["requested_end"],
                    "dates_processed": result["dates_processed"],
                    "missing_assets": result["missing_assets"],
                }
            break
    else:
        raise ValueError(f"Chunk index {chunk_index} not found in plan.")

    plan["completed_chunks"] = sum(
        1 for chunk in plan["chunks"] if chunk["status"] == "completed"
    )
    BACKFILL_PLAN_FILE.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    return plan
