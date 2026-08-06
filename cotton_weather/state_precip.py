from __future__ import annotations

from datetime import date, datetime, timedelta
import json
from pathlib import Path

import numpy as np
import pandas as pd
import rasterio
from rasterio.warp import Resampling, reproject

from cotton_weather.cdl import load_cdl_footprint_summary, summarize_downloaded_cdl_rasters
from cotton_weather.config import (
    CDL_COTTON_CLASS_CODE,
    CDL_RAW_DIR,
    COTTON_STATE_FIPS,
    DEFAULT_HISTORY_DAYS,
    DEFAULT_REPROCESS_DAYS,
    RAW_DIR,
    STATE_PRECIP_BACKFILL_PLAN_FILE,
    STATE_PRECIP_FILE,
    STATE_PRECIP_METADATA_FILE,
    STATE_PRECIP_PROGRESS_FILE,
    STATE_WEIGHT_DIR,
)
from cotton_weather.prism import PrismDownloadError, ensure_prism_asset


def _date_range(start_date: date, end_date: date) -> list[date]:
    return [start_date + timedelta(days=offset) for offset in range((end_date - start_date).days + 1)]


def _chunk_ranges(start_date: date, end_date: date, chunk_days: int) -> list[tuple[date, date]]:
    ranges: list[tuple[date, date]] = []
    cursor = start_date
    while cursor <= end_date:
        chunk_end = min(cursor + timedelta(days=chunk_days - 1), end_date)
        ranges.append((cursor, chunk_end))
        cursor = chunk_end + timedelta(days=1)
    return ranges


def _chunk_ranges_reverse(start_date: date, end_date: date, chunk_days: int) -> list[tuple[date, date]]:
    ranges: list[tuple[date, date]] = []
    cursor = end_date
    while cursor >= start_date:
        chunk_start = max(cursor - timedelta(days=chunk_days - 1), start_date)
        ranges.append((chunk_start, cursor))
        cursor = chunk_start - timedelta(days=1)
    return ranges


def _load_existing_state_dataset() -> pd.DataFrame:
    if not STATE_PRECIP_FILE.exists():
        return pd.DataFrame()
    dataset = pd.read_parquet(STATE_PRECIP_FILE)
    dataset["date"] = pd.to_datetime(dataset["date"])
    return dataset


def _write_state_dataset(dataset: pd.DataFrame) -> None:
    STATE_PRECIP_FILE.parent.mkdir(parents=True, exist_ok=True)
    dataset.to_parquet(STATE_PRECIP_FILE, index=False)


def _merge_state_rows(existing: pd.DataFrame, day_rows: pd.DataFrame) -> pd.DataFrame:
    if existing.empty:
        return day_rows.sort_values(["date", "state"]).reset_index(drop=True)
    replacement_keys = set(zip(day_rows["date"].dt.date, day_rows["state"]))
    retained = existing.loc[
        ~existing.apply(lambda row: (row["date"].date(), row["state"]) in replacement_keys, axis=1)
    ].copy()
    merged = pd.concat([retained, day_rows], ignore_index=True)
    return merged.sort_values(["date", "state"]).reset_index(drop=True)


def _write_progress(progress: dict) -> None:
    STATE_PRECIP_PROGRESS_FILE.write_text(json.dumps(progress, indent=2), encoding="utf-8")


def cleanup_state_precip_raw_files(start_date: date, end_date: date) -> dict:
    deleted = {"zip_files": 0, "tif_files": 0, "day_dirs": 0, "bytes_freed": 0}
    for day in _date_range(start_date, end_date):
        year_dir = RAW_DIR / "ppt" / day.strftime("%Y")
        zip_path = year_dir / f"prism_ppt_us_30s_{day.strftime('%Y%m%d')}.zip"
        tif_dir = year_dir / day.strftime("%Y%m%d")
        tif_path = tif_dir / f"prism_ppt_us_30s_{day.strftime('%Y%m%d')}.tif"

        for path, key in ((tif_path, "tif_files"), (zip_path, "zip_files")):
            if path.exists():
                deleted["bytes_freed"] += path.stat().st_size
                path.unlink()
                deleted[key] += 1

        if tif_dir.exists():
            try:
                tif_dir.rmdir()
                deleted["day_dirs"] += 1
            except OSError:
                pass
    return deleted


def _build_metadata(
    final: pd.DataFrame,
    resolved_start: date,
    resolved_end: date,
    selected_states: list[str],
    footprint_year: int,
    missing_assets: list[str],
) -> dict:
    return {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "requested_start": resolved_start.isoformat(),
        "requested_end": resolved_end.isoformat(),
        "rows_written": int(len(final)),
        "dates_processed": len(_date_range(resolved_start, resolved_end)),
        "states": selected_states,
        "footprint_year": footprint_year,
        "missing_assets": missing_assets,
        "dataset_start": final["date"].min().date().isoformat() if not final.empty else None,
        "dataset_end": final["date"].max().date().isoformat() if not final.empty else None,
    }


def _determine_start_date(
    existing: pd.DataFrame,
    history_days: int,
    reprocess_days: int,
    end_date: date,
) -> date:
    if existing.empty:
        return end_date - timedelta(days=history_days - 1)
    latest_existing = existing["date"].max().date()
    return min(latest_existing - timedelta(days=reprocess_days - 1), end_date)


def _normalize_states(states: list[str] | None) -> list[str]:
    if not states:
        return sorted(COTTON_STATE_FIPS.keys())
    normalized = [state.strip().upper() for state in states]
    invalid = sorted(set(normalized).difference(COTTON_STATE_FIPS))
    if invalid:
        invalid_text = ", ".join(invalid)
        raise ValueError(f"Unsupported cotton state abbreviations: {invalid_text}")
    return normalized


def _cdl_raster_path(state_abbr: str, footprint_year: int) -> Path:
    return CDL_RAW_DIR / str(footprint_year) / f"{state_abbr.lower()}_{footprint_year}_cdl.tif"


def _state_weight_path(state_abbr: str, footprint_year: int) -> Path:
    return STATE_WEIGHT_DIR / str(footprint_year) / f"{state_abbr.lower()}_prism_weights.tif"


def ensure_state_prism_weights(
    state_abbr: str,
    prism_tif_path: Path,
    footprint_year: int = 2024,
) -> Path:
    weight_path = _state_weight_path(state_abbr, footprint_year)
    if weight_path.exists():
        return weight_path

    cdl_path = _cdl_raster_path(state_abbr, footprint_year)
    if not cdl_path.exists():
        raise FileNotFoundError(f"Missing CDL raster for {state_abbr}: {cdl_path}")

    STATE_WEIGHT_DIR.mkdir(parents=True, exist_ok=True)
    weight_path.parent.mkdir(parents=True, exist_ok=True)

    with rasterio.open(cdl_path) as cdl_ds, rasterio.open(prism_tif_path) as prism_ds:
        cdl_band = cdl_ds.read(1, masked=True)
        source_mask = np.where(
            (cdl_band == CDL_COTTON_CLASS_CODE) & ~cdl_band.mask,
            1.0,
            0.0,
        ).astype("float32")
        destination = np.zeros((prism_ds.height, prism_ds.width), dtype="float32")
        reproject(
            source=source_mask,
            destination=destination,
            src_transform=cdl_ds.transform,
            src_crs=cdl_ds.crs,
            src_nodata=0.0,
            dst_transform=prism_ds.transform,
            dst_crs=prism_ds.crs,
            dst_nodata=0.0,
            resampling=Resampling.average,
        )
        metadata = prism_ds.meta.copy()
        metadata.update(dtype="float32", count=1, compress="lzw", nodata=0.0)
        with rasterio.open(weight_path, "w", **metadata) as output:
            output.write(destination, 1)

    return weight_path


def aggregate_state_precipitation_for_day(
    date_value: date,
    footprint_year: int = 2024,
    states: list[str] | None = None,
) -> pd.DataFrame:
    selected_states = _normalize_states(states)
    prism_asset = ensure_prism_asset(variable="ppt", date_value=date_value, raw_dir=RAW_DIR)
    state_area = summarize_downloaded_cdl_rasters(year=footprint_year)
    if state_area.empty:
        state_area = load_cdl_footprint_summary()
    area_lookup = state_area.set_index("state")["cotton_area_acres_est"].to_dict() if not state_area.empty else {}

    rows: list[dict] = []
    with rasterio.open(prism_asset.tif_path) as prism_ds:
        ppt_band = prism_ds.read(1, masked=True)
        for state_abbr in selected_states:
            weight_path = _state_weight_path(state_abbr, footprint_year)
            if not weight_path.exists() and not _cdl_raster_path(state_abbr, footprint_year).exists():
                continue
            if not weight_path.exists():
                weight_path = ensure_state_prism_weights(
                    state_abbr=state_abbr,
                    prism_tif_path=prism_asset.tif_path,
                    footprint_year=footprint_year,
                )
            with rasterio.open(weight_path) as weight_ds:
                weights = weight_ds.read(1, masked=True)

            valid_mask = (~ppt_band.mask) & (weights.data > 0)
            if not np.any(valid_mask):
                weighted_mean = None
                wet_share = None
                weight_sum = 0.0
            else:
                valid_weights = weights.data[valid_mask]
                valid_ppt = ppt_band.data[valid_mask]
                weight_sum = float(valid_weights.sum())
                weighted_mean = float(np.average(valid_ppt, weights=valid_weights))
                wet_share = float(valid_weights[valid_ppt > 0].sum() / weight_sum) if weight_sum > 0 else None

            rows.append(
                {
                    "date": pd.to_datetime(date_value),
                    "state": state_abbr,
                    "ppt_mm": weighted_mean,
                    "wet_cell_share": wet_share,
                    "prism_weight_sum": weight_sum,
                    "cotton_area_acres_est": area_lookup.get(state_abbr, weight_sum if weight_sum > 0 else None),
                }
            )

    return pd.DataFrame(rows)


def update_state_precipitation_dataset(
    history_days: int = DEFAULT_HISTORY_DAYS,
    reprocess_days: int = DEFAULT_REPROCESS_DAYS,
    start_date: date | None = None,
    end_date: date | None = None,
    footprint_year: int = 2024,
    states: list[str] | None = None,
    progress_callback=None,
) -> dict:
    selected_states = _normalize_states(states)
    existing = _load_existing_state_dataset()
    resolved_end = end_date or (date.today() - timedelta(days=1))
    resolved_start = start_date or _determine_start_date(
        existing=existing,
        history_days=history_days,
        reprocess_days=reprocess_days,
        end_date=resolved_end,
    )
    if resolved_start > resolved_end:
        raise ValueError("start_date cannot be after end_date")

    total_days = len(_date_range(resolved_start, resolved_end))
    missing_assets: list[str] = []
    current = existing.copy()
    for index, day in enumerate(_date_range(resolved_start, resolved_end), start=1):
        if progress_callback:
            progress_callback(
                {
                    "phase": "running",
                    "current_day": index,
                    "total_days": total_days,
                    "date": day.isoformat(),
                    "states": selected_states,
                }
            )
        try:
            day_rows = aggregate_state_precipitation_for_day(
                date_value=day,
                footprint_year=footprint_year,
                states=selected_states,
            )
            current = _merge_state_rows(current, day_rows)
            _write_state_dataset(current)
            _write_progress(
                {
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                    "phase": "running",
                    "current_day": index,
                    "total_days": total_days,
                    "date": day.isoformat(),
                    "states": selected_states,
                    "footprint_year": footprint_year,
                }
            )
        except PrismDownloadError:
            missing_assets.append(day.isoformat())
            _write_progress(
                {
                    "updated_at": datetime.now().isoformat(timespec="seconds"),
                    "phase": "missing_asset",
                    "current_day": index,
                    "total_days": total_days,
                    "date": day.isoformat(),
                    "states": selected_states,
                    "footprint_year": footprint_year,
                }
            )

    final = current.sort_values(["date", "state"]).reset_index(drop=True)
    _write_state_dataset(final)

    metadata = _build_metadata(
        final=final,
        resolved_start=resolved_start,
        resolved_end=resolved_end,
        selected_states=selected_states,
        footprint_year=footprint_year,
        missing_assets=missing_assets,
    )
    STATE_PRECIP_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    _write_progress(
        {
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "phase": "completed",
            "current_day": total_days,
            "total_days": total_days,
            "date": resolved_end.isoformat(),
            "states": selected_states,
            "footprint_year": footprint_year,
            "rows_written": metadata["rows_written"],
        }
    )
    if progress_callback:
        progress_callback(
            {
                "phase": "completed",
                "current_day": total_days,
                "total_days": total_days,
                "date": resolved_end.isoformat(),
                "states": selected_states,
            }
        )
    return metadata


def load_state_precipitation_data() -> pd.DataFrame:
    return _load_existing_state_dataset()


def load_state_precipitation_metadata() -> dict:
    if not STATE_PRECIP_METADATA_FILE.exists():
        return {}
    return json.loads(STATE_PRECIP_METADATA_FILE.read_text(encoding="utf-8"))


def load_state_precipitation_progress() -> dict:
    if not STATE_PRECIP_PROGRESS_FILE.exists():
        return {}
    return json.loads(STATE_PRECIP_PROGRESS_FILE.read_text(encoding="utf-8"))


def create_state_precip_backfill_plan(
    start_date: date,
    end_date: date,
    chunk_days: int = 90,
    footprint_year: int = 2024,
    states: list[str] | None = None,
    newest_first: bool = False,
) -> dict:
    if start_date > end_date:
        raise ValueError("start_date cannot be after end_date")
    if chunk_days <= 0:
        raise ValueError("chunk_days must be positive")

    selected_states = _normalize_states(states)
    chunk_ranges = _chunk_ranges_reverse(start_date, end_date, chunk_days) if newest_first else _chunk_ranges(start_date, end_date, chunk_days)
    chunks = [
        {
            "index": index,
            "start_date": chunk_start.isoformat(),
            "end_date": chunk_end.isoformat(),
            "status": "pending",
        }
        for index, (chunk_start, chunk_end) in enumerate(chunk_ranges, start=1)
    ]
    plan = {
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "start_date": start_date.isoformat(),
        "end_date": end_date.isoformat(),
        "chunk_days": chunk_days,
        "chunk_count": len(chunks),
        "completed_chunks": 0,
        "footprint_year": footprint_year,
        "states": selected_states,
        "newest_first": newest_first,
        "chunks": chunks,
    }
    STATE_PRECIP_BACKFILL_PLAN_FILE.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    return plan


def load_state_precip_backfill_plan() -> dict:
    if not STATE_PRECIP_BACKFILL_PLAN_FILE.exists():
        return {}
    return json.loads(STATE_PRECIP_BACKFILL_PLAN_FILE.read_text(encoding="utf-8"))


def update_state_precip_backfill_chunk(
    chunk_index: int,
    status: str,
    result: dict | None = None,
) -> dict:
    plan = load_state_precip_backfill_plan()
    if not plan:
        raise ValueError("No state precipitation backfill plan found.")

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
        raise ValueError(f"Chunk index {chunk_index} not found in state precipitation backfill plan.")

    plan["completed_chunks"] = sum(1 for chunk in plan["chunks"] if chunk["status"] == "completed")
    STATE_PRECIP_BACKFILL_PLAN_FILE.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    return plan


def run_state_precip_backfill_chunk(chunk_index: int) -> dict:
    return run_state_precip_backfill_chunk_with_progress(chunk_index=chunk_index, progress_callback=None)


def run_state_precip_backfill_chunk_with_progress(
    chunk_index: int,
    progress_callback=None,
    cleanup_raw: bool = False,
) -> dict:
    plan = load_state_precip_backfill_plan()
    if not plan:
        raise ValueError("No state precipitation backfill plan found.")

    chunk = next((item for item in plan["chunks"] if item["index"] == chunk_index), None)
    if chunk is None:
        raise ValueError(f"Chunk index {chunk_index} not found in state precipitation backfill plan.")

    update_state_precip_backfill_chunk(chunk_index, "running")
    try:
        result = update_state_precipitation_dataset(
            start_date=date.fromisoformat(chunk["start_date"]),
            end_date=date.fromisoformat(chunk["end_date"]),
            footprint_year=plan["footprint_year"],
            states=plan["states"],
            progress_callback=progress_callback,
        )
    except Exception:
        update_state_precip_backfill_chunk(chunk_index, "failed")
        raise

    cleanup_result = None
    if cleanup_raw:
        cleanup_result = cleanup_state_precip_raw_files(
            start_date=date.fromisoformat(chunk["start_date"]),
            end_date=date.fromisoformat(chunk["end_date"]),
        )
        result["cleanup"] = cleanup_result

    update_state_precip_backfill_chunk(chunk_index, "completed", result=result)
    return result
