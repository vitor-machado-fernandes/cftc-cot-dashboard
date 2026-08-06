from __future__ import annotations

from datetime import date, timedelta
import json
from pathlib import Path

import numpy as np
import pandas as pd

from cotton_weather.config import PROCESSED_DIR, RAW_DIR


PRISM_MAP_PREVIEW_CACHE_DIR = PROCESSED_DIR / "prism_map_previews"


def list_available_prism_dates(variable: str = "ppt") -> list[date]:
    variable_dir = RAW_DIR / variable
    if not variable_dir.exists():
        return []

    dates: list[date] = []
    for tif_path in variable_dir.glob("*/*/prism_*.tif"):
        try:
            dates.append(date.fromisoformat(tif_path.parent.name))
        except ValueError:
            continue
    return sorted(set(dates))


def _build_window_dates(end_date: date, window_days: int) -> list[date]:
    return [end_date - timedelta(days=offset) for offset in range(window_days - 1, -1, -1)]


def _resolve_tif_path(date_value: date, variable: str = "ppt") -> Path:
    token = date_value.strftime("%Y%m%d")
    return RAW_DIR / variable / date_value.strftime("%Y") / token / f"prism_{variable}_us_30s_{token}.tif"


def available_map_windows(window_days: int, variable: str = "ppt") -> list[date]:
    available_dates = set(list_available_prism_dates(variable=variable))
    valid_end_dates: list[date] = []
    for candidate in sorted(available_dates):
        required = _build_window_dates(candidate, window_days)
        if all(day in available_dates for day in required):
            valid_end_dates.append(candidate)
    return valid_end_dates


def _cache_paths(
    window_days: int,
    end_date: date,
    downsample_factor: int,
    min_precip_mm: float,
) -> tuple[Path, Path]:
    end_token = end_date.isoformat()
    min_token = str(float(min_precip_mm)).replace(".", "p")
    stem = f"prism_ppt_{window_days}day_{end_token}_ds{downsample_factor}_min{min_token}"
    return (
        PRISM_MAP_PREVIEW_CACHE_DIR / f"{stem}.parquet",
        PRISM_MAP_PREVIEW_CACHE_DIR / f"{stem}.json",
    )


def available_cached_map_windows(
    window_days: int,
    downsample_factor: int = 28,
    min_precip_mm: float = 0.1,
) -> list[date]:
    if not PRISM_MAP_PREVIEW_CACHE_DIR.exists():
        return []

    suffix = f"_ds{downsample_factor}_min{str(float(min_precip_mm)).replace('.', 'p')}.parquet"
    prefix = f"prism_ppt_{window_days}day_"
    end_dates: list[date] = []
    for path in PRISM_MAP_PREVIEW_CACHE_DIR.glob(f"{prefix}*{suffix}"):
        date_token = path.name.removeprefix(prefix).removesuffix(suffix)
        try:
            end_dates.append(date.fromisoformat(date_token))
        except ValueError:
            continue
    return sorted(set(end_dates))


def load_precipitation_map_preview_cached(
    window_days: int,
    end_date: date,
    downsample_factor: int = 28,
    min_precip_mm: float = 0.1,
    build_if_missing: bool = True,
) -> tuple[pd.DataFrame, dict]:
    preview_path, meta_path = _cache_paths(
        window_days=window_days,
        end_date=end_date,
        downsample_factor=downsample_factor,
        min_precip_mm=min_precip_mm,
    )
    if preview_path.exists() and meta_path.exists():
        try:
            return pd.read_parquet(preview_path), json.loads(meta_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            if not build_if_missing:
                return pd.DataFrame(), {
                    "window_days": window_days,
                    "available": False,
                    "end_date": end_date.isoformat(),
                    "cache_file": str(preview_path),
                }

    if not build_if_missing:
        return pd.DataFrame(), {
            "window_days": window_days,
            "available": False,
            "end_date": end_date.isoformat(),
            "cache_file": str(preview_path),
        }

    preview, metadata = build_precipitation_map_preview(
        window_days=window_days,
        end_date=end_date,
        downsample_factor=downsample_factor,
        min_precip_mm=min_precip_mm,
    )
    if not preview.empty:
        PRISM_MAP_PREVIEW_CACHE_DIR.mkdir(parents=True, exist_ok=True)
        preview.to_parquet(preview_path, index=False)
        metadata["cache_file"] = str(preview_path)
        meta_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return preview, metadata


def prebuild_precipitation_map_previews(
    window_days_options: tuple[int, ...] = (1, 2, 7, 14),
    recent_end_dates: int = 7,
    downsample_factor: int = 28,
    min_precip_mm: float = 0.1,
) -> list[dict]:
    built: list[dict] = []
    for window_days in window_days_options:
        end_dates = available_map_windows(window_days=window_days, variable="ppt")
        for end_date in end_dates[-recent_end_dates:]:
            _, metadata = load_precipitation_map_preview_cached(
                window_days=window_days,
                end_date=end_date,
                downsample_factor=downsample_factor,
                min_precip_mm=min_precip_mm,
                build_if_missing=True,
            )
            built.append(metadata)
    return built


def build_precipitation_map_preview(
    window_days: int,
    end_date: date | None = None,
    downsample_factor: int = 28,
    min_precip_mm: float = 0.1,
) -> tuple[pd.DataFrame, dict]:
    try:
        from affine import Affine
        import rasterio
        from rasterio.enums import Resampling
        from rasterio.transform import xy
        from rasterio.warp import transform
    except ModuleNotFoundError as exc:
        missing_module = getattr(exc, "name", None) or "a raster-map dependency"
        raise ModuleNotFoundError(
            f"The '{missing_module}' package is required to render PRISM precipitation maps. "
            "Install the CoT requirements or run the app from the project .conda environment."
        ) from exc

    valid_end_dates = available_map_windows(window_days=window_days, variable="ppt")
    if not valid_end_dates:
        return pd.DataFrame(), {"window_days": window_days, "available": False}

    resolved_end = end_date or valid_end_dates[-1]
    if resolved_end not in valid_end_dates:
        raise ValueError(f"No complete {window_days}-day precipitation window ends on {resolved_end.isoformat()}.")

    window_dates = _build_window_dates(resolved_end, window_days)
    sample_path = _resolve_tif_path(window_dates[-1], variable="ppt")
    with rasterio.open(sample_path) as sample_ds:
        out_height = max(1, sample_ds.height // downsample_factor)
        out_width = max(1, sample_ds.width // downsample_factor)
        preview_transform = sample_ds.transform * Affine.scale(
            sample_ds.width / out_width,
            sample_ds.height / out_height,
        )
        preview_crs = sample_ds.crs

    accumulated = np.zeros((out_height, out_width), dtype="float32")
    valid_cells = np.zeros((out_height, out_width), dtype=bool)
    for day in window_dates:
        tif_path = _resolve_tif_path(day, variable="ppt")
        with rasterio.open(tif_path) as dataset:
            data = dataset.read(
                1,
                out_shape=(out_height, out_width),
                resampling=Resampling.average,
                masked=True,
            )
        filled = data.filled(0).astype("float32")
        accumulated += filled
        valid_cells |= ~data.mask

    row_idx, col_idx = np.indices(accumulated.shape)
    xs, ys = xy(preview_transform, row_idx, col_idx)
    longitudes = np.asarray(xs).ravel()
    latitudes = np.asarray(ys).ravel()
    totals = accumulated.ravel()
    valid = valid_cells.ravel() & np.isfinite(totals) & (totals >= float(min_precip_mm))

    if preview_crs and preview_crs.to_string() not in {"EPSG:4326", "EPSG:4269"}:
        longitudes, latitudes = transform(
            preview_crs,
            "EPSG:4326",
            longitudes.tolist(),
            latitudes.tolist(),
        )
        longitudes = np.asarray(longitudes)
        latitudes = np.asarray(latitudes)

    preview = pd.DataFrame(
        {
            "longitude": longitudes[valid],
            "latitude": latitudes[valid],
            "precip_mm": totals[valid],
        }
    )

    metadata = {
        "window_days": window_days,
        "available": True,
        "end_date": resolved_end.isoformat(),
        "start_date": window_dates[0].isoformat(),
        "grid_points": int(len(preview)),
        "downsample_factor": downsample_factor,
    }
    return preview, metadata
