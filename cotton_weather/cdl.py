from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
import hashlib
import json
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
import pandas as pd
import rasterio
import requests
from rasterio.features import rasterize
from rasterio.transform import xy
from rasterio.warp import transform, transform_geom

from cotton_weather.config import (
    CDL_CONUS_START_YEAR,
    CDL_COTTON_CLASS_CODE,
    CDL_FAQ_URL,
    CDL_METADATA_URL,
    CDL_PORTAL_URL,
    CDL_RAW_DIR,
    CDL_SERVICE_URL,
    COTTON_STATE_FIPS,
    PROCESSED_DIR,
)
from cotton_weather.counties import load_county_features
from cotton_weather.county_irrigation import (
    COUNTY_IRRIGATION_SOURCE_FILE,
    load_county_irrigation_share_source,
)


CDL_FOOTPRINT_PLAN_FILE = PROCESSED_DIR / "cdl_footprint_plan.json"
CDL_FOOTPRINT_SUMMARY_FILE = PROCESSED_DIR / "cdl_cotton_footprint_summary.parquet"
CDL_FOOTPRINT_METADATA_FILE = PROCESSED_DIR / "cdl_cotton_footprint_metadata.json"
CDL_PREVIEW_CACHE_DIR = PROCESSED_DIR / "cdl_preview_cache"
CDL_SUMMARY_CACHE_DIR = PROCESSED_DIR / "cdl_summary_cache"


class CdlDownloadError(RuntimeError):
    """Raised when a CDL asset cannot be downloaded or parsed."""


@dataclass(frozen=True)
class CdlFootprintPlan:
    crop_name: str
    class_code: int
    target_year: int
    source_portal: str
    source_metadata: str
    source_faq: str
    conus_full_coverage_start_year: int
    methodology: str
    geometry_type: str
    status: str
    created_at: str
    notes: list[str]


def create_cdl_footprint_plan(target_year: int) -> dict:
    if target_year < CDL_CONUS_START_YEAR:
        raise ValueError(
            f"Target year {target_year} is earlier than the CDL full-CONUS coverage year "
            f"{CDL_CONUS_START_YEAR}."
        )

    plan = CdlFootprintPlan(
        crop_name="Cotton",
        class_code=CDL_COTTON_CLASS_CODE,
        target_year=target_year,
        source_portal=CDL_PORTAL_URL,
        source_metadata=CDL_METADATA_URL,
        source_faq=CDL_FAQ_URL,
        conus_full_coverage_start_year=CDL_CONUS_START_YEAR,
        methodology=(
            "Use the most recent USDA NASS Cropland Data Layer raster as the crop-footprint base. "
            "Download state CDL rasters for cotton-producing states, filter pixels with class code 2 "
            "(Cotton), and use the resulting cotton mask as the fixed footprint for PRISM aggregation."
        ),
        geometry_type="Raster-derived cotton mask",
        status="planned",
        created_at=datetime.now().isoformat(timespec="seconds"),
        notes=[
            "The CDL class available in official metadata is 'Cotton' (code 2).",
            "The official CDL class list reviewed here does not distinguish upland cotton from pima cotton.",
            "For this project we will initially treat the latest CDL Cotton class as a proxy for upland cotton unless a better separating dataset is added.",
            "The weather history can be overlaid on one fixed recent-year cotton footprint, which avoids having to rebuild crop area for every historical year.",
        ],
    )
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    serialized = asdict(plan)
    CDL_FOOTPRINT_PLAN_FILE.write_text(json.dumps(serialized, indent=2), encoding="utf-8")
    return serialized


def load_cdl_footprint_plan() -> dict:
    if not CDL_FOOTPRINT_PLAN_FILE.exists():
        return {}
    return json.loads(CDL_FOOTPRINT_PLAN_FILE.read_text(encoding="utf-8"))


def _parse_return_url(xml_text: str) -> str:
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise CdlDownloadError("Could not parse CDL service response.") from exc

    for element in root.iter():
        if element.tag.endswith("returnURL") and element.text:
            return element.text.strip()
    raise CdlDownloadError("CDL service response did not include a returnURL.")


def request_cdl_file_url(year: int, fips_code: str) -> str:
    response = requests.post(
        CDL_SERVICE_URL,
        data={"year": str(year), "fips": fips_code},
        timeout=120,
    )
    response.raise_for_status()
    return _parse_return_url(response.text)


def download_cdl_raster(year: int, state_abbr: str, source_url: str, force: bool = False):
    if state_abbr not in COTTON_STATE_FIPS:
        raise ValueError(f"Unsupported cotton state abbreviation: {state_abbr}")

    destination = CDL_RAW_DIR / str(year) / f"{state_abbr.lower()}_{year}_cdl.tif"
    if destination.exists() and not force:
        return destination

    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_suffix(".part")
    try:
        with requests.get(source_url, stream=True, timeout=300) as response:
            response.raise_for_status()
            with temp_path.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        output.write(chunk)
        temp_path.replace(destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return destination


def summarize_cotton_pixels(raster_path, state_abbr: str, year: int) -> dict:
    with rasterio.open(raster_path) as dataset:
        band = dataset.read(1, masked=True)
        cotton_mask = (band == CDL_COTTON_CLASS_CODE) & ~band.mask
        cotton_pixels = int(np.count_nonzero(cotton_mask))
        total_pixels = int(np.count_nonzero(~band.mask))
        pixel_area_square_meters = abs(dataset.transform.a * dataset.transform.e)
        cotton_area_acres = cotton_pixels * pixel_area_square_meters / 4046.8564224

    return {
        "state": state_abbr,
        "year": year,
        "raster_path": str(raster_path),
        "cotton_pixels": cotton_pixels,
        "total_valid_pixels": total_pixels,
        "cotton_share": (cotton_pixels / total_pixels) if total_pixels else 0.0,
        "cotton_area_acres_est": cotton_area_acres,
    }


def _normalize_states(states: list[str] | None) -> list[str]:
    if not states:
        return sorted(COTTON_STATE_FIPS.keys())
    normalized = [state.strip().upper() for state in states]
    invalid = sorted(set(normalized).difference(COTTON_STATE_FIPS))
    if invalid:
        invalid_text = ", ".join(invalid)
        raise ValueError(f"Unsupported cotton state abbreviations: {invalid_text}")
    return normalized


def build_latest_cotton_footprint(
    year: int,
    force_download: bool = False,
    states: list[str] | None = None,
) -> dict:
    if year < CDL_CONUS_START_YEAR:
        raise ValueError(
            f"Target year {year} is earlier than the CDL full-CONUS coverage year "
            f"{CDL_CONUS_START_YEAR}."
        )

    selected_states = _normalize_states(states)
    summaries = []
    download_sources: dict[str, str] = {}
    for state_abbr in selected_states:
        fips_code = COTTON_STATE_FIPS[state_abbr]
        source_url = request_cdl_file_url(year=year, fips_code=fips_code)
        download_sources[state_abbr] = source_url
        raster_path = download_cdl_raster(
            year=year,
            state_abbr=state_abbr,
            source_url=source_url,
            force=force_download,
        )
        summaries.append(
            summarize_cotton_pixels(
                raster_path=raster_path,
                state_abbr=state_abbr,
                year=year,
            )
        )

    output = pd.DataFrame(summaries).sort_values("cotton_area_acres_est", ascending=False)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    output.to_parquet(CDL_FOOTPRINT_SUMMARY_FILE, index=False)

    metadata = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "target_year": year,
        "class_code": CDL_COTTON_CLASS_CODE,
        "crop_name": "Cotton",
        "state_count": len(output),
        "states": selected_states,
        "download_sources": download_sources,
        "summary_file": str(CDL_FOOTPRINT_SUMMARY_FILE),
        "notes": [
            "This is a latest-year cotton footprint summary by state raster.",
            "It is intended to define the fixed crop area that later PRISM aggregation will use.",
        ],
    }
    CDL_FOOTPRINT_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return metadata


def list_downloaded_cdl_rasters(year: int | None = None) -> list:
    base_dir = CDL_RAW_DIR / str(year) if year is not None else CDL_RAW_DIR
    if not base_dir.exists():
        return []
    return sorted(base_dir.rglob("*_cdl.tif"))


def _raster_signature(raster_paths: list[Path]) -> str:
    if not raster_paths:
        return "empty"
    tokens = [
        f"{path.name}:{int(path.stat().st_size)}:{int(path.stat().st_mtime)}"
        for path in raster_paths
    ]
    joined = "|".join(tokens)
    return hashlib.sha1(joined.encode("utf-8")).hexdigest()[:16]


def _cache_paths(cache_dir: Path, prefix: str, signature: str) -> tuple[Path, Path]:
    cache_dir.mkdir(parents=True, exist_ok=True)
    data_path = cache_dir / f"{prefix}_{signature}.parquet"
    meta_path = cache_dir / f"{prefix}_{signature}.json"
    return data_path, meta_path


def load_cdl_footprint_summary() -> pd.DataFrame:
    if not CDL_FOOTPRINT_SUMMARY_FILE.exists():
        return pd.DataFrame()
    return pd.read_parquet(CDL_FOOTPRINT_SUMMARY_FILE)


def load_cdl_footprint_metadata() -> dict:
    if not CDL_FOOTPRINT_METADATA_FILE.exists():
        return {}
    return json.loads(CDL_FOOTPRINT_METADATA_FILE.read_text(encoding="utf-8"))


def summarize_downloaded_cdl_rasters(year: int | None = None) -> pd.DataFrame:
    raster_paths = list_downloaded_cdl_rasters(year=year)
    if not raster_paths:
        return pd.DataFrame()

    signature = _raster_signature(raster_paths)
    summary_path, meta_path = _cache_paths(CDL_SUMMARY_CACHE_DIR, "downloaded_cdl_summary", signature)
    if summary_path.exists() and meta_path.exists():
        return pd.read_parquet(summary_path)

    summaries = []
    for raster_path in raster_paths:
        state_abbr = raster_path.stem.split("_")[0].upper()
        raster_year = int(raster_path.stem.split("_")[1])
        summaries.append(
            summarize_cotton_pixels(
                raster_path=raster_path,
                state_abbr=state_abbr,
                year=raster_year,
            )
        )
    output = pd.DataFrame(summaries).sort_values("cotton_area_acres_est", ascending=False)
    output.to_parquet(summary_path, index=False)
    meta_path.write_text(
        json.dumps(
            {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "signature": signature,
                "rows": int(len(output)),
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return output


def _allocate_preview_points(
    summary: pd.DataFrame,
    max_total_points: int,
    min_points_per_state: int,
) -> dict[str, int]:
    total_pixels = int(summary["cotton_pixels"].sum())
    if total_pixels <= 0:
        return {}

    allocations: dict[str, int] = {}
    remainders: list[tuple[float, str]] = []
    allocated = 0
    for row in summary.itertuples(index=False):
        raw_points = max_total_points * (row.cotton_pixels / total_pixels)
        points = min(int(row.cotton_pixels), max(min_points_per_state, int(np.floor(raw_points))))
        allocations[row.state] = points
        allocated += points
        remainders.append((raw_points - np.floor(raw_points), row.state))

    if allocated > max_total_points:
        for _, state in sorted(remainders):
            if allocated <= max_total_points:
                break
            if allocations[state] > min_points_per_state:
                allocations[state] -= 1
                allocated -= 1
    elif allocated < max_total_points:
        for _, state in sorted(remainders, reverse=True):
            if allocated >= max_total_points:
                break
            allocations[state] += 1
            allocated += 1

    return allocations


def build_cotton_map_preview(
    max_total_points: int = 12000,
    min_points_per_state: int = 50,
    states: list[str] | None = None,
    irrigation_mode: str = "none",
) -> pd.DataFrame:
    raster_paths = list_downloaded_cdl_rasters()
    selected_states = set(_normalize_states(states))
    if not raster_paths:
        return pd.DataFrame()

    signature = _raster_signature(raster_paths)
    state_key = "-".join(sorted(selected_states))
    irrigation_signature = "none"
    if irrigation_mode == "county_share" and COUNTY_IRRIGATION_SOURCE_FILE.exists():
        irrigation_signature = f"county_share_{int(COUNTY_IRRIGATION_SOURCE_FILE.stat().st_mtime)}_{int(COUNTY_IRRIGATION_SOURCE_FILE.stat().st_size)}"
    preview_path, meta_path = _cache_paths(
        CDL_PREVIEW_CACHE_DIR,
        f"cotton_preview_{state_key}_{max_total_points}_{min_points_per_state}_{irrigation_mode}",
        f"{signature}_{irrigation_signature}",
    )
    if preview_path.exists() and meta_path.exists():
        return pd.read_parquet(preview_path)

    state_summary = summarize_downloaded_cdl_rasters()
    if state_summary.empty:
        return pd.DataFrame()
    state_summary = state_summary[state_summary["state"].isin(selected_states)].copy()
    point_allocations = _allocate_preview_points(
        summary=state_summary,
        max_total_points=max_total_points,
        min_points_per_state=min_points_per_state,
    )
    county_features = load_county_features() if irrigation_mode == "county_share" else []
    county_by_state: dict[str, list[dict]] = {}
    county_share_lookup = {}
    if irrigation_mode == "county_share":
        county_share_frame = load_county_irrigation_share_source()
        county_share_lookup = county_share_frame.set_index("geoid")["irrigated_share"].to_dict() if not county_share_frame.empty else {}
        for feature in county_features:
            county_by_state.setdefault(feature["properties"]["STATE_ABBR"], []).append(feature)

    preview_frames: list[pd.DataFrame] = []
    for raster_path in raster_paths:
        state_abbr = raster_path.stem.split("_")[0].upper()
        if state_abbr not in selected_states:
            continue
        target_points = point_allocations.get(state_abbr, 0)
        if target_points <= 0:
            continue
        year = int(raster_path.stem.split("_")[1])
        with rasterio.open(raster_path) as dataset:
            band = dataset.read(1, masked=True)
            cotton_rows, cotton_cols = np.where((band == CDL_COTTON_CLASS_CODE) & ~band.mask)
            pixel_count = len(cotton_rows)
            if pixel_count == 0:
                continue

            if pixel_count > target_points:
                sample_idx = np.linspace(0, pixel_count - 1, num=target_points, dtype=int)
                cotton_rows = cotton_rows[sample_idx]
                cotton_cols = cotton_cols[sample_idx]

            x_coords, y_coords = xy(
                dataset.transform,
                cotton_rows,
                cotton_cols,
                offset="center",
            )
            if irrigation_mode == "county_share":
                geoid_values = np.full(len(x_coords), "", dtype=object)
                shapes: list[tuple[dict, int]] = []
                for feature in county_by_state.get(state_abbr, []):
                    geom_proj = transform_geom("EPSG:4269", dataset.crs, feature["geometry"])
                    geoid_int = int(feature["properties"]["GEOID"])
                    shapes.append((geom_proj, geoid_int))
                if shapes:
                    county_raster = rasterize(
                        shapes,
                        out_shape=band.shape,
                        transform=dataset.transform,
                        fill=0,
                        dtype="int32",
                    )
                    sampled_geoids = county_raster[cotton_rows, cotton_cols]
                    geoid_values = np.where(sampled_geoids > 0, pd.Series(sampled_geoids).astype(str).str.zfill(5), "")
                point_frame = pd.DataFrame({"geoid": geoid_values})
                point_frame["irrigated_share"] = point_frame["geoid"].map(county_share_lookup)
                irrigation_frame = pd.DataFrame(
                    {
                        "irrigation_class": np.where(point_frame["irrigated_share"].notna(), "county_share", "unknown"),
                        "irrigation_value": point_frame["irrigated_share"].tolist(),
                        "county_geoid": point_frame["geoid"].tolist(),
                    }
                )
            else:
                irrigation_frame = pd.DataFrame(
                    {
                        "irrigation_class": ["unknown"] * len(x_coords),
                        "irrigation_value": [pd.NA] * len(x_coords),
                    }
                )

            longitudes, latitudes = x_coords, y_coords
            if dataset.crs and dataset.crs.to_string() != "EPSG:4326":
                longitudes, latitudes = transform(
                    dataset.crs,
                    "EPSG:4326",
                    list(longitudes),
                    list(latitudes),
                )
            preview_frames.append(
                pd.DataFrame(
                    {
                        "state": state_abbr,
                        "year": year,
                        "latitude": latitudes,
                        "longitude": longitudes,
                        "state_preview_points": len(latitudes),
                        "irrigation_class": irrigation_frame["irrigation_class"].tolist(),
                        "irrigation_value": irrigation_frame["irrigation_value"].tolist(),
                        "county_geoid": irrigation_frame["county_geoid"].tolist() if "county_geoid" in irrigation_frame.columns else [pd.NA] * len(latitudes),
                    }
                )
            )

    if not preview_frames:
        return pd.DataFrame()
    preview = pd.concat(preview_frames, ignore_index=True)
    preview.to_parquet(preview_path, index=False)
    meta_path.write_text(
        json.dumps(
            {
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "signature": signature,
                "rows": int(len(preview)),
                "states": sorted(selected_states),
                "max_total_points": max_total_points,
                "min_points_per_state": min_points_per_state,
                "irrigation_mode": irrigation_mode,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return preview
