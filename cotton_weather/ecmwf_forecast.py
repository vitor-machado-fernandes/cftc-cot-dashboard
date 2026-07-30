from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from ecmwf.opendata import Client
import numpy as np
import pandas as pd
import shapefile
import rasterio
from rasterio.features import rasterize
from rasterio.transform import from_origin
from rasterio.warp import Resampling, reproject
import xarray as xr

from cotton_weather.cdl import summarize_downloaded_cdl_rasters
from cotton_weather.config import (
    BOUNDARY_RAW_DIR,
    CDL_COTTON_CLASS_CODE,
    CDL_RAW_DIR,
    COTTON_STATE_FIPS,
    DATA_DIR,
)


ECMWF_RAW_DIR = DATA_DIR / "raw" / "ecmwf"
ECMWF_PROCESSED_DIR = DATA_DIR / "processed" / "ecmwf"
ECMWF_WEIGHT_DIR = ECMWF_PROCESSED_DIR / "state_weights"
COUNTY_BOUNDARY_ZIP = BOUNDARY_RAW_DIR / "cb_2024_us_county_500k.zip"
ECMWF_FORECAST_STEP_HOURS = 120
ECMWF_DAILY_STEP_HOURS = tuple(range(24, 361, 24))
ECMWF_FORECAST_WINDOWS = {
    "3 days": 72,
    "5 days": 120,
    "7 days": 168,
    "15 days": 360,
}
ECMWF_REGION_MAP = {
    "Delta Region": ["AR", "TN", "MO", "MS", "AL", "LA"],
    "Southeast Region": ["GA", "FL", "SC", "NC", "VA"],
    "Southwest Region": ["TX", "OK", "KS", "NM"],
    "Far West Region": ["AZ", "CA"],
}
ECMWF_REGION_ORDER = ["Delta Region", "Southeast Region", "Southwest Region", "Far West Region"]
ECMWF_VERIFY_SSL = False


def latest_long_range_cycle(now_utc: datetime | None = None) -> datetime:
    current = now_utc or datetime.now(timezone.utc)
    base = current.replace(minute=0, second=0, microsecond=0)
    if current.hour >= 18:
        return base.replace(hour=12)
    if current.hour >= 6:
        return base.replace(hour=0)
    previous_day = base - timedelta(days=1)
    return previous_day.replace(hour=12)


def refresh_latest_ecmwf_precip(
    step_hours: int = ECMWF_FORECAST_STEP_HOURS,
    verify: bool = ECMWF_VERIFY_SSL,
    cycle_time: datetime | None = None,
) -> Path:
    ECMWF_RAW_DIR.mkdir(parents=True, exist_ok=True)
    target = ECMWF_RAW_DIR / f"ecmwf_tp_step{step_hours}.grib2"
    selected_cycle = cycle_time or latest_long_range_cycle()
    client = Client(source="ecmwf", verify=verify)
    client.retrieve(
        model="ifs",
        stream="oper",
        type="fc",
        levtype="sfc",
        param="tp",
        date=selected_cycle.strftime("%Y%m%d"),
        time=selected_cycle.hour,
        step=step_hours,
        target=str(target),
    )
    return target


def refresh_latest_ecmwf_precip_suite(
    step_hours_list: tuple[int, ...] = ECMWF_DAILY_STEP_HOURS,
    verify: bool = ECMWF_VERIFY_SSL,
) -> list[Path]:
    selected_cycle = latest_long_range_cycle()
    refreshed: list[Path] = []
    for step_hours in step_hours_list:
        refreshed.append(refresh_latest_ecmwf_precip(step_hours=step_hours, verify=verify, cycle_time=selected_cycle))
    return refreshed


def latest_ecmwf_file(step_hours: int = ECMWF_FORECAST_STEP_HOURS) -> Path | None:
    direct = ECMWF_RAW_DIR / f"ecmwf_tp_step{step_hours}.grib2"
    if direct.exists():
        return direct
    candidates = sorted(ECMWF_RAW_DIR.glob(f"*step{step_hours}*.grib2"), key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def _load_us_land_geometries() -> list[dict]:
    with ZipFile(COUNTY_BOUNDARY_ZIP) as archive:
        shp_name = next(name for name in archive.namelist() if name.endswith(".shp"))
        shx_name = next(name for name in archive.namelist() if name.endswith(".shx"))
        dbf_name = next(name for name in archive.namelist() if name.endswith(".dbf"))
        reader = shapefile.Reader(
            shp=BytesIO(archive.read(shp_name)),
            shx=BytesIO(archive.read(shx_name)),
            dbf=BytesIO(archive.read(dbf_name)),
        )
        return [shape_record.shape.__geo_interface__ for shape_record in reader.iterShapeRecords()]


def _land_mask_cache_path(
    lon_bounds: tuple[float, float],
    lat_bounds: tuple[float, float],
    interpolation_step_degrees: float,
) -> Path:
    west, east = lon_bounds
    south, north = lat_bounds
    step_token = str(interpolation_step_degrees).replace(".", "p")
    return ECMWF_PROCESSED_DIR / (
        f"us_land_mask_w{west}_e{east}_s{south}_n{north}_step{step_token}.npz"
    )


def load_us_land_mask(
    target_lons: np.ndarray,
    target_lats: np.ndarray,
    lon_bounds: tuple[float, float],
    lat_bounds: tuple[float, float],
    interpolation_step_degrees: float,
) -> np.ndarray:
    cache_path = _land_mask_cache_path(lon_bounds, lat_bounds, interpolation_step_degrees)
    if cache_path.exists():
        with np.load(cache_path) as cached:
            return cached["mask"].astype(bool)

    ECMWF_PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    geometries = _load_us_land_geometries()
    transform = from_origin(
        float(target_lons.min()) - interpolation_step_degrees / 2,
        float(target_lats.max()) + interpolation_step_degrees / 2,
        interpolation_step_degrees,
        interpolation_step_degrees,
    )
    mask = rasterize(
        [(geometry, 1) for geometry in geometries],
        out_shape=(len(target_lats), len(target_lons)),
        transform=transform,
        fill=0,
        dtype="uint8",
        all_touched=False,
    ).astype(bool)
    np.savez_compressed(cache_path, mask=mask)
    return mask


def _cdl_raster_path(state_abbr: str, footprint_year: int) -> Path:
    return CDL_RAW_DIR / str(footprint_year) / f"{state_abbr.lower()}_{footprint_year}_cdl.tif"


def ensure_state_ecmwf_weights(
    state_abbr: str,
    ecmwf_grib_path: Path,
    footprint_year: int = 2024,
) -> Path:
    weight_path = ECMWF_WEIGHT_DIR / str(footprint_year) / f"{state_abbr.lower()}_ecmwf_weights.tif"
    if weight_path.exists():
        try:
            with rasterio.open(weight_path):
                return weight_path
        except Exception:
            weight_path.unlink(missing_ok=True)

    cdl_path = _cdl_raster_path(state_abbr, footprint_year)
    if not cdl_path.exists():
        raise FileNotFoundError(f"Missing CDL raster for {state_abbr}: {cdl_path}")

    ECMWF_WEIGHT_DIR.mkdir(parents=True, exist_ok=True)
    weight_path.parent.mkdir(parents=True, exist_ok=True)

    with rasterio.open(cdl_path) as cdl_ds, rasterio.open(ecmwf_grib_path) as ecmwf_ds:
        cdl_band = cdl_ds.read(1, masked=True)
        source_mask = np.where(
            (cdl_band == CDL_COTTON_CLASS_CODE) & ~cdl_band.mask,
            1.0,
            0.0,
        ).astype("float32")
        destination = np.zeros((ecmwf_ds.height, ecmwf_ds.width), dtype="float32")
        reproject(
            source=source_mask,
            destination=destination,
            src_transform=cdl_ds.transform,
            src_crs=cdl_ds.crs,
            src_nodata=0.0,
            dst_transform=ecmwf_ds.transform,
            dst_crs=ecmwf_ds.crs,
            dst_nodata=0.0,
            resampling=Resampling.average,
        )
        metadata = ecmwf_ds.meta.copy()
        metadata.update(dtype="float32", count=1, compress="lzw", nodata=0.0)
        with rasterio.open(weight_path, "w", **metadata) as output:
            output.write(destination, 1)

    return weight_path


def aggregate_state_ecmwf_forecast(
    step_hours: int,
    footprint_year: int = 2024,
    states: list[str] | None = None,
) -> pd.DataFrame:
    source_path = latest_ecmwf_file(step_hours=step_hours)
    if source_path is None:
        return pd.DataFrame()

    selected_states = states or sorted(COTTON_STATE_FIPS.keys())
    state_area = summarize_downloaded_cdl_rasters(year=footprint_year)
    area_lookup = state_area.set_index("state")["cotton_area_acres_est"].to_dict() if not state_area.empty else {}

    with xr.open_dataset(source_path, engine="cfgrib") as dataset:
        precip_mm = dataset["tp"].values * 1000.0
        valid_time = pd.to_datetime(dataset["valid_time"].values).isoformat() if "valid_time" in dataset.coords else None
        issue_time = pd.to_datetime(dataset["time"].values).isoformat() if "time" in dataset.coords else None

    rows: list[dict] = []
    for state_abbr in selected_states:
        if not _cdl_raster_path(state_abbr, footprint_year).exists():
            continue
        weight_path = ensure_state_ecmwf_weights(
            state_abbr=state_abbr,
            ecmwf_grib_path=source_path,
            footprint_year=footprint_year,
        )
        with rasterio.open(weight_path) as weight_ds:
            weights = weight_ds.read(1, masked=True)

        valid_mask = (~np.isnan(precip_mm)) & (weights.data > 0)
        if not np.any(valid_mask):
            weighted_mean = None
            weight_sum = 0.0
        else:
            valid_weights = weights.data[valid_mask]
            valid_precip = precip_mm[valid_mask]
            weight_sum = float(valid_weights.sum())
            weighted_mean = float(np.average(valid_precip, weights=valid_weights)) if weight_sum > 0 else None

        rows.append(
            {
                "step_hours": step_hours,
                "forecast_day": step_hours // 24,
                "state": state_abbr,
                "ppt_mm": weighted_mean,
                "ecmwf_weight_sum": weight_sum,
                "cotton_area_acres_est": area_lookup.get(state_abbr),
                "issue_time": issue_time,
                "valid_time": valid_time,
            }
        )

    return pd.DataFrame(rows)


def build_ecmwf_regional_forecast_index(
    max_step_hours: int,
    footprint_year: int = 2024,
) -> tuple[pd.DataFrame, dict]:
    available_steps = [step for step in ECMWF_DAILY_STEP_HOURS if step <= max_step_hours and latest_ecmwf_file(step) is not None]
    if not available_steps:
        return pd.DataFrame(), {"available": False, "max_step_hours": max_step_hours}

    all_state_rows: list[pd.DataFrame] = []
    for step_hours in available_steps:
        step_rows = aggregate_state_ecmwf_forecast(step_hours=step_hours, footprint_year=footprint_year)
        if not step_rows.empty:
            all_state_rows.append(step_rows)

    if not all_state_rows:
        return pd.DataFrame(), {"available": False, "max_step_hours": max_step_hours}

    state_forecast = pd.concat(all_state_rows, ignore_index=True)
    national_area = float(state_forecast[["state", "cotton_area_acres_est"]].drop_duplicates()["cotton_area_acres_est"].fillna(0.0).sum())
    if national_area <= 0:
        return pd.DataFrame(), {"available": False, "max_step_hours": max_step_hours}

    state_forecast["national_contribution_mm"] = (
        state_forecast["ppt_mm"].fillna(0.0) * state_forecast["cotton_area_acres_est"].fillna(0.0) / national_area
    )

    regional_rows: list[dict] = []
    for step_hours in available_steps:
        step_frame = state_forecast.loc[state_forecast["step_hours"] == step_hours].copy()
        issue_time = step_frame["issue_time"].dropna().iloc[0] if not step_frame["issue_time"].dropna().empty else None
        valid_time = step_frame["valid_time"].dropna().iloc[0] if not step_frame["valid_time"].dropna().empty else None
        for region_name in ECMWF_REGION_ORDER:
            region_states = ECMWF_REGION_MAP[region_name]
            region_frame = step_frame.loc[step_frame["state"].isin(region_states)]
            regional_rows.append(
                {
                    "forecast_day": step_hours // 24,
                    "step_hours": step_hours,
                    "region": region_name,
                    "cumulative_precip_mm": float(region_frame["national_contribution_mm"].sum()),
                    "valid_time": valid_time,
                    "issue_time": issue_time,
                }
            )

    regional_df = pd.DataFrame(regional_rows)
    if regional_df.empty:
        return pd.DataFrame(), {"available": False, "max_step_hours": max_step_hours}

    regional_df["valid_date"] = pd.to_datetime(regional_df["valid_time"]).dt.date.astype(str)

    issue_time = regional_df["issue_time"].dropna().iloc[0] if not regional_df["issue_time"].dropna().empty else None
    if issue_time is not None:
        issue_date = pd.to_datetime(issue_time).date().isoformat()
        zero_rows = pd.DataFrame(
            [
                {
                    "forecast_day": 0,
                    "step_hours": 0,
                    "region": region_name,
                    "cumulative_precip_mm": 0.0,
                    "valid_time": issue_time,
                    "issue_time": issue_time,
                    "valid_date": issue_date,
                }
                for region_name in ECMWF_REGION_ORDER
            ]
        )
        regional_df = pd.concat([zero_rows, regional_df], ignore_index=True)

    regional_df["region"] = pd.Categorical(regional_df["region"], categories=ECMWF_REGION_ORDER, ordered=True)
    regional_df = regional_df.sort_values(["forecast_day", "region"]).reset_index(drop=True)

    metadata = {
        "available": True,
        "max_step_hours": max_step_hours,
        "available_steps": available_steps,
        "issue_time": regional_df["issue_time"].dropna().iloc[0] if not regional_df["issue_time"].dropna().empty else None,
        "latest_valid_time": regional_df["valid_time"].dropna().iloc[-1] if not regional_df["valid_time"].dropna().empty else None,
        "national_cotton_area_acres": national_area,
    }
    return regional_df, metadata


def build_ecmwf_precip_map_preview(
    step_hours: int = ECMWF_FORECAST_STEP_HOURS,
    min_precip_mm: float = 0.1,
    lon_bounds: tuple[float, float] = (-127.0, -65.0),
    lat_bounds: tuple[float, float] = (22.0, 52.0),
    interpolation_step_degrees: float = 0.1,
) -> tuple[pd.DataFrame, dict]:
    source_path = latest_ecmwf_file(step_hours=step_hours)
    if source_path is None:
        return pd.DataFrame(), {"available": False, "step_hours": step_hours}

    dataset = xr.open_dataset(source_path, engine="cfgrib")
    try:
        precip = dataset["tp"]
        precip_mm = precip * 1000.0
        cropped = precip_mm.sel(
            longitude=slice(lon_bounds[0], lon_bounds[1]),
            latitude=slice(lat_bounds[1], lat_bounds[0]),
        )
        target_lons = np.arange(lon_bounds[0], lon_bounds[1] + interpolation_step_degrees, interpolation_step_degrees)
        target_lats = np.arange(lat_bounds[1], lat_bounds[0] - interpolation_step_degrees, -interpolation_step_degrees)
        interpolated = cropped.interp(longitude=target_lons, latitude=target_lats, method="linear")
        land_mask = load_us_land_mask(
            target_lons=target_lons,
            target_lats=target_lats,
            lon_bounds=lon_bounds,
            lat_bounds=lat_bounds,
            interpolation_step_degrees=interpolation_step_degrees,
        )
        lon_grid, lat_grid = np.meshgrid(target_lons, target_lats)
        frame = pd.DataFrame(
            {
                "longitude": lon_grid.ravel(),
                "latitude": lat_grid.ravel(),
                "precip_mm": interpolated.values.ravel(),
                "is_land": land_mask.ravel(),
            }
        )
        frame = frame.loc[
            frame["is_land"]
            & frame["precip_mm"].notna()
            & (frame["precip_mm"] >= float(min_precip_mm))
        ].drop(columns="is_land").copy()
        valid_time = pd.to_datetime(interpolated["valid_time"].values).isoformat() if "valid_time" in interpolated.coords else None
        issue_time = pd.to_datetime(interpolated["time"].values).isoformat() if "time" in interpolated.coords else None
    finally:
        dataset.close()

    metadata = {
        "available": True,
        "source_path": str(source_path),
        "step_hours": step_hours,
        "grid_points": int(len(frame)),
        "valid_time": valid_time,
        "issue_time": issue_time,
        "min_precip_mm": min_precip_mm,
        "interpolation_step_degrees": interpolation_step_degrees,
        "land_clipped": True,
    }
    return frame, metadata
