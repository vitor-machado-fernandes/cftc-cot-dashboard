from __future__ import annotations

from io import BytesIO
import json
from zipfile import ZipFile

import pandas as pd
import requests
import rasterio
from rasterio.features import geometry_mask, geometry_window
from rasterio.warp import transform_geom

from cotton_weather.config import (
    BOUNDARY_RAW_DIR,
    CENSUS_COUNTY_BOUNDARY_URL,
    CDL_RAW_DIR,
    CDL_COTTON_CLASS_CODE,
    COTTON_STATE_FIPS,
    PROCESSED_DIR,
)


COUNTY_BOUNDARY_ZIP = BOUNDARY_RAW_DIR / "cb_2024_us_county_500k.zip"
COUNTY_GEOJSON_FILE = PROCESSED_DIR / "cotton_counties_2024.geojson"
COUNTY_WEIGHT_FILE = PROCESSED_DIR / "cotton_county_weights_2024.parquet"
STATE_NAME_BY_FIPS = {fips: state for state, fips in COTTON_STATE_FIPS.items()}


def download_county_boundaries(force: bool = False) -> Path:
    if COUNTY_BOUNDARY_ZIP.exists() and not force:
        return COUNTY_BOUNDARY_ZIP

    COUNTY_BOUNDARY_ZIP.parent.mkdir(parents=True, exist_ok=True)
    temp_path = COUNTY_BOUNDARY_ZIP.with_suffix(".part")
    try:
        with requests.get(CENSUS_COUNTY_BOUNDARY_URL, stream=True, timeout=300) as response:
            response.raise_for_status()
            with temp_path.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        output.write(chunk)
        temp_path.replace(COUNTY_BOUNDARY_ZIP)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return COUNTY_BOUNDARY_ZIP


def _load_county_features() -> list[dict]:
    try:
        import shapefile
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "The 'pyshp' package is required for county boundary processing. "
            "Install dependencies with `pip install -r requirements.txt`."
        ) from exc

    zip_path = download_county_boundaries()
    with ZipFile(zip_path) as archive:
        shp_name = next(name for name in archive.namelist() if name.endswith(".shp"))
        shx_name = next(name for name in archive.namelist() if name.endswith(".shx"))
        dbf_name = next(name for name in archive.namelist() if name.endswith(".dbf"))
        reader = shapefile.Reader(
            shp=BytesIO(archive.read(shp_name)),
            shx=BytesIO(archive.read(shx_name)),
            dbf=BytesIO(archive.read(dbf_name)),
        )
        records = []
        valid_state_fips = set(COTTON_STATE_FIPS.values())
        for shape_record in reader.iterShapeRecords():
            record = shape_record.record.as_dict()
            if record["STATEFP"] not in valid_state_fips:
                continue
            records.append(
                {
                    "type": "Feature",
                    "properties": {
                        "GEOID": record["GEOID"],
                        "NAME": record["NAME"],
                        "STATEFP": record["STATEFP"],
                        "STATE_ABBR": STATE_NAME_BY_FIPS[record["STATEFP"]],
                    },
                    "geometry": shape_record.shape.__geo_interface__,
                }
            )
        return records


def load_county_features() -> list[dict]:
    return _load_county_features()


def export_county_geojson() -> Path:
    features = _load_county_features()
    COUNTY_GEOJSON_FILE.parent.mkdir(parents=True, exist_ok=True)
    geojson = {"type": "FeatureCollection", "features": features}
    COUNTY_GEOJSON_FILE.write_text(json.dumps(geojson), encoding="utf-8")
    return COUNTY_GEOJSON_FILE


def _county_rows_for_state(features: list[dict], state_fips: str) -> list[dict]:
    return [feature for feature in features if feature["properties"]["STATEFP"] == state_fips]


def build_county_weights(year: int = 2024) -> pd.DataFrame:
    features = _load_county_features()
    rows: list[dict] = []

    for state_abbr, state_fips in COTTON_STATE_FIPS.items():
        raster_path = CDL_RAW_DIR / str(year) / f"{state_abbr.lower()}_{year}_cdl.tif"
        if not raster_path.exists():
            continue

        state_features = _county_rows_for_state(features, state_fips)
        with rasterio.open(raster_path) as dataset:
            pixel_area_square_meters = abs(dataset.transform.a * dataset.transform.e)
            for feature in state_features:
                geom_proj = transform_geom("EPSG:4269", dataset.crs, feature["geometry"])
                try:
                    window = geometry_window(dataset, [geom_proj], pad_x=0, pad_y=0)
                except Exception:
                    continue

                band = dataset.read(1, window=window, masked=True)
                window_transform = dataset.window_transform(window)
                county_mask = geometry_mask(
                    [geom_proj],
                    transform=window_transform,
                    out_shape=band.shape,
                    invert=True,
                )
                cotton_mask = (band == CDL_COTTON_CLASS_CODE) & ~band.mask & county_mask
                cotton_pixels = int(cotton_mask.sum())
                if cotton_pixels == 0:
                    continue

                rows.append(
                    {
                        "geoid": feature["properties"]["GEOID"],
                        "county_name": feature["properties"]["NAME"],
                        "state": state_abbr,
                        "state_name": state_abbr,
                        "year": year,
                        "cotton_pixels": cotton_pixels,
                        "cotton_area_acres_est": cotton_pixels * pixel_area_square_meters / 4046.8564224,
                    }
                )

    if not rows:
        return pd.DataFrame()

    output = pd.DataFrame(rows).sort_values("cotton_area_acres_est", ascending=False).reset_index(drop=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    output.to_parquet(COUNTY_WEIGHT_FILE, index=False)
    export_county_geojson()
    return output


def load_county_weights() -> pd.DataFrame:
    if not COUNTY_WEIGHT_FILE.exists():
        return pd.DataFrame()
    return pd.read_parquet(COUNTY_WEIGHT_FILE)


def load_county_geojson() -> dict:
    if not COUNTY_GEOJSON_FILE.exists():
        return {}
    return json.loads(COUNTY_GEOJSON_FILE.read_text(encoding="utf-8"))
