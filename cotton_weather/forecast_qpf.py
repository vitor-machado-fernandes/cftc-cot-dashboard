from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
import io
import json
from pathlib import Path
import tarfile

import pandas as pd
import requests
import urllib3

from cotton_weather.config import DATA_DIR


WPC_QPF_BASE_URL = "https://ftp.wpc.ncep.noaa.gov/shapefiles/qpf"
WPC_QPF_MAPSERVER_BASE_URL = "https://mapservices.weather.noaa.gov/vector/rest/services/precip/wpc_qpf/MapServer"
WPC_QPF_RAW_DIR = DATA_DIR / "raw" / "wpc_qpf"
WPC_QPF_CACHE_DIR = DATA_DIR / "processed" / "wpc_qpf_cache"
WPC_QPF_IMAGE_CACHE_DIR = DATA_DIR / "processed" / "wpc_qpf_images"
WPC_REFRESH_HOURS = 6

WPC_QPF_PRODUCTS = {
    "Next 24h": ("day1", "QPF24hr_Day1_latest.tar"),
    "Day 2": ("day2", "QPF24hr_Day2_latest.tar"),
    "Day 3": ("day3", "QPF24hr_Day3_latest.tar"),
    "Days 4-5": ("day45", "QPF48hr_Day4-5_latest.tar"),
    "Days 6-7": ("day67", "QPF48hr_Day6-7_latest.tar"),
    "Next 5 days": ("5day", "QPF120hr_Day1-5_latest.tar"),
    "Next 7 days": ("7day", "QPF168hr_Day1-7_latest.tar"),
}

WPC_QPF_IMAGE_URLS = {
    "Next 24h": "https://www.wpc.ncep.noaa.gov/qpf/fill_94qwbg.gif",
    "Day 2": "https://www.wpc.ncep.noaa.gov/qpf/fill_98qwbg.gif",
    "Day 3": "https://www.wpc.ncep.noaa.gov/qpf/fill_99qwbg.gif",
    "Days 4-5": "https://www.wpc.ncep.noaa.gov/qpf/95ep48iwbg_fill.gif",
    "Days 6-7": "https://www.wpc.ncep.noaa.gov/qpf/97ep48iwbg_fill.gif",
    "Next 5 days": "https://www.wpc.ncep.noaa.gov/qpf/p120i.gif",
    "Next 7 days": "https://www.wpc.ncep.noaa.gov/qpf/p168i.gif",
}

WPC_QPF_DAILY_IMAGE_URLS = {
    1: "https://www.wpc.ncep.noaa.gov/qpf/fill_94qwbg.gif",
    2: "https://www.wpc.ncep.noaa.gov/qpf/fill_98qwbg.gif",
    3: "https://www.wpc.ncep.noaa.gov/qpf/fill_99qwbg.gif",
    4: "https://www.wpc.ncep.noaa.gov/qpf/day4p24iwbg_fill.gif",
    5: "https://www.wpc.ncep.noaa.gov/qpf/day5p24iwbg_fill.gif",
    6: "https://www.wpc.ncep.noaa.gov/qpf/day6p24iwbg_fill.gif",
    7: "https://www.wpc.ncep.noaa.gov/qpf/day7p24iwbg_fill.gif",
}


WPC_QPF_MAPSERVER_LAYERS = {
    "Next 7 days custom": 11,
}


@dataclass(frozen=True)
class WpcQpfAsset:
    product_label: str
    source_url: str
    tar_path: Path
    extract_dir: Path
    shp_path: Path


def _is_stale(path: Path, max_age_hours: int = WPC_REFRESH_HOURS) -> bool:
    if not path.exists():
        return True
    age = datetime.now() - datetime.fromtimestamp(path.stat().st_mtime)
    return age > timedelta(hours=max_age_hours)


def _download_tar(url: str, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_suffix(destination.suffix + ".part")
    try:
        try:
            with requests.get(url, timeout=180) as response:
                response.raise_for_status()
                temp_path.write_bytes(response.content)
        except requests.exceptions.SSLError:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            with requests.get(url, timeout=180, verify=False) as response:
                response.raise_for_status()
                temp_path.write_bytes(response.content)
        temp_path.replace(destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _download_binary(url: str, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_suffix(destination.suffix + ".part")
    try:
        try:
            with requests.get(url, timeout=180) as response:
                response.raise_for_status()
                temp_path.write_bytes(response.content)
        except requests.exceptions.SSLError:
            # Some Windows environments fail cert validation against WPC even though the
            # content is otherwise reachable. Restrict the fallback to these forecast images.
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            with requests.get(url, timeout=180, verify=False) as response:
                response.raise_for_status()
                temp_path.write_bytes(response.content)
        temp_path.replace(destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def ensure_wpc_qpf_asset(product_label: str) -> WpcQpfAsset:
    if product_label not in WPC_QPF_PRODUCTS:
        raise ValueError(f"Unsupported WPC QPF product: {product_label}")

    subdir, filename = WPC_QPF_PRODUCTS[product_label]
    source_url = f"{WPC_QPF_BASE_URL}/{subdir}/{filename}"
    tar_path = WPC_QPF_RAW_DIR / subdir / filename
    extract_dir = WPC_QPF_RAW_DIR / subdir / filename.removesuffix(".tar")

    tar_is_stale = _is_stale(tar_path)
    if tar_is_stale:
        _download_tar(source_url, tar_path)

    extract_dir.mkdir(parents=True, exist_ok=True)
    if tar_is_stale:
        for existing_file in extract_dir.glob('*'):
            if existing_file.is_file():
                existing_file.unlink()
    if tar_is_stale or not any(extract_dir.glob("*.shp")):
        with tarfile.open(tar_path, mode="r") as archive:
            archive.extractall(extract_dir)

    shp_candidates = sorted(extract_dir.glob("*.shp"))
    if not shp_candidates:
        raise FileNotFoundError(f"No shapefile extracted from {tar_path.name}")

    return WpcQpfAsset(
        product_label=product_label,
        source_url=source_url,
        tar_path=tar_path,
        extract_dir=extract_dir,
        shp_path=shp_candidates[0],
    )


def ensure_wpc_qpf_image(image_url: str) -> Path:
    file_name = image_url.rstrip("/").split("/")[-1]
    destination = WPC_QPF_IMAGE_CACHE_DIR / file_name
    if _is_stale(destination):
        _download_binary(image_url, destination)
    return destination


def refresh_wpc_qpf_image_cache() -> list[Path]:
    cached_paths: list[Path] = []
    for product_label in WPC_QPF_PRODUCTS:
        load_wpc_qpf_geojson(product_label)
    for product_label in WPC_QPF_MAPSERVER_LAYERS:
        load_wpc_qpf_mapserver_geojson(product_label)
    for image_url in list(WPC_QPF_IMAGE_URLS.values()) + list(WPC_QPF_DAILY_IMAGE_URLS.values()):
        cached_paths.append(ensure_wpc_qpf_image(image_url))
    return cached_paths


def load_wpc_qpf_geojson(product_label: str) -> tuple[dict, pd.DataFrame, dict]:
    asset = ensure_wpc_qpf_asset(product_label)
    cache_key = asset.tar_path.stem + f"_{int(asset.tar_path.stat().st_mtime)}_v3"
    geojson_path = WPC_QPF_CACHE_DIR / f"{cache_key}.geojson"
    records_path = WPC_QPF_CACHE_DIR / f"{cache_key}.parquet"
    meta_path = WPC_QPF_CACHE_DIR / f"{cache_key}.json"

    if geojson_path.exists() and records_path.exists() and meta_path.exists():
        geojson = json.loads(geojson_path.read_text(encoding="utf-8"))
        records = pd.read_parquet(records_path)
        metadata = json.loads(meta_path.read_text(encoding="utf-8"))
        return geojson, records, metadata

    WPC_QPF_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        import shapefile
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "The 'pyshp' package is required to parse fresh NOAA/WPC QPF shapefiles. "
            "Install it with `pip install pyshp` or run the app from the project .conda environment."
        ) from exc

    reader = shapefile.Reader(str(asset.shp_path))
    fields = [field[0] for field in reader.fields[1:]]

    features = []
    records = []
    for index, shape_record in enumerate(reader.iterShapeRecords()):
        record = dict(zip(fields, shape_record.record))
        qpf_value = float(record.get("QPF", 0.0))
        feature_id = f"{asset.tar_path.stem}_{index}"
        geometry = shape_record.shape.__geo_interface__
        features.append(
            {
                "type": "Feature",
                "id": feature_id,
                "properties": {
                    "feature_id": feature_id,
                    "qpf": qpf_value,
                    "units": record.get("UNITS"),
                    "valid_time": record.get("VALID_TIME"),
                    "product": record.get("PRODUCT"),
                    "issue_time": record.get("ISSUE_TIME"),
                    "start_time": record.get("START_TIME"),
                    "end_time": record.get("END_TIME"),
                },
                "geometry": geometry,
            }
        )
        records.append(
            {
                "feature_id": feature_id,
                "qpf": qpf_value,
                "units": record.get("UNITS"),
                "valid_time": record.get("VALID_TIME"),
                "product": record.get("PRODUCT"),
                "issue_time": record.get("ISSUE_TIME"),
                "start_time": record.get("START_TIME"),
                "end_time": record.get("END_TIME"),
            }
        )

    geojson = {"type": "FeatureCollection", "features": features}
    records_df = pd.DataFrame(records).sort_values("qpf").reset_index(drop=True)
    metadata = {
        "product_label": product_label,
        "source_url": asset.source_url,
        "feature_count": int(len(records_df)),
        "valid_time": records_df["valid_time"].iloc[0] if not records_df.empty else None,
        "issue_time": records_df["issue_time"].iloc[0] if not records_df.empty else None,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }

    geojson_path.write_text(json.dumps(geojson), encoding="utf-8")
    records_df.to_parquet(records_path, index=False)
    meta_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return geojson, records_df, metadata


def load_wpc_qpf_mapserver_geojson(product_label: str) -> tuple[dict, pd.DataFrame, dict]:
    if product_label not in WPC_QPF_MAPSERVER_LAYERS:
        raise ValueError(f"Unsupported WPC MapServer product: {product_label}")

    layer_id = WPC_QPF_MAPSERVER_LAYERS[product_label]
    cache_key = f"mapserver_layer_{layer_id}"
    geojson_path = WPC_QPF_CACHE_DIR / f"{cache_key}.geojson"
    records_path = WPC_QPF_CACHE_DIR / f"{cache_key}.parquet"
    meta_path = WPC_QPF_CACHE_DIR / f"{cache_key}.json"

    if (
        geojson_path.exists() and records_path.exists() and meta_path.exists()
        and not _is_stale(meta_path)
    ):
        geojson = json.loads(geojson_path.read_text(encoding="utf-8"))
        records = pd.read_parquet(records_path)
        metadata = json.loads(meta_path.read_text(encoding="utf-8"))
        return geojson, records, metadata

    WPC_QPF_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    query_url = f"{WPC_QPF_MAPSERVER_BASE_URL}/{layer_id}/query"
    query_params = {
        "where": "1=1",
        "outFields": "*",
        "returnGeometry": "true",
        "f": "geojson",
    }
    try:
        response = requests.get(query_url, params=query_params, timeout=180)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(query_url, params=query_params, timeout=180, verify=False)
    response.raise_for_status()
    geojson = response.json()

    records = []
    features = geojson.get("features", [])
    for index, feature in enumerate(features):
        properties = feature.get("properties", {}) or {}
        feature_id = str(properties.get("OBJECTID") or properties.get("FID") or index)
        feature["id"] = feature_id
        properties["feature_id"] = feature_id
        feature["properties"] = properties
        qpf_value = properties.get("qpf", properties.get("QPF"))
        try:
            qpf_value = float(qpf_value)
        except (TypeError, ValueError):
            qpf_value = None
        records.append(
            {
                "feature_id": feature_id,
                "qpf": qpf_value,
                "product": properties.get("product") or properties.get("PRODUCT"),
                "valid_time": properties.get("validTime") or properties.get("VALID_TIME"),
                "issue_time": properties.get("issueTime") or properties.get("ISSUE_TIME"),
                "start_time": properties.get("startTime") or properties.get("START_TIME"),
                "end_time": properties.get("endTime") or properties.get("END_TIME"),
            }
        )

    records_df = pd.DataFrame(records)
    if not records_df.empty and "qpf" in records_df.columns:
        records_df = records_df.sort_values("qpf").reset_index(drop=True)

    metadata = {
        "product_label": product_label,
        "source_url": query_url,
        "layer_id": layer_id,
        "feature_count": int(len(features)),
        "valid_time": records_df["valid_time"].dropna().iloc[0] if not records_df.empty and records_df["valid_time"].notna().any() else None,
        "issue_time": records_df["issue_time"].dropna().iloc[0] if not records_df.empty and records_df["issue_time"].notna().any() else None,
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }

    geojson_path.write_text(json.dumps(geojson), encoding="utf-8")
    records_df.to_parquet(records_path, index=False)
    meta_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return geojson, records_df, metadata
