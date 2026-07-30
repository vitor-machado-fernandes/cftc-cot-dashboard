from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
from zipfile import BadZipFile, ZipFile, is_zipfile

import pandas as pd
import rasterio
from rasterio.warp import transform
import requests

from cotton_weather.config import PRISM_BASE_URL, SUPPORTED_VARIABLES


class PrismDownloadError(RuntimeError):
    """Raised when a PRISM asset cannot be downloaded."""


@dataclass(frozen=True)
class PrismAsset:
    variable: str
    date_value: date
    zip_path: Path
    tif_path: Path
    source_url: str


def build_prism_url(variable: str, date_value: date) -> str:
    if variable not in SUPPORTED_VARIABLES:
        raise ValueError(f"Unsupported PRISM variable: {variable}")
    year = date_value.strftime("%Y")
    day_token = date_value.strftime("%Y%m%d")
    file_name = f"prism_{variable}_us_30s_{day_token}.zip"
    return f"{PRISM_BASE_URL}/{variable}/daily/{year}/{file_name}"


def _stream_download(url: str, temp_path: Path, timeout: int, verify: bool) -> None:
    with requests.get(url, stream=True, timeout=timeout, verify=verify) as response:
        if response.status_code == 404:
            raise PrismDownloadError(f"PRISM data not available yet: {url}")
        response.raise_for_status()
        with temp_path.open("wb") as output:
            for chunk in response.iter_content(chunk_size=1024 * 512):
                if chunk:
                    output.write(chunk)


def _download_file(url: str, destination: Path, timeout: int = 120) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_path = destination.with_suffix(destination.suffix + ".part")
    try:
        try:
            _stream_download(url=url, temp_path=temp_path, timeout=timeout, verify=True)
        except requests.exceptions.SSLError:
            _stream_download(url=url, temp_path=temp_path, timeout=timeout, verify=False)
        temp_path.replace(destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def _validate_zip(zip_path: Path) -> None:
    if not zip_path.exists() or zip_path.stat().st_size == 0:
        raise PrismDownloadError(f"Downloaded file is empty or missing: {zip_path.name}")
    if not is_zipfile(zip_path):
        raise PrismDownloadError(f"Downloaded file is not a valid zip archive: {zip_path.name}")


def _extract_tif(zip_path: Path, extract_dir: Path) -> Path:
    extract_dir.mkdir(parents=True, exist_ok=True)
    try:
        with ZipFile(zip_path) as archive:
            tif_members = [member for member in archive.namelist() if member.lower().endswith(".tif")]
            if not tif_members:
                raise PrismDownloadError(f"No GeoTIFF found inside {zip_path.name}")
            tif_member = tif_members[0]
            tif_path = extract_dir / Path(tif_member).name
            if not tif_path.exists():
                tif_path.write_bytes(archive.read(tif_member))
            return tif_path
    except BadZipFile as exc:
        raise PrismDownloadError(f"Corrupted PRISM zip file: {zip_path.name}") from exc


def ensure_prism_asset(variable: str, date_value: date, raw_dir: Path) -> PrismAsset:
    day_token = date_value.strftime("%Y%m%d")
    variable_dir = raw_dir / variable / date_value.strftime("%Y")
    zip_path = variable_dir / f"prism_{variable}_us_30s_{day_token}.zip"
    tif_dir = variable_dir / day_token
    source_url = build_prism_url(variable, date_value)

    if not zip_path.exists():
        _download_file(source_url, zip_path)
    try:
        _validate_zip(zip_path)
    except PrismDownloadError:
        _download_file(source_url, zip_path)
        _validate_zip(zip_path)

    tif_path = _extract_tif(zip_path, tif_dir)
    return PrismAsset(
        variable=variable,
        date_value=date_value,
        zip_path=zip_path,
        tif_path=tif_path,
        source_url=source_url,
    )


def sample_points(asset: PrismAsset, locations: pd.DataFrame) -> pd.DataFrame:
    coordinates = list(zip(locations["longitude"], locations["latitude"]))
    with rasterio.open(asset.tif_path) as dataset:
        xs = [coord[0] for coord in coordinates]
        ys = [coord[1] for coord in coordinates]
        if dataset.crs and dataset.crs.to_string() not in {"EPSG:4326", "EPSG:4269"}:
            xs, ys = transform("EPSG:4326", dataset.crs, xs, ys)
        sampled = list(dataset.sample(zip(xs, ys), masked=True))

    values = []
    for sample in sampled:
        if len(sample) == 0:
            values.append(None)
            continue
        cell_value = sample[0]
        values.append(None if getattr(cell_value, "mask", False) else float(cell_value))

    return pd.DataFrame(
        {
            "location_id": locations["location_id"].values,
            "date": pd.to_datetime(asset.date_value),
            asset.variable: values,
        }
    )
