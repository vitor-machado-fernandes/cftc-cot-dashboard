from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
from urllib.parse import urljoin

from bs4 import BeautifulSoup
import pandas as pd
import requests

from cotton_weather.config import PROCESSED_DIR, REFERENCE_DIR
from cotton_weather.counties import load_county_weights


COUNTY_IRRIGATION_SOURCE_FILE = REFERENCE_DIR / "cotton_irrigated_share_by_county_2022.csv"
COUNTY_IRRIGATION_METADATA_FILE = PROCESSED_DIR / "cotton_irrigated_share_by_county_2022_metadata.json"
USDA_RAW_DIR = REFERENCE_DIR.parent / "raw" / "usda"
USDA_WORKBOOK_FILE = USDA_RAW_DIR / "nass_ag_census_web_maps_2022.xlsx"
USDA_AG_CENSUS_WEB_MAPS_URL = "https://www.nass.usda.gov/Publications/AgCensus/2022/Online_Resources/Ag_Census_Web_Maps/Data_download/index.php"
USDA_AG_CENSUS_DOC_URL = "https://www.nass.usda.gov/Publications/AgCensus/2022/Online_Resources/Ag_Census_Web_Maps/Data_download/NASSAgCensusWebMaps2022_Doc.pdf"
USDA_MAP_ID = "y22_M121"
USDA_SHEET_NAME = "Crops and Plants"


def file_signature(path: Path) -> tuple:
    if not path.exists():
        return (str(path), 0, 0)
    stat = path.stat()
    return (str(path), int(stat.st_size), int(stat.st_mtime))


def _normalize_geoid(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.zfill(5)


def _scrape_workbook_url(download_page_url: str = USDA_AG_CENSUS_WEB_MAPS_URL) -> str:
    response = requests.get(download_page_url, timeout=120)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if href.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return urljoin(download_page_url, href)
    raise RuntimeError("Could not find the USDA Ag Census Web Maps workbook link on the official download page.")


def download_county_irrigation_workbook(force: bool = False) -> Path:
    if USDA_WORKBOOK_FILE.exists() and not force:
        return USDA_WORKBOOK_FILE

    workbook_url = _scrape_workbook_url()
    USDA_WORKBOOK_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp_path = USDA_WORKBOOK_FILE.with_suffix(".part")
    try:
        with requests.get(workbook_url, stream=True, timeout=300) as response:
            response.raise_for_status()
            with temp_path.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        output.write(chunk)
        temp_path.replace(USDA_WORKBOOK_FILE)
    finally:
        if temp_path.exists():
            temp_path.unlink()
    return USDA_WORKBOOK_FILE


def build_county_irrigation_share_source(force_download: bool = False) -> pd.DataFrame:
    workbook_path = download_county_irrigation_workbook(force=force_download)
    data = pd.read_excel(workbook_path, sheet_name=USDA_SHEET_NAME)
    geoid_column = "FIPSTEXT" if "FIPSTEXT" in data.columns else "FIPS"
    numeric_column = f"{USDA_MAP_ID}_valueNumeric"
    text_column = f"{USDA_MAP_ID}_valueText"
    required_columns = {geoid_column, numeric_column}
    missing = sorted(column for column in required_columns if column not in data.columns)
    if missing:
        missing_text = ", ".join(missing)
        raise ValueError(f"The official USDA workbook is missing expected columns: {missing_text}")

    output = pd.DataFrame(
        {
            "geoid": _normalize_geoid(data[geoid_column]),
            "irrigated_share": pd.to_numeric(data[numeric_column], errors="coerce") / 100.0,
            "source_text": data[text_column] if text_column in data.columns else pd.NA,
        }
    )
    output["irrigated_share"] = output["irrigated_share"].clip(lower=0.0, upper=1.0)
    output = output.loc[output["geoid"].notna()].drop_duplicates(subset=["geoid"]).reset_index(drop=True)

    COUNTY_IRRIGATION_SOURCE_FILE.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(COUNTY_IRRIGATION_SOURCE_FILE, index=False)

    metadata = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_workbook": str(workbook_path),
        "source_url": USDA_AG_CENSUS_WEB_MAPS_URL,
        "source_doc_url": USDA_AG_CENSUS_DOC_URL,
        "source_map_id": USDA_MAP_ID,
        "row_count": int(len(output)),
    }
    COUNTY_IRRIGATION_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return output


def load_county_irrigation_share_source() -> pd.DataFrame:
    if not COUNTY_IRRIGATION_SOURCE_FILE.exists():
        return pd.DataFrame()
    output = pd.read_csv(COUNTY_IRRIGATION_SOURCE_FILE, dtype={"geoid": str})
    output["geoid"] = _normalize_geoid(output["geoid"])
    output["irrigated_share"] = pd.to_numeric(output["irrigated_share"], errors="coerce").clip(lower=0.0, upper=1.0)
    return output


def load_county_irrigation_metadata() -> dict:
    if not COUNTY_IRRIGATION_METADATA_FILE.exists():
        return {}
    return json.loads(COUNTY_IRRIGATION_METADATA_FILE.read_text(encoding="utf-8"))


def summarize_state_irrigated_acres(year: int = 2024) -> pd.DataFrame:
    county_weights = load_county_weights()
    county_share = load_county_irrigation_share_source()
    if county_weights.empty or county_share.empty:
        return pd.DataFrame()

    working = county_weights.loc[county_weights["year"] == year].copy()
    if working.empty:
        return pd.DataFrame()

    working = working.merge(
        county_share[["geoid", "irrigated_share"]],
        on="geoid",
        how="left",
    )
    working["irrigated_share"] = working["irrigated_share"].fillna(0.0)
    working["irrigated_cotton_area_acres_est"] = working["cotton_area_acres_est"] * working["irrigated_share"]
    working["rainfed_cotton_area_acres_est"] = working["cotton_area_acres_est"] - working["irrigated_cotton_area_acres_est"]

    summary = (
        working.groupby("state", as_index=False)
        .agg(
            cotton_area_acres_est=("cotton_area_acres_est", "sum"),
            irrigated_cotton_area_acres_est=("irrigated_cotton_area_acres_est", "sum"),
            rainfed_cotton_area_acres_est=("rainfed_cotton_area_acres_est", "sum"),
        )
        .sort_values("cotton_area_acres_est", ascending=False)
        .reset_index(drop=True)
    )
    summary["irrigated_share"] = summary["irrigated_cotton_area_acres_est"] / summary["cotton_area_acres_est"]
    return summary
