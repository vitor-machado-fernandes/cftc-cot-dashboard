from __future__ import annotations

import json

import pandas as pd

from cotton_weather.config import COTTON_COUNTIES_FILE, USDA_WAREHOUSE_FILE


WAREHOUSE_COLUMNS = [
    "current_date",
    "ingest_date",
    "warehouse_name",
    "business_entity",
    "city",
    "county",
    "state",
    "capacity_bales",
    "license_number",
    "license_type",
    "warehouse_status",
    "ccc_status",
    "commodity",
    "county_name",
    "latitude",
    "longitude",
]


def _flatten_coordinates(coordinates: list) -> tuple[list[float], list[float]]:
    longitudes: list[float] = []
    latitudes: list[float] = []

    def visit(node) -> None:
        if not isinstance(node, (list, tuple)) or not node:
            return
        first = node[0]
        if isinstance(first, (int, float)) and len(node) >= 2:
            longitudes.append(float(node[0]))
            latitudes.append(float(node[1]))
            return
        for item in node:
            visit(item)

    visit(coordinates)
    return longitudes, latitudes


def _geometry_center(geometry: dict | None) -> tuple[float | None, float | None]:
    if not geometry:
        return None, None
    longitudes, latitudes = _flatten_coordinates(geometry.get("coordinates", []))
    if not longitudes or not latitudes:
        return None, None
    return (min(latitudes) + max(latitudes)) / 2.0, (min(longitudes) + max(longitudes)) / 2.0


def _normalize_county_name(value: object) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().upper()
    if not text:
        return None
    if "," in text:
        text = text.split(",", 1)[0].strip()
    text = text.replace("SAINT ", "ST. ")
    return text


def _parse_capacity(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _load_county_centers() -> pd.DataFrame:
    if not COTTON_COUNTIES_FILE.exists():
        return pd.DataFrame(columns=["state", "county_name", "latitude", "longitude"])

    with COTTON_COUNTIES_FILE.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    rows: list[dict] = []
    for feature in payload.get("features", []):
        properties = feature.get("properties", {})
        latitude, longitude = _geometry_center(feature.get("geometry"))
        if latitude is None or longitude is None:
            continue
        county_name = _normalize_county_name(properties.get("NAME"))
        state = str(properties.get("STATE_ABBR", "")).strip().upper()
        if not county_name or not state:
            continue
        rows.append(
            {
                "state": state,
                "county_name": county_name,
                "latitude": latitude,
                "longitude": longitude,
            }
        )
    return pd.DataFrame(rows).drop_duplicates(subset=["state", "county_name"], keep="first")


def load_usda_warehouse_data() -> pd.DataFrame:
    if not USDA_WAREHOUSE_FILE.exists():
        return pd.DataFrame(columns=WAREHOUSE_COLUMNS)

    raw = pd.read_csv(USDA_WAREHOUSE_FILE)
    if raw.empty:
        return pd.DataFrame(columns=WAREHOUSE_COLUMNS)

    warehouses = pd.DataFrame(
        {
            "current_date": raw.get("Current Date"),
            "ingest_date": raw.get("Ingest Date"),
            "warehouse_name": raw.get("Warehouse Name"),
            "business_entity": raw.get("Business Entity"),
            "city": raw.get("City"),
            "county": raw.get("County"),
            "state": raw.get("State"),
            "capacity_bales": raw.get("Capacity"),
            "license_number": raw.get("License Number"),
            "license_type": raw.get("License Type"),
            "warehouse_status": raw.get("Warehouse Status"),
            "ccc_status": raw.get("CCC Status"),
            "commodity": raw.get("Commodity*").fillna(raw.get("Commodity")),
        }
    )
    warehouses["state"] = warehouses["state"].astype("string").str.strip().str.upper()
    warehouses["current_date"] = warehouses["current_date"].astype("string").str.strip()
    warehouses["ingest_date"] = warehouses["ingest_date"].astype("string").str.strip()
    warehouses["commodity"] = warehouses["commodity"].astype("string").str.strip()
    warehouses["county_name"] = warehouses["county"].apply(_normalize_county_name)
    warehouses["capacity_bales"] = warehouses["capacity_bales"].apply(_parse_capacity)
    warehouses = warehouses.loc[
        warehouses["commodity"].str.contains("COTTON", case=False, na=False)
        & warehouses["state"].notna()
        & warehouses["county_name"].notna()
        & warehouses["capacity_bales"].notna()
        & (warehouses["capacity_bales"] > 0)
    ].copy()
    if warehouses.empty:
        return pd.DataFrame(columns=WAREHOUSE_COLUMNS)

    county_centers = _load_county_centers()
    if county_centers.empty:
        warehouses["latitude"] = pd.NA
        warehouses["longitude"] = pd.NA
    else:
        warehouses = warehouses.merge(county_centers, on=["state", "county_name"], how="left")

    warehouses["city"] = warehouses["city"].astype("string").str.strip().str.title()
    warehouses["county"] = warehouses["county_name"].astype("string").str.title()
    warehouses["warehouse_name"] = warehouses["warehouse_name"].astype("string").str.strip()
    warehouses["business_entity"] = warehouses["business_entity"].astype("string").str.strip()
    warehouses["license_number"] = warehouses["license_number"].astype("string").str.strip()
    warehouses["license_type"] = warehouses["license_type"].astype("string").str.strip()
    warehouses["warehouse_status"] = warehouses["warehouse_status"].astype("string").str.strip()
    warehouses["ccc_status"] = warehouses["ccc_status"].astype("string").str.strip()
    warehouses["commodity"] = warehouses["commodity"].astype("string").str.strip()
    warehouses["capacity_bales"] = warehouses["capacity_bales"].astype(float)

    return warehouses[WAREHOUSE_COLUMNS].sort_values(
        ["capacity_bales", "state", "warehouse_name"],
        ascending=[False, True, True],
    ).reset_index(drop=True)
