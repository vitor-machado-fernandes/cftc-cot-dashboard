from __future__ import annotations

import csv
from datetime import datetime, timezone
import json
from pathlib import Path
import time

import pandas as pd
import requests

from cotton_weather.config import (
    COTTON_PORT_EXPORTS_FILE,
    COTTON_PORTS_FILE,
    COTTON_PORTS_METADATA_FILE,
)


PORT_QUERY_MAP = {
    "Savannah, GA (District)": "Port of Savannah, Georgia, United States",
    "Los Angeles, CA (District)": "Port of Los Angeles, California, United States",
    "Houston-Galveston, TX (District)": "Port of Houston, Texas, United States",
    "Norfolk, VA (District)": "Norfolk International Terminals, Norfolk, Virginia, United States",
    "Mobile, AL (District)": "Port of Mobile, Alabama, United States",
    "Charleston, SC (District)": "Port of Charleston, South Carolina, United States",
    "New Orleans, LA (District)": "Port of New Orleans, Louisiana, United States",
    "Wilmington, NC (District)": "Port of Wilmington, North Carolina, United States",
    "Tampa, FL (District)": "Port Tampa Bay, Florida, United States",
    "San Francisco, CA (District)": "Port of Oakland, California, United States",
    "Miami, FL (District)": "Port of Miami, Florida, United States",
    "Seattle, WA (District)": "Port of Seattle, Washington, United States",
    "San Diego, CA (District)": "Port of San Diego, California, United States",
    "New York City, NY (District)": "Port Newark Container Terminal, Newark, New Jersey, United States",
    "Detroit, MI (District)": "Port of Detroit, Michigan, United States",
    "Duluth, MN (District)": "Port of Duluth, Minnesota, United States",
    "San Juan, PR (District)": "Port of San Juan, Puerto Rico, United States",
}


def _parse_export_rows(source_path: Path = COTTON_PORT_EXPORTS_FILE) -> tuple[pd.DataFrame, str | None]:
    if not source_path.exists():
        return pd.DataFrame(), None

    with source_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        next(reader, None)
        current_date_row = next(reader, None)
        header = next(reader, None)
        if not header:
            return pd.DataFrame(), None
        header = [column for column in header if column != ""]

        rows: list[dict] = []
        for row in reader:
            if len(row) < len(header):
                continue
            rows.append(dict(zip(header, row[: len(header)])))

    current_date = current_date_row[0] if current_date_row else None
    frame = pd.DataFrame(rows)
    return frame, current_date


def _parse_numeric(value: object) -> float:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def load_port_export_summary_2025(source_path: Path = COTTON_PORT_EXPORTS_FILE) -> tuple[pd.DataFrame, str | None]:
    frame, current_date = _parse_export_rows(source_path)
    if frame.empty:
        return frame, current_date

    frame = frame.loc[frame["Time"] == "2025"].copy()
    frame["vessel_total_exports_kg"] = frame["Vessel Total Exports SWT (kg)"].apply(_parse_numeric)
    frame["containerized_exports_kg"] = frame["Containerized Vessel Total Exports SWT (kg)"].apply(_parse_numeric)
    frame = frame.loc[frame["vessel_total_exports_kg"] > 0].copy()
    frame = frame.sort_values("vessel_total_exports_kg", ascending=False).reset_index(drop=True)
    return frame, current_date


def _geocode_query(query: str, session: requests.Session) -> dict:
    response = session.get(
        "https://nominatim.openstreetmap.org/search",
        params={"q": query, "format": "jsonv2", "limit": 1},
        headers={"User-Agent": "weather-monitor/1.0"},
        timeout=60,
    )
    response.raise_for_status()
    data = response.json()
    if not data:
        return {"matched_name": None, "matched_display_name": None, "latitude": None, "longitude": None}
    top = data[0]
    return {
        "matched_name": top.get("name"),
        "matched_display_name": top.get("display_name"),
        "latitude": float(top["lat"]) if top.get("lat") else None,
        "longitude": float(top["lon"]) if top.get("lon") else None,
    }


def build_cotton_ports_2025(source_path: Path = COTTON_PORT_EXPORTS_FILE) -> pd.DataFrame:
    frame, current_date = load_port_export_summary_2025(source_path)
    if frame.empty:
        return frame

    session = requests.Session()
    geocoded_rows: list[dict] = []
    for _, row in frame.iterrows():
        port = row["Port"]
        query = PORT_QUERY_MAP.get(port, port.replace(" (District)", "") + ", United States")
        result = _geocode_query(query, session)
        geocoded_rows.append(
            {
                "port": port,
                "query_used": query,
                "matched_name": result.get("matched_name"),
                "matched_display_name": result.get("matched_display_name"),
                "latitude": result.get("latitude"),
                "longitude": result.get("longitude"),
            }
        )
        time.sleep(1.0)

    frame = frame.rename(columns={"Port": "port", "Commodity": "commodity", "Country": "country", "Time": "time"})
    output = frame.merge(pd.DataFrame(geocoded_rows), on="port", how="left")
    output["vessel_total_exports_kg"] = pd.to_numeric(output["vessel_total_exports_kg"], errors="coerce")
    output["containerized_exports_kg"] = pd.to_numeric(output["containerized_exports_kg"], errors="coerce")

    COTTON_PORTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(COTTON_PORTS_FILE, index=False)

    metadata = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source_file": str(source_path),
        "source_current_date": current_date,
        "row_count": int(len(output)),
        "mapped_rows": int(output["latitude"].notna().sum()),
        "unmapped_rows": int(output["latitude"].isna().sum()),
        "note": "Port district locations are approximated using the main port facility or terminal associated with each district label.",
    }
    COTTON_PORTS_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return output


def load_cotton_ports_2025() -> pd.DataFrame:
    if not COTTON_PORTS_FILE.exists():
        return pd.DataFrame()
    return pd.read_csv(COTTON_PORTS_FILE)


def load_cotton_ports_metadata_2025() -> dict:
    if not COTTON_PORTS_METADATA_FILE.exists():
        return {}
    return json.loads(COTTON_PORTS_METADATA_FILE.read_text(encoding="utf-8"))
