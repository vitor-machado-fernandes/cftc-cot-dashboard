from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

from cotton_weather.config import COTTON_RAIL_FILE, COTTON_RAIL_METADATA_FILE, STB_RAW_DIR


STB_FCS_PAGE = "https://www.stb.gov/reports-data/economic-data/freight-commodity-statistics/"
FRA_RAIL_LINES_URL = (
    "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/"
    "NTAD_North_American_Rail_Network_Lines/FeatureServer/0/query"
)

RAILROAD_CONFIG = {
    "BNSF": {
        "stb_name": "BNSF",
        "fra_codes": ["BNSF"],
        "color": "#c2410c",
    },
    "CPKC": {
        "stb_name": "CPKC",
        "fra_codes": ["CP", "CPKC", "KCS", "SOO", "DME", "ICE", "DMVW"],
        "color": "#7c2d12",
    },
    "NS": {
        "stb_name": "NS",
        "fra_codes": ["NS", "NSR"],
        "color": "#111827",
    },
    "UP": {
        "stb_name": "UP",
        "fra_codes": ["UP"],
        "color": "#f59e0b",
    },
}


def _download_text(url: str, timeout: int = 120) -> str:
    response = requests.get(url, timeout=timeout)
    response.raise_for_status()
    return response.text


def _download_bytes(url: str, timeout: int = 240) -> bytes:
    response = requests.get(url, timeout=timeout)
    response.raise_for_status()
    return response.content


def _annual_stb_links_2025() -> dict[str, str]:
    soup = BeautifulSoup(_download_text(STB_FCS_PAGE), "html.parser")
    hrefs = [link.get("href") for link in soup.find_all("a", href=True)]
    lookup = {
        "BNSF": "STB_FCS_PUBLIC_BNSF_2025_Q4",
        "CPKC": "STB_FCS_PUBLIC_CPKC_2025_04",
        "CSX": "STB_FCS_PUBLIC_CSXT_2025_Q4",
        "GTC": "STB_FCS_PUBLIC_GTC_2025_Q4",
        "NS": "STB_FCS_PUBLIC_NS_2025_Q4",
        "UP": "STB_FCS_PUBLIC_UP_2025_Q4",
    }
    output: dict[str, str] = {}
    for railroad, token in lookup.items():
        href = next(item for item in hrefs if item and token in item)
        output[railroad] = "https://www.stb.gov" + href if href.startswith("/") else href
    return output


def _save_stb_workbooks(target_dir: Path) -> dict[str, Path]:
    target_dir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}
    for railroad, url in _annual_stb_links_2025().items():
        path = target_dir / Path(url).name
        if not path.exists():
            path.write_bytes(_download_bytes(url))
        paths[railroad] = path
    return paths


def _extract_cotton_tons(workbook_path: Path) -> float:
    df = pd.read_excel(workbook_path, sheet_name="QCS_data")
    stcc_code = df.iloc[:, 0].astype(str).str.strip()
    cotton_row = df.loc[stcc_code == "0112---"].copy()
    if cotton_row.empty:
        return 0.0
    ton_columns = [column for column in df.columns if "tons" in str(column).lower()]
    return float(cotton_row[ton_columns].fillna(0).sum(axis=1).iloc[0])


def load_stb_cotton_railroad_summary_2025() -> pd.DataFrame:
    workbook_paths = _save_stb_workbooks(STB_RAW_DIR / "freight_commodity_statistics")
    rows: list[dict] = []
    for railroad, workbook_path in workbook_paths.items():
        tons = _extract_cotton_tons(workbook_path)
        rows.append(
            {
                "railroad": railroad,
                "cotton_tons_2025": tons,
                "source_file": str(workbook_path),
            }
        )
    return pd.DataFrame(rows).sort_values("cotton_tons_2025", ascending=False).reset_index(drop=True)


def _fra_where_clause(codes: list[str]) -> str:
    fields = ["RROWNER1", "RROWNER2", "RROWNER3"] + [f"TRKRGHTS{idx}" for idx in range(1, 10)]
    quoted_codes = ", ".join(f"'{code}'" for code in codes)
    tests = [f"{field} IN ({quoted_codes})" for field in fields]
    return "(" + " OR ".join(tests) + ") AND COUNTRY='US'"


def _fetch_fra_features(codes: list[str]) -> list[dict]:
    where = _fra_where_clause(codes)
    offset = 0
    page_size = 2000
    features: list[dict] = []
    while True:
        response = requests.get(
            FRA_RAIL_LINES_URL,
            params={
                "where": where,
                "outFields": "RROWNER1,RROWNER2,RROWNER3,TRKRGHTS1,TRKRGHTS2,TRKRGHTS3,TRKRGHTS4,TRKRGHTS5,TRKRGHTS6,TRKRGHTS7,TRKRGHTS8,TRKRGHTS9,STATEAB,MILES",
                "returnGeometry": "true",
                "f": "geojson",
                "resultOffset": offset,
                "resultRecordCount": page_size,
            },
            timeout=240,
        )
        response.raise_for_status()
        payload = response.json()
        batch = payload.get("features", [])
        if not batch:
            break
        features.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size
    return features


def _multiline_parts(geometry: dict | None) -> list[list[list[float]]]:
    if not geometry:
        return []
    geometry_type = geometry.get("type")
    coordinates = geometry.get("coordinates", [])
    if geometry_type == "LineString":
        return [coordinates]
    if geometry_type == "MultiLineString":
        return coordinates
    return []


def _simplify_line(line: list[list[float]], max_points: int = 18) -> list[list[float]]:
    if len(line) <= max_points:
        return [[round(float(point[0]), 4), round(float(point[1]), 4)] for point in line if len(point) >= 2]
    stride = max(1, len(line) // (max_points - 1))
    simplified = [line[index] for index in range(0, len(line), stride)]
    if simplified[-1] != line[-1]:
        simplified.append(line[-1])
    return [[round(float(point[0]), 4), round(float(point[1]), 4)] for point in simplified if len(point) >= 2]


def build_cotton_rail_layer_2025() -> tuple[dict, dict]:
    summary = load_stb_cotton_railroad_summary_2025()
    active = summary.loc[summary["cotton_tons_2025"] > 0].copy()
    if active.empty:
        geojson = {"type": "FeatureCollection", "features": []}
        metadata = {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "row_count": 0,
            "railroads": [],
        }
        COTTON_RAIL_FILE.write_text(json.dumps(geojson), encoding="utf-8")
        COTTON_RAIL_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
        return geojson, metadata

    max_tons = float(active["cotton_tons_2025"].max())
    features: list[dict] = []
    for record in active.to_dict(orient="records"):
        railroad = str(record["railroad"])
        config = RAILROAD_CONFIG.get(railroad)
        if not config:
            continue
        tons = float(record["cotton_tons_2025"])
        line_width = 1.5 + (8.0 * (tons / max_tons if max_tons else 0.0))
        multiline_coordinates: list[list[list[float]]] = []
        for feature in _fetch_fra_features(config["fra_codes"]):
            for line in _multiline_parts(feature.get("geometry")):
                simplified = _simplify_line(line)
                if len(simplified) >= 2:
                    multiline_coordinates.append(simplified)
        if not multiline_coordinates:
            continue
        features.append(
            {
                "type": "Feature",
                "properties": {
                    "railroad": railroad,
                    "cotton_tons_2025": tons,
                    "line_width": line_width,
                    "line_color": config["color"],
                    "line_count": len(multiline_coordinates),
                },
                "geometry": {
                    "type": "MultiLineString",
                    "coordinates": multiline_coordinates,
                },
            }
        )

    geojson = {"type": "FeatureCollection", "features": features}
    metadata = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "row_count": len(features),
        "railroads": active.to_dict(orient="records"),
        "caveat": (
            "Best-available approximation using STB 2025 Freight Commodity Statistics by Class I railroad "
            "and FRA rail network ownership/trackage-rights geometry. Thickness reflects railroad-system raw cotton tons, "
            "not verified segment-level traffic."
        ),
    }
    COTTON_RAIL_FILE.parent.mkdir(parents=True, exist_ok=True)
    COTTON_RAIL_FILE.write_text(json.dumps(geojson), encoding="utf-8")
    COTTON_RAIL_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return geojson, metadata


def load_cotton_rail_layer_2025() -> dict:
    if not COTTON_RAIL_FILE.exists():
        return {}
    return json.loads(COTTON_RAIL_FILE.read_text(encoding="utf-8"))


def load_cotton_rail_metadata_2025() -> dict:
    if not COTTON_RAIL_METADATA_FILE.exists():
        return {}
    return json.loads(COTTON_RAIL_METADATA_FILE.read_text(encoding="utf-8"))
