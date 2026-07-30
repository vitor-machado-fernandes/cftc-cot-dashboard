from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import zlib

import pandas as pd
import requests

from cotton_weather.config import (
    US_COTTON_GINS_FILE,
    US_COTTON_GINS_METADATA_FILE,
    US_COTTON_GINS_PDF,
)
from cotton_weather.counties import load_county_features
from cotton_weather.usda_warehouses import _geometry_center, _normalize_county_name


TEXT_PATTERN = re.compile(
    r"BT /F\d+ (?P<font>\d+(?:\.\d+)?) Tf .*? (?P<x>-?\d+(?:\.\d+)?) (?P<y>-?\d+(?:\.\d+)?) Td "
    r"\((?P<text>(?:\\.|[^\\)])*)\) Tj",
    re.S,
)
CITY_STATE_ZIP_PATTERN = re.compile(
    r"^(?P<city>.+?),\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)$"
)


@dataclass
class TextLine:
    y: float
    max_font: float
    items: list[tuple[float, str]]

    @property
    def min_x(self) -> float:
        return min(item[0] for item in self.items)

    def text(self) -> str:
        parts = [piece for _, piece in sorted(self.items, key=lambda item: item[0])]
        return " ".join(part.strip() for part in parts if part and part.strip()).strip()

    def text_at(self, min_x: float) -> str:
        parts = [text for x, text in sorted(self.items, key=lambda item: item[0]) if x >= min_x]
        return " ".join(part.strip() for part in parts if part and part.strip()).strip()

    def text_between(self, min_x: float, max_x: float) -> str:
        parts = [
            text
            for x, text in sorted(self.items, key=lambda item: item[0])
            if min_x <= x < max_x
        ]
        return " ".join(part.strip() for part in parts if part and part.strip()).strip()


def _decode_pdf_text(value: str) -> str:
    output: list[str] = []
    index = 0
    while index < len(value):
        char = value[index]
        if char != "\\":
            output.append(char)
            index += 1
            continue
        index += 1
        if index >= len(value):
            break
        escaped = value[index]
        mapping = {"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f", "\\": "\\", "(": "(", ")": ")"}
        output.append(mapping.get(escaped, escaped))
        index += 1
    return "".join(output).replace("\r", " ").replace("\n", " ").strip()


def _extract_stream_text(pdf_path: Path) -> list[list[TextLine]]:
    data = pdf_path.read_bytes()
    pages: list[list[TextLine]] = []
    for match in re.finditer(br"stream\r?\n(.*?)\r?\nendstream", data, re.S):
        chunk = match.group(1)
        try:
            decoded = zlib.decompress(chunk).decode("latin1", errors="ignore")
        except Exception:
            continue
        fragments: list[tuple[float, float, float, str]] = []
        for text_match in TEXT_PATTERN.finditer(decoded):
            text = _decode_pdf_text(text_match.group("text"))
            if not text:
                continue
            fragments.append(
                (
                    float(text_match.group("y")),
                    float(text_match.group("x")),
                    float(text_match.group("font")),
                    text,
                )
            )
        if not fragments:
            continue

        lines_by_y: dict[float, list[tuple[float, float, str]]] = {}
        for y, x, font, text in fragments:
            line_key = round(y, 3)
            lines_by_y.setdefault(line_key, []).append((x, font, text))

        page_lines: list[TextLine] = []
        for y, items in sorted(lines_by_y.items(), key=lambda item: item[0], reverse=True):
            sorted_items = sorted(items, key=lambda item: item[0])
            max_font = max(item[1] for item in sorted_items)
            page_lines.append(
                TextLine(
                    y=y,
                    max_font=max_font,
                    items=[(item[0], item[2]) for item in sorted_items],
                )
            )
        pages.append(page_lines)
    return pages


def _is_state_heading(line: TextLine) -> bool:
    text = line.text()
    return line.max_font >= 18 and line.min_x < 100 and "Active Gins Count" not in text


def _is_county_heading(line: TextLine) -> bool:
    text = line.text()
    return 14 <= line.max_font < 18 and line.min_x < 100 and text.upper() == text and not any(char.isdigit() for char in text)


def _is_gin_code_line(line: TextLine) -> bool:
    first_cell = line.text_at(0)
    code_match = re.match(r"^(\d{5})\b", first_cell)
    return bool(code_match) and line.min_x < 50


def _parse_city_state_zip(value: str) -> tuple[str | None, str | None, str | None]:
    match = CITY_STATE_ZIP_PATTERN.match(re.sub(r"\s+", " ", value.strip()))
    if not match:
        return None, None, None
    return match.group("city").strip(), match.group("state").strip(), match.group("zip").strip()


def extract_gin_rows(pdf_path: Path = US_COTTON_GINS_PDF) -> pd.DataFrame:
    if not pdf_path.exists():
        return pd.DataFrame()

    rows: list[dict] = []
    current_state: str | None = None
    current_county: str | None = None

    for page_lines in _extract_stream_text(pdf_path):
        index = 0
        while index < len(page_lines):
            line = page_lines[index]
            if _is_state_heading(line):
                current_state = line.text().strip()
                index += 1
                continue
            if _is_county_heading(line):
                current_county = line.text().strip().title()
                index += 1
                continue
            if not _is_gin_code_line(line):
                index += 1
                continue

            code_match = re.match(r"^(\d{5})\b", line.text_at(0))
            gin_code = code_match.group(1) if code_match else None
            gin_name = line.text_between(90, 330)
            address_parts = [line.text_at(330)]

            look_ahead = index + 1
            while look_ahead < len(page_lines):
                next_line = page_lines[look_ahead]
                if _is_state_heading(next_line) or _is_county_heading(next_line) or _is_gin_code_line(next_line):
                    break
                address_text = next_line.text_at(330)
                if address_text:
                    address_parts.append(address_text)
                look_ahead += 1

            address_parts = [part for part in address_parts if part]
            street_address = address_parts[0] if address_parts else None
            city_state_zip = " ".join(address_parts[1:]).strip() if len(address_parts) > 1 else None
            city, state_abbr, postal_code = _parse_city_state_zip(city_state_zip or "")
            full_address_parts = [part for part in [street_address, city_state_zip] if part]
            full_address = ", ".join(full_address_parts) if full_address_parts else None

            rows.append(
                {
                    "gin_code": gin_code,
                    "gin_name": gin_name,
                    "county": current_county,
                    "state_name": current_state,
                    "state": state_abbr,
                    "street_address": street_address,
                    "city_state_zip": city_state_zip,
                    "city": city,
                    "postal_code": postal_code,
                    "full_address": full_address,
                }
            )
            index = look_ahead

    output = pd.DataFrame(rows)
    if output.empty:
        return output
    output = output.dropna(subset=["gin_code", "gin_name"]).drop_duplicates(subset=["gin_code"], keep="first").reset_index(drop=True)
    return output


def _geocode_one_line_address(session: requests.Session, address: str) -> dict:
    url = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress"
    params = {
        "address": address,
        "benchmark": "Public_AR_Current",
        "format": "json",
    }
    response = session.get(url, params=params, timeout=30)
    response.raise_for_status()
    payload = response.json()
    matches = payload.get("result", {}).get("addressMatches", [])
    if not matches:
        return {"matched_address": None, "latitude": None, "longitude": None, "geocoder_status": "no_match"}
    match = matches[0]
    coordinates = match.get("coordinates", {})
    return {
        "matched_address": match.get("matchedAddress"),
        "latitude": coordinates.get("y"),
        "longitude": coordinates.get("x"),
        "geocoder_status": "match",
        "tiger_line_id": match.get("tigerLine", {}).get("tigerLineId"),
        "side": match.get("tigerLine", {}).get("side"),
    }


def _approximation_candidates(record: dict) -> list[tuple[str, str]]:
    candidates: list[tuple[str, str]] = []
    street = record.get("street_address")
    city = record.get("city")
    state = record.get("state")
    postal_code = record.get("postal_code")
    county = record.get("county")

    zip5 = postal_code[:5] if postal_code else None
    if street and city and state:
        exact = ", ".join(part for part in [street, city, state, postal_code] if part)
        compact = ", ".join(part for part in [street, city, state, zip5] if part)
        candidates.append(("match", exact))
        if compact and compact != exact:
            candidates.append(("match", compact))
    if city and state:
        city_candidate = ", ".join(part for part in [city, state, zip5] if part)
        candidates.append(("city_fallback", city_candidate))
    if county and state:
        candidates.append(("county_fallback", f"{county} County, {state}"))

    deduped: list[tuple[str, str]] = []
    seen: set[str] = set()
    for status, candidate in candidates:
        normalized = re.sub(r"\s+", " ", candidate).strip()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        deduped.append((status, normalized))
    return deduped


def _load_state_county_centers() -> pd.DataFrame:
    rows: list[dict] = []
    for feature in load_county_features():
        properties = feature.get("properties", {})
        latitude, longitude = _geometry_center(feature.get("geometry"))
        county_name = _normalize_county_name(properties.get("NAME"))
        state = str(properties.get("STATE_ABBR", "")).strip().upper()
        if county_name and state and latitude is not None and longitude is not None:
            rows.append(
                {
                    "state": state,
                    "county_name": county_name.title(),
                    "latitude": latitude,
                    "longitude": longitude,
                }
            )
    return pd.DataFrame(rows).drop_duplicates(subset=["state", "county_name"], keep="first")


def build_us_cotton_gin_dataset(source_path: Path = US_COTTON_GINS_PDF) -> pd.DataFrame:
    gins = extract_gin_rows(source_path)
    if gins.empty:
        return gins

    session = requests.Session()
    geocode_rows: list[dict] = []
    for record in gins.to_dict(orient="records"):
        result = {"matched_address": None, "latitude": None, "longitude": None, "geocoder_status": "missing_address"}
        for status, candidate in _approximation_candidates(record):
            result = _geocode_one_line_address(session, candidate)
            if result.get("latitude") is not None and result.get("longitude") is not None:
                result["geocoder_status"] = status
                break
        geocode_rows.append(result)

    geocoded = pd.concat([gins.reset_index(drop=True), pd.DataFrame(geocode_rows)], axis=1)
    geocoded["latitude"] = pd.to_numeric(geocoded["latitude"], errors="coerce")
    geocoded["longitude"] = pd.to_numeric(geocoded["longitude"], errors="coerce")

    county_centers = _load_state_county_centers()
    if not county_centers.empty:
        geocoded = geocoded.merge(
            county_centers.rename(columns={"latitude": "county_latitude", "longitude": "county_longitude"}),
            left_on=["state", "county"],
            right_on=["state", "county_name"],
            how="left",
        )
        missing_coords = geocoded["latitude"].isna() | geocoded["longitude"].isna()
        geocoded.loc[missing_coords, "latitude"] = geocoded.loc[missing_coords, "county_latitude"]
        geocoded.loc[missing_coords, "longitude"] = geocoded.loc[missing_coords, "county_longitude"]
        geocoded.loc[
            missing_coords & geocoded["latitude"].notna() & geocoded["longitude"].notna(),
            "geocoder_status",
        ] = "county_center_fallback"
        geocoded = geocoded.drop(columns=["county_name", "county_latitude", "county_longitude"], errors="ignore")

    US_COTTON_GINS_FILE.parent.mkdir(parents=True, exist_ok=True)
    geocoded.to_csv(US_COTTON_GINS_FILE, index=False)

    metadata = {
        "source_file": str(source_path),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "row_count": int(len(geocoded)),
        "mapped_rows": int(geocoded["latitude"].notna().sum()),
        "exact_match_rows": int((geocoded["geocoder_status"] == "match").sum()),
        "county_fallback_rows": int((geocoded["geocoder_status"] == "county_center_fallback").sum()),
        "unmatched_rows": int(geocoded["latitude"].isna().sum()),
    }
    US_COTTON_GINS_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return geocoded


def load_us_cotton_gins() -> pd.DataFrame:
    if not US_COTTON_GINS_FILE.exists():
        return pd.DataFrame()
    return pd.read_csv(
        US_COTTON_GINS_FILE,
        dtype={
            "gin_code": "string",
            "state": "string",
            "state_name": "string",
            "county": "string",
            "city": "string",
            "postal_code": "string",
            "street_address": "string",
            "full_address": "string",
            "matched_address": "string",
            "geocoder_status": "string",
        },
    )


def load_us_cotton_gins_metadata() -> dict:
    if not US_COTTON_GINS_METADATA_FILE.exists():
        return {}
    return json.loads(US_COTTON_GINS_METADATA_FILE.read_text(encoding="utf-8"))
