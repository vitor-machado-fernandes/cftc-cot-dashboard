from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import StringIO
import re
import time

import pandas as pd
import plotly.graph_objects as go
import requests
import urllib3


GLAM_GETTBL_URL = "https://glam1.gsfc.nasa.gov/api/gettbl/v4"
GLAM_CROP_MASK_URL = "https://glam1.gsfc.nasa.gov/api/doc/crop-idname/crop-idname-v17.txt"
GLAM_ADM_L0_URL = "https://glam1.gsfc.nasa.gov/api/doc/shape-id/v1/ADM_L0.txt"
GLAM_ADM_L1_URL = "https://glam1.gsfc.nasa.gov/api/doc/shape-id/v1/ADM_L1.txt"
DEFAULT_GLAM_VERSION = "v17"

SATELLITE_OPTIONS = {
    "VNP": "VNP - S-NPP VIIRS, 8-day, 500 m",
    "VJ1": "VJ1 - NOAA-20 VIIRS, 8-day, 500 m",
    "OS3": "OS3 - Sentinel-3 OLCI, 10-day, 300 m",
}

SEASON_WINDOWS = {
    "United States": {"start_month": 1, "num_months": 12, "note": "Calendar-year cotton monitoring window."},
    "Brazil": {"start_month": 9, "num_months": 12, "note": "Southern Hemisphere crop-year window starting in September."},
    "Australia": {"start_month": 9, "num_months": 12, "note": "Southern Hemisphere crop-year window starting in September."},
    "Mexico": {"start_month": 1, "num_months": 12, "note": "Calendar-year monitoring window."},
    "Turkey": {"start_month": 3, "num_months": 10, "note": "Spring-to-harvest monitoring window."},
    "Zambia": {"start_month": 9, "num_months": 12, "note": "Southern Hemisphere crop-year window starting in September."},
    "Egypt": {"start_month": 3, "num_months": 10, "note": "Spring-to-harvest monitoring window."},
    "Greece": {"start_month": 3, "num_months": 10, "note": "Spring-to-harvest monitoring window."},
}

FALLBACK_L0_IDS = {
    "Australia": "28633",
    "Brazil": "28430",
    "Egypt": "28533",
    "Greece": "28535",
    "Mexico": "28441",
    "Turkey": "28538",
    "United States": "27258",
    "Zambia": "28609",
}

FALLBACK_COTTON_MASKS = [
    ("GDA-CropID-Australia_2021_Cotton", "GDA Australia 2021 Cotton"),
    ("GDA-CropID-Australia_2022_Cotton-v2", "GDA Australia 2022 Cotton"),
    ("GDA-CropID-Australia_2023_Cotton", "GDA Australia 2023 Cotton"),
    ("GDA-CropID-Brazil_2023_Cotton", "GDA Brazil 2023 Cotton"),
    ("GDA-CropID-Brazil_2024_Cotton", "GDA Brazil 2024 Cotton"),
    ("GDA-CropID-Turkey_2019_Cotton", "GDA Turkey 2019 Cotton"),
    ("GDA-CropID-Turkey_2020_Cotton", "GDA Turkey 2020 Cotton"),
    ("GDA-CropID-Turkey_2021_Cotton", "GDA Turkey 2021 Cotton"),
    ("GDA-CropID-Zambia_2025_Cotton", "GDA Zambia 2025 Cotton"),
    ("USDA-ASRC-Brazil_2022_cotton", "ASRC Brazil 2022 Cotton"),
    ("USDA-ASRC-Egypt_2017_cotton", "ASRC Egypt 2017 Cotton"),
    ("USDA-ASRC-Greece_2017_cotton", "ASRC Greece 2017 Cotton"),
    ("USDA-ASRC-Mexico_2020_cotton", "ASRC Mexico 2020 Cotton"),
    ("USDA-NASS-CDL_2018-2023_cotton-50pp", "NASS 2018-2023 50% PR Cotton"),
    ("NASS_2011-2016_cotton", "NASS 2011-2016 50% PR Cotton"),
]


@dataclass(frozen=True)
class GlamMask:
    mask_id: str
    display_name: str
    family: str
    geography: str
    vintage: str
    start_month: int
    num_months: int


@dataclass(frozen=True)
class AdmRegion:
    region_id: str
    country: str
    region_name: str
    level: str


class GlamNdviError(RuntimeError):
    """Raised when GLAM NDVI data cannot be downloaded or parsed."""


def _request_text(url: str, params: list[tuple[str, str]] | None = None, timeout: int = 45) -> str:
    last_error: Exception | None = None
    for attempt in range(3):
        for trust_env in (True, False):
            for verify in (True, False):
                session = requests.Session()
                session.trust_env = trust_env
                try:
                    if not verify:
                        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                    response = session.get(url, params=params, timeout=timeout, verify=verify)
                    response.raise_for_status()
                    return response.text
                except (requests.exceptions.RequestException, requests.exceptions.SSLError) as exc:
                    last_error = exc
                finally:
                    session.close()
        time.sleep(0.75 * (2 ** attempt))
    raise GlamNdviError(f"Could not fetch GLAM data from {url}: {last_error}")


def _split_compact_lines(text: str) -> list[str]:
    return re.findall(r"(?m)(?:^|\s)(\d+\|[^\n\r]*?)(?=\s+\d+\||$)", text)


def _mask_family(mask_id: str) -> str:
    if mask_id.startswith("GDA-CropID-"):
        return "GDA CropID"
    if mask_id.startswith("USDA-ASRC-"):
        return "ASRC"
    if mask_id.startswith("USDA-NASS-CDL") or mask_id.startswith("NASS_"):
        return "USDA/NASS CDL"
    if mask_id.startswith("IFPRI-SPAM"):
        return "IFPRI SPAM"
    return mask_id.split("_", 1)[0].replace("-", " ")


def _display_geography(token: str) -> str:
    replacements = {
        "USA": "United States",
        "US": "United States",
        "UnitedStates": "United States",
        "SouthAmerica": "South America",
        "CentralAsia": "Central Asia",
    }
    if token in replacements:
        return replacements[token]
    return re.sub(r"(?<!^)(?=[A-Z])", " ", token).replace("-", " ").strip()


def _mask_geography(mask_id: str, display_name: str) -> str:
    if mask_id.startswith("USDA-NASS-CDL") or mask_id.startswith("NASS_"):
        return "United States"
    if mask_id.startswith("GDA-CropID-"):
        return _display_geography(mask_id.removeprefix("GDA-CropID-").split("_", 1)[0])
    if mask_id.startswith("USDA-ASRC-"):
        return _display_geography(mask_id.removeprefix("USDA-ASRC-").split("_", 1)[0])
    parts = display_name.split()
    return parts[1] if len(parts) > 2 else "Global"


def _mask_vintage(mask_id: str, display_name: str) -> str:
    year_match = re.search(r"(20\d{2}(?:-\d{4})?)", mask_id)
    if year_match:
        return year_match.group(1)
    display_match = re.search(r"(20\d{2}(?:-\d{4})?)", display_name)
    return display_match.group(1) if display_match else "current"


def _season_for_geography(geography: str) -> dict:
    return SEASON_WINDOWS.get(geography, {"start_month": 1, "num_months": 12, "note": "Calendar-year monitoring window."})


def _mask_records_from_lines(lines: list[str]) -> list[GlamMask]:
    masks: list[GlamMask] = []
    for line in lines:
        if "|" not in line:
            continue
        mask_id, display_name = line.split("|", 1)
        if "cotton" not in f"{mask_id} {display_name}".lower():
            continue
        geography = _mask_geography(mask_id, display_name)
        if geography not in FALLBACK_L0_IDS:
            continue
        season = _season_for_geography(geography)
        masks.append(
            GlamMask(
                mask_id=mask_id,
                display_name=display_name or mask_id,
                family=_mask_family(mask_id),
                geography=geography,
                vintage=_mask_vintage(mask_id, display_name),
                start_month=int(season["start_month"]),
                num_months=int(season["num_months"]),
            )
        )
    return sorted(masks, key=lambda item: (item.family, item.geography, item.vintage, item.display_name))


def get_cotton_masks() -> list[GlamMask]:
    """Return cotton-specific GLAM masks with known ADM geography support."""
    try:
        text = _request_text(GLAM_CROP_MASK_URL)
        rows = text.splitlines()[1:]
        masks = _mask_records_from_lines(rows)
        if masks:
            return masks
    except GlamNdviError:
        pass
    return _mask_records_from_lines([f"{mask_id}|{display_name}" for mask_id, display_name in FALLBACK_COTTON_MASKS])


def _parse_l0_regions(text: str) -> list[AdmRegion]:
    rows = _split_compact_lines(text)
    if not rows:
        rows = text.splitlines()[1:]
    regions: list[AdmRegion] = []
    for row in rows:
        parts = row.split("|")
        if len(parts) >= 2:
            regions.append(AdmRegion(parts[0], parts[1], "National aggregate", "L0"))
    return regions


def _parse_l1_regions(text: str) -> list[AdmRegion]:
    regions: list[AdmRegion] = []
    for row in text.splitlines()[1:]:
        parts = row.split("|")
        if len(parts) >= 3:
            regions.append(AdmRegion(parts[0], parts[1], parts[2], "L1"))
    return regions


def get_adm_regions() -> list[AdmRegion]:
    """Return GLAM ADM national and first-level regions."""
    regions: list[AdmRegion] = []
    try:
        regions.extend(_parse_l0_regions(_request_text(GLAM_ADM_L0_URL)))
        regions.extend(_parse_l1_regions(_request_text(GLAM_ADM_L1_URL)))
    except GlamNdviError:
        regions.extend(AdmRegion(region_id, country, "National aggregate", "L0") for country, region_id in FALLBACK_L0_IDS.items())
    return regions


def regions_for_geography(regions: list[AdmRegion], geography: str) -> list[AdmRegion]:
    matching = [region for region in regions if region.country == geography and region.level == "L0"]
    matching.extend(region for region in regions if region.country == geography and region.level == "L1")
    return sorted(matching, key=lambda item: (item.level != "L0", item.region_name))


def fetch_glam_ndvi(
    *,
    mask_id: str,
    region_id: str,
    years: list[int],
    satellite: str = "VNP",
    start_month: int = 1,
    num_months: int = 12,
    version: str = DEFAULT_GLAM_VERSION,
) -> str:
    """Fetch a seasonal NDVI CSV from GLAM gettbl v4."""
    params: list[tuple[str, str]] = [
        ("version", version),
        ("sat", satellite),
        ("layer", "NDVI"),
        ("mask", mask_id),
        ("shape", "ADM"),
        ("ids", region_id),
        ("start_month", str(start_month)),
        ("num_months", str(num_months)),
        ("ts_type", "seasonal"),
        ("format", "csv"),
    ]
    params.extend(("years", str(year)) for year in years)
    text = _request_text(GLAM_GETTBL_URL, params=params)
    if "ORDINAL DATE" not in text:
        raise GlamNdviError("GLAM returned no seasonal NDVI table for the selected combination.")
    return text


def parse_glam_csv(text: str, region_name: str) -> tuple[pd.DataFrame, dict]:
    """Parse GLAM's text-plus-CSV response into rows and response metadata."""
    lines = text.splitlines()
    header_index = next((index for index, line in enumerate(lines) if line.startswith("ORDINAL DATE,")), None)
    if header_index is None:
        raise GlamNdviError("The GLAM response did not include a CSV table header.")

    metadata = {}
    for line in lines[:header_index]:
        if "," in line:
            key, value = line.split(",", 1)
            metadata[key.strip()] = value.strip()

    csv_text = "\n".join(lines[header_index:])
    frame = pd.read_csv(StringIO(csv_text))
    if frame.empty:
        raise GlamNdviError("The GLAM response table was empty.")

    rename = {
        "ORDINAL DATE": "ordinal_date",
        "START DATE": "start_date",
        "END DATE": "end_date",
        "SOURCE": "source",
        "SAMPLE VALUE": "ndvi",
        "SAMPLE COUNT": "sample_count",
        "MEAN VALUE": "mean_ndvi",
        "MEAN COUNT": "mean_count",
        "ANOM VALUE": "anom_ndvi",
        "MIN VALUE": "min_ndvi",
        "MAX VALUE": "max_ndvi",
    }
    frame = frame.rename(columns=rename)
    for column in ["ndvi", "mean_ndvi", "anom_ndvi", "min_ndvi", "max_ndvi", "sample_count", "mean_count"]:
        if column in frame.columns:
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
    frame["start_date"] = pd.to_datetime(frame["start_date"], errors="coerce")
    frame["end_date"] = pd.to_datetime(frame["end_date"], errors="coerce")
    frame["year"] = frame["start_date"].dt.year
    frame["region"] = region_name
    frame = frame.dropna(subset=["start_date"]).copy()
    return frame, metadata


def add_seasonal_plot_dates(frame: pd.DataFrame, start_month: int) -> pd.DataFrame:
    output = frame.copy()
    season_start_year = output["start_date"].dt.year.where(output["start_date"].dt.month >= start_month, output["start_date"].dt.year - 1)
    season_start = pd.to_datetime(
        {
            "year": season_start_year.astype(int),
            "month": start_month,
            "day": 1,
        }
    )
    output["season_year"] = season_start_year.astype(int)
    output["season_day"] = (output["start_date"] - season_start).dt.days
    output["plot_date"] = pd.Timestamp("2000-01-01") + pd.to_timedelta(output["season_day"], unit="D")
    return output.sort_values(["season_year", "plot_date"])


def build_ndvi_seasonal_chart(frame: pd.DataFrame, selected_years: list[int], start_month: int) -> go.Figure:
    plot_df = add_seasonal_plot_dates(frame, start_month=start_month)
    current_year = max(selected_years)
    figure = go.Figure()

    if {"min_ndvi", "max_ndvi"}.issubset(plot_df.columns):
        baseline = plot_df.drop_duplicates("season_day").sort_values("plot_date")
        if baseline["min_ndvi"].notna().any() and baseline["max_ndvi"].notna().any():
            figure.add_trace(
                go.Scatter(
                    x=baseline["plot_date"],
                    y=baseline["max_ndvi"],
                    line={"width": 0},
                    mode="lines",
                    showlegend=False,
                    hoverinfo="skip",
                )
            )
            figure.add_trace(
                go.Scatter(
                    x=baseline["plot_date"],
                    y=baseline["min_ndvi"],
                    fill="tonexty",
                    fillcolor="rgba(148, 163, 184, 0.22)",
                    line={"width": 0},
                    mode="lines",
                    name="Historical range",
                    hovertemplate="Historical range<br>%{x|%b %d}<br>%{y:.3f}<extra></extra>",
                )
            )

    if "mean_ndvi" in plot_df.columns:
        baseline = plot_df.drop_duplicates("season_day").sort_values("plot_date")
        if baseline["mean_ndvi"].notna().any():
            figure.add_trace(
                go.Scatter(
                    x=baseline["plot_date"],
                    y=baseline["mean_ndvi"],
                    mode="lines",
                    name="Long-term average",
                    line={"color": "#d97706", "width": 3, "dash": "dash"},
                    hovertemplate="Average<br>%{x|%b %d}<br>NDVI: %{y:.3f}<extra></extra>",
                )
            )

    palette = ["#64748b", "#94a3b8", "#3b82f6", "#10b981", "#8b5cf6", "#f97316"]
    for index, year in enumerate(sorted(selected_years)):
        year_df = plot_df.loc[plot_df["year"] == year].sort_values("plot_date")
        if year_df.empty:
            continue
        is_current = year == current_year
        figure.add_trace(
            go.Scatter(
                x=year_df["plot_date"],
                y=year_df["ndvi"],
                mode="lines+markers",
                name=str(year),
                connectgaps=False,
                line={"color": "#0f172a" if is_current else palette[index % len(palette)], "width": 4 if is_current else 2},
                marker={"size": 5 if is_current else 4},
                customdata=year_df[["start_date", "end_date", "region", "sample_count"]],
                hovertemplate=(
                    "%{customdata[2]}<br>"
                    "Year: " + str(year) + "<br>"
                    "Window: %{customdata[0]|%Y-%m-%d} to %{customdata[1]|%Y-%m-%d}<br>"
                    "NDVI: %{y:.3f}<br>"
                    "Sample count: %{customdata[3]:,.0f}<extra></extra>"
                ),
            )
        )

    figure.update_layout(
        height=560,
        hovermode="x unified",
        xaxis_title="Season",
        yaxis_title="NDVI",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "left", "x": 0},
        margin={"l": 20, "r": 20, "t": 70, "b": 40},
    )
    figure.update_xaxes(tickformat="%b", dtick="M1")
    figure.update_yaxes(range=[0, 1])
    return figure
