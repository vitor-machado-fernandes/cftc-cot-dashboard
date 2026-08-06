from __future__ import annotations

from pathlib import Path
from urllib.parse import urljoin

import requests
import urllib3
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from cotton_weather.config import (
    PROCESSED_DIR,
    RAW_DIR,
    STATE_PRECIP_FILE,
    STATE_PRECIP_METADATA_FILE,
)
from cotton_weather.forecast_qpf import (
    WPC_QPF_IMAGE_URLS,
    ensure_wpc_qpf_image,
    load_wpc_qpf_geojson,
)
from cotton_weather.precip_maps import (
    available_map_windows,
    available_cached_map_windows,
    load_precipitation_map_preview_cached,
)
from cotton_weather.state_precip import (
    load_state_precipitation_data,
    load_state_precipitation_metadata,
    load_state_precipitation_progress,
)
from cotton_weather.state_precip_section import render_cotton_state_precipitation


FORECAST_WINDOW = "Next 7 days"
FOOTPRINT_STATES = "AL-AR-AZ-CA-FL-GA-KS-LA-MO-MS-NC-NM-OK-SC-TN-TX-VA"
DROUGHT_MONITOR_URL = "https://droughtmonitor.unl.edu/"


def _file_signature(path: Path) -> tuple:
    if not path.exists():
        return (str(path), 0, 0)
    stat = path.stat()
    return (str(path), int(stat.st_size), int(stat.st_mtime))


def _latest_prism_signature(variable: str = "ppt", limit: int = 21) -> tuple:
    variable_dir = RAW_DIR / variable
    if not variable_dir.exists():
        return tuple()
    tif_paths = sorted(
        variable_dir.glob("*/*/prism_*.tif"),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    return tuple((path.name, int(path.stat().st_size), int(path.stat().st_mtime)) for path in tif_paths[:limit])


def _latest_prism_preview_signature(limit: int = 50) -> tuple:
    cache_dir = PROCESSED_DIR / "prism_map_previews"
    if not cache_dir.exists():
        return tuple()
    preview_paths = sorted(cache_dir.glob("*.parquet"), key=lambda item: item.stat().st_mtime, reverse=True)
    return tuple((path.name, int(path.stat().st_size), int(path.stat().st_mtime)) for path in preview_paths[:limit])


def _latest_preview_cache(irrigation_mode: str) -> Path | None:
    cache_dir = PROCESSED_DIR / "cdl_preview_cache"
    if not cache_dir.exists():
        return None
    pattern = f"cotton_preview_{FOOTPRINT_STATES}_12000_50_{irrigation_mode}_*.parquet"
    candidates = sorted(cache_dir.glob(pattern), key=lambda item: item.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def _preview_signature(irrigation_mode: str) -> tuple:
    return _file_signature(_latest_preview_cache(irrigation_mode) or Path("__missing__"))


def _state_precip_signature() -> tuple:
    return (
        _file_signature(STATE_PRECIP_FILE),
        _file_signature(STATE_PRECIP_METADATA_FILE),
    )


@st.cache_data(ttl=1800)
def _load_forecast_map(product_label: str) -> tuple[dict, pd.DataFrame, dict]:
    return load_wpc_qpf_geojson(product_label)


@st.cache_data(ttl=1800)
def _load_forecast_image(image_url: str) -> bytes:
    return ensure_wpc_qpf_image(image_url).read_bytes()


def _get_with_ssl_fallback(url: str, timeout: int = 30) -> requests.Response:
    last_error: Exception | None = None
    for trust_env in (True, False):
        for verify in (True, False):
            session = requests.Session()
            session.trust_env = trust_env
            try:
                if not verify:
                    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
                response = session.get(url, timeout=timeout, verify=verify)
                response.raise_for_status()
                return response
            except (requests.exceptions.ProxyError, requests.exceptions.SSLError) as exc:
                last_error = exc
            finally:
                session.close()

    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Could not fetch {url}")


@st.cache_data(ttl=21600)
def _load_drought_monitor_image() -> tuple[bytes, dict]:
    response = _get_with_ssl_fallback(DROUGHT_MONITOR_URL)

    soup = BeautifulSoup(response.text, "html.parser")
    png_href = None
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "/data/png/" in href and href.lower().endswith("_usdm.png"):
            png_href = href
            break

    if png_href is None:
        raise ValueError("Could not find the current U.S. Drought Monitor PNG link.")

    image_url = urljoin(DROUGHT_MONITOR_URL, png_href)
    image_response = _get_with_ssl_fallback(image_url)

    page_text = soup.get_text("\n", strip=True)
    released = next((line.replace("Map released:", "").strip() for line in page_text.splitlines() if line.startswith("Map released:")), "")
    valid = next((line.replace("Data valid:", "").strip() for line in page_text.splitlines() if line.startswith("Data valid:")), "")
    return image_response.content, {"image_url": image_url, "released": released, "valid": valid}


@st.cache_data(ttl=300)
def _load_precip_window_dates(window_days: int, _preview_signature: tuple, _prism_signature: tuple) -> list[str]:
    cached_dates = available_cached_map_windows(window_days=window_days)
    raw_dates = available_map_windows(window_days=window_days)
    return [value.isoformat() for value in sorted(set(cached_dates).union(raw_dates))]


@st.cache_data(ttl=300)
def _load_precip_map_preview(
    window_days: int,
    end_date_iso: str,
    _preview_signature: tuple,
    _prism_signature: tuple,
) -> tuple[pd.DataFrame, dict]:
    end_date = pd.Timestamp(end_date_iso).date()
    cached_dates = set(available_cached_map_windows(window_days=window_days))
    return load_precipitation_map_preview_cached(
        window_days=window_days,
        end_date=end_date,
        build_if_missing=end_date not in cached_dates,
    )


@st.cache_data(ttl=900)
def _load_cached_footprint_preview(irrigation_mode: str, _signature: tuple) -> pd.DataFrame:
    cache_path = _latest_preview_cache(irrigation_mode)
    if cache_path is None:
        return pd.DataFrame()
    return pd.read_parquet(cache_path)


@st.cache_data(ttl=900)
def _load_state_precip_data(_signature: tuple) -> pd.DataFrame:
    return load_state_precipitation_data()


@st.cache_data(ttl=900)
def _load_state_precip_meta(_signature: tuple) -> dict:
    return load_state_precipitation_metadata()


@st.cache_data(ttl=30)
def _load_state_precip_progress() -> dict:
    return load_state_precipitation_progress()


def _inches_to_mm(value: float | int | None) -> float | None:
    if value is None or pd.isna(value):
        return None
    return float(value) * 25.4


def _interpolate_hex_color(start_hex: str, end_hex: str, fraction: float | None) -> str:
    if fraction is None or pd.isna(fraction):
        return start_hex
    clipped = max(0.0, min(1.0, float(fraction)))
    start = tuple(int(start_hex[index:index + 2], 16) for index in (1, 3, 5))
    end = tuple(int(end_hex[index:index + 2], 16) for index in (1, 3, 5))
    blended = tuple(int(round(start[idx] + (end[idx] - start[idx]) * clipped)) for idx in range(3))
    return "#{:02x}{:02x}{:02x}".format(*blended)


def _add_plain_footprint_overlay(base_figure, preview_df: pd.DataFrame):
    overlay = px.scatter_geo(
        preview_df,
        lat="latitude",
        lon="longitude",
        scope="usa",
        opacity=0.12,
    )
    overlay.update_traces(
        marker={"size": 2.2, "color": "#111111"},
        hovertemplate="Cotton footprint<br>State: %{customdata[0]}<extra></extra>",
        customdata=np.stack([preview_df["state"]], axis=-1),
        showlegend=False,
    )
    for trace in overlay.data:
        base_figure.add_trace(trace)


def _add_county_share_footprint_overlay(base_figure, preview_df: pd.DataFrame):
    known = preview_df.loc[preview_df["irrigation_value"].notna()].copy()
    unknown = preview_df.loc[preview_df["irrigation_value"].isna()].copy()
    if not known.empty:
        known["overlay_color"] = known["irrigation_value"].apply(
            lambda value: _interpolate_hex_color("#facc15", "#b91c1c", value)
        )
        overlay = px.scatter_geo(
            known,
            lat="latitude",
            lon="longitude",
            scope="usa",
            opacity=0.14,
            hover_data={
                "state": True,
                "county_geoid": True,
                "irrigation_value": ":.0%",
                "latitude": False,
                "longitude": False,
            },
        )
        overlay.update_traces(
            marker={"size": 2.4, "color": known["overlay_color"].tolist()},
            showlegend=False,
        )
        for trace in overlay.data:
            base_figure.add_trace(trace)
    if not unknown.empty:
        overlay_unknown = px.scatter_geo(
            unknown,
            lat="latitude",
            lon="longitude",
            scope="usa",
            opacity=0.14,
        )
        overlay_unknown.update_traces(
            marker={"size": 2.4, "color": "#facc15"},
            hovertemplate="State: %{customdata[0]}<br>County GEOID: %{customdata[1]}<br>County irrigated share: unavailable<extra></extra>",
            customdata=np.stack([unknown["state"], unknown["county_geoid"]], axis=-1),
            showlegend=False,
        )
        for trace in overlay_unknown.data:
            base_figure.add_trace(trace)


def _render_forecast_map():
    st.subheader("Forecast Map")
    image_url = WPC_QPF_IMAGE_URLS[FORECAST_WINDOW]
    try:
        forecast_image = _load_forecast_image(image_url)
    except Exception as exc:
        st.error(f"Could not load the NOAA WPC 7-day forecast map: {exc}")
        return

    forecast_records = pd.DataFrame()
    forecast_meta = {}
    try:
        _, forecast_records, forecast_meta = _load_forecast_map(FORECAST_WINDOW)
    except Exception:
        pass

    st.image(forecast_image, use_container_width=True)
    max_qpf_in = float(forecast_records["qpf"].max()) if not forecast_records.empty else None
    max_qpf_mm = _inches_to_mm(max_qpf_in)
    caption = "NOAA WPC 7-day accumulated precipitation outlook."
    if forecast_meta:
        caption = (
            f"{caption} Valid window: {forecast_meta.get('valid_time')}. "
            f"Issued: {forecast_meta.get('issue_time')}."
        )
    if max_qpf_mm is not None:
        caption = f"{caption} Highest contour shown is about {max_qpf_mm:,.0f} mm."
    st.caption(caption)


def _render_drought_monitor_map():
    st.subheader("U.S. Drought Monitor")
    try:
        image_bytes, image_meta = _load_drought_monitor_image()
    except Exception as exc:
        st.error(f"Could not load the U.S. Drought Monitor map: {exc}")
        return

    st.image(image_bytes, use_container_width=True)
    caption_parts = []
    if image_meta.get("released"):
        caption_parts.append(f"Released: {image_meta['released']}.")
    if image_meta.get("valid"):
        caption_parts.append(f"Valid: {image_meta['valid']}.")
    st.caption(" ".join(caption_parts))


def _render_precipitation_map():
    st.subheader("National Precipitation Map")
    window_col, end_date_col = st.columns(2)
    with window_col:
        map_window_label = st.selectbox(
            "Accumulation window",
            options=["24h", "2 days", "7 days", "14 days"],
            index=2,
        )
    window_days = {"24h": 1, "2 days": 2, "7 days": 7, "14 days": 14}[map_window_label]
    prism_signature = _latest_prism_signature("ppt")
    preview_signature = _latest_prism_preview_signature()
    window_end_dates = _load_precip_window_dates(window_days, preview_signature, prism_signature)

    if not window_end_dates:
        st.info(f"No complete {window_days}-day national PRISM precipitation window is currently available on disk.")
        return

    latest_window_end_date = window_end_dates[-1]
    st.caption(f"Latest complete local PRISM window ends on {latest_window_end_date}.")
    with end_date_col:
        selected_end_date = st.selectbox(
            "Map end date",
            options=list(reversed(window_end_dates)),
            index=0,
            key=f"cot_precip_map_end_{window_days}_{latest_window_end_date}",
        )
    try:
        precip_map_df, precip_map_meta = _load_precip_map_preview(
            window_days,
            selected_end_date,
            preview_signature,
            prism_signature,
        )
    except ModuleNotFoundError as exc:
        st.error(str(exc))
        return

    if precip_map_df.empty:
        st.warning("The selected precipitation window could not be rendered from the current PRISM cache.")
        return

    color_max = float(precip_map_df["precip_mm"].quantile(0.98))
    if color_max <= 0:
        color_max = float(precip_map_df["precip_mm"].max())
    precip_map = px.scatter_geo(
        precip_map_df,
        lat="latitude",
        lon="longitude",
        color="precip_mm",
        color_continuous_scale="YlGnBu",
        range_color=(0, color_max if color_max > 0 else None),
        scope="usa",
        opacity=0.42,
        hover_data={"precip_mm": ":.1f", "latitude": False, "longitude": False},
        labels={"precip_mm": "Precipitation (mm)"},
        title=f"PRISM precipitation accumulation: {precip_map_meta['start_date']} to {precip_map_meta['end_date']}",
    )
    precip_map.update_traces(marker={"size": 4})

    overlay_choice = st.radio(
        "Footprint overlay",
        options=["Cotton footprint", "Rainfed / irrigated footprint", "None"],
        horizontal=True,
    )
    if overlay_choice == "Cotton footprint":
        cdl_preview = _load_cached_footprint_preview("none", _preview_signature("none"))
        if not cdl_preview.empty:
            _add_plain_footprint_overlay(precip_map, cdl_preview)
        else:
            st.warning(
                "Cotton footprint overlay data was not found. "
                "Expected a `data/processed/cdl_preview_cache/*12000_50_none*.parquet` file."
            )
    elif overlay_choice == "Rainfed / irrigated footprint":
        county_preview = _load_cached_footprint_preview("county_share", _preview_signature("county_share"))
        if not county_preview.empty:
            _add_county_share_footprint_overlay(precip_map, county_preview)
        else:
            st.warning(
                "Rainfed / irrigated footprint overlay data was not found. "
                "Expected a `data/processed/cdl_preview_cache/*12000_50_county_share*.parquet` file."
            )

    precip_map.update_layout(height=540)
    st.plotly_chart(precip_map, use_container_width=True, config={"scrollZoom": False})
    st.caption(
        f"National PRISM {window_days}-day accumulated precipitation ending on {precip_map_meta['end_date']}."
    )
    if overlay_choice == "Cotton footprint":
        st.caption("The black overlay marks the sampled 2024 USDA CDL cotton footprint preview.")
    elif overlay_choice == "Rainfed / irrigated footprint":
        st.caption("The colored overlay maps county irrigated cotton share onto the sampled cotton footprint: yellow is more rain-fed, red is more irrigated.")


def render_weather():
    st.header("Weather")
    forecast_col, drought_col = st.columns(2)
    with forecast_col:
        _render_forecast_map()
    with drought_col:
        _render_drought_monitor_map()
    _render_precipitation_map()
    state_precip_signature = _state_precip_signature()
    render_cotton_state_precipitation(
        state_precip=_load_state_precip_data(state_precip_signature),
        state_precip_meta=_load_state_precip_meta(state_precip_signature),
        state_precip_progress=_load_state_precip_progress(),
        key_prefix="weather_cotton_state_precip",
    )
