from __future__ import annotations

from datetime import date

import pandas as pd
import streamlit as st

from cotton_weather.glam_ndvi import (
    DEFAULT_GLAM_VERSION,
    SATELLITE_OPTIONS,
    SEASON_WINDOWS,
    GlamMask,
    GlamNdviError,
    AdmRegion,
    build_ndvi_seasonal_chart,
    fetch_glam_ndvi,
    get_adm_regions,
    get_cotton_masks,
    parse_glam_csv,
    regions_for_geography,
)


@st.cache_data(ttl=24 * 60 * 60)
def _load_cotton_masks() -> list[GlamMask]:
    return get_cotton_masks()


@st.cache_data(ttl=24 * 60 * 60)
def _load_adm_regions() -> list[AdmRegion]:
    return get_adm_regions()


@st.cache_data(ttl=6 * 60 * 60)
def _fetch_ndvi_csv(
    mask_id: str,
    region_id: str,
    years: tuple[int, ...],
    satellite: str,
    start_month: int,
    num_months: int,
    version: str,
) -> str:
    return fetch_glam_ndvi(
        mask_id=mask_id,
        region_id=region_id,
        years=list(years),
        satellite=satellite,
        start_month=start_month,
        num_months=num_months,
        version=version,
    )


def _mask_label(mask: GlamMask) -> str:
    return f"{mask.vintage} - {mask.display_name}"


def _year_options(start_year: int = 2013) -> list[int]:
    current_year = date.today().year
    return list(range(current_year, start_year - 1, -1))


def _default_years(options: list[int]) -> list[int]:
    return options[:6]


def _safe_index(options: list[str], value: str) -> int:
    return options.index(value) if value in options else 0


def render_ndvi_index():
    st.header("NDVI index")

    try:
        masks = _load_cotton_masks()
        regions = _load_adm_regions()
    except GlamNdviError as exc:
        st.error(f"Could not load GLAM metadata: {exc}")
        return

    if not masks:
        st.warning("No GLAM cotton-specific masks are currently available to this app.")
        return

    families = sorted({mask.family for mask in masks})
    control_row_1 = st.columns(3)
    with control_row_1[0]:
        family = st.selectbox("Crop-mask source", families, index=_safe_index(families, "GDA CropID"))

    family_masks = [mask for mask in masks if mask.family == family]
    geographies = sorted({mask.geography for mask in family_masks})
    default_geography = "United States" if "United States" in geographies else geographies[0]
    with control_row_1[1]:
        geography = st.selectbox("Country or geography", geographies, index=_safe_index(geographies, default_geography))

    geography_masks = [mask for mask in family_masks if mask.geography == geography]
    mask_options = {_mask_label(mask): mask for mask in sorted(geography_masks, key=lambda item: item.vintage, reverse=True)}
    with control_row_1[2]:
        mask_label = st.selectbox("Crop-mask vintage", list(mask_options.keys()))
    selected_mask = mask_options[mask_label]

    region_options = regions_for_geography(regions, geography)
    if not region_options:
        st.warning(f"No GLAM ADM regions are available for {geography}.")
        return
    region_label_map = {
        ("National aggregate" if region.level == "L0" else region.region_name): region
        for region in region_options
    }

    season = SEASON_WINDOWS.get(
        geography,
        {"start_month": selected_mask.start_month, "num_months": selected_mask.num_months, "note": "Calendar-year monitoring window."},
    )
    year_options = _year_options()
    control_row_2 = st.columns([1.1, 1.3, 1])
    with control_row_2[0]:
        region_label = st.selectbox("Administrative region", list(region_label_map.keys()), index=0)
    with control_row_2[1]:
        selected_years = st.multiselect(
            "NDVI observation years",
            options=year_options,
            default=_default_years(year_options),
            max_selections=8,
        )
    with control_row_2[2]:
        satellite_label = st.selectbox(
            "Satellite dataset",
            options=list(SATELLITE_OPTIONS.values()),
            index=0,
        )

    if not selected_years:
        st.info("Select at least one NDVI observation year.")
        return

    satellite_lookup = {label: code for code, label in SATELLITE_OPTIONS.items()}
    satellite = satellite_lookup[satellite_label]
    selected_region = region_label_map[region_label]
    start_month = int(season["start_month"])
    num_months = int(season["num_months"])

    try:
        csv_text = _fetch_ndvi_csv(
            selected_mask.mask_id,
            selected_region.region_id,
            tuple(sorted(selected_years)),
            satellite,
            start_month,
            num_months,
            DEFAULT_GLAM_VERSION,
        )
        ndvi_df, response_meta = parse_glam_csv(csv_text, region_name=region_label)
    except GlamNdviError as exc:
        st.error(f"GLAM NDVI request failed: {exc}")
        return
    except Exception as exc:
        st.error(f"Could not render the GLAM NDVI chart for this selection: {exc}")
        return

    selected_years_sorted = sorted(selected_years)
    chart = build_ndvi_seasonal_chart(ndvi_df, selected_years=selected_years_sorted, start_month=start_month)
    chart.update_layout(
        title=(
            f"{geography} cotton NDVI: {region_label} "
            f"({selected_mask.vintage} mask, {satellite})"
        )
    )
    st.plotly_chart(chart, use_container_width=True, config={"scrollZoom": False})

    latest_date = pd.to_datetime(ndvi_df["end_date"]).max()
    meta_bits = [
        f"GLAM database {response_meta.get('DB VERSION', DEFAULT_GLAM_VERSION)}",
        response_meta.get("PRODUCT", satellite),
        f"mask `{selected_mask.mask_id}`",
        f"latest observation ending {latest_date.date() if pd.notna(latest_date) else 'n/a'}",
    ]
    st.caption(" | ".join(meta_bits))
    st.caption(str(season.get("note", "")))
