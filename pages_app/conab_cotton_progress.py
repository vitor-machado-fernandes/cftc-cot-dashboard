"""Streamlit tab for Brazilian cotton crop progress from CONAB."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
import streamlit as st

from conab.chart import available_seasons, build_progress_chart, cotton_rows
from conab.parser import parse_workbook
from conab.scraper import crawl_bulletins, download_bulletins
from conab.storage import latest_bulletin_date, load_data, upsert_data

LOGGER = logging.getLogger(__name__)
BASE_DIR = Path(__file__).resolve().parents[1]
RAW_DIR = BASE_DIR / "data" / "raw" / "conab"


@st.cache_data(ttl=60 * 30, show_spinner=False)
def load_conab_progress_data() -> pd.DataFrame:
    return load_data()


def run_conab_update() -> tuple[int, int, int]:
    latest = latest_bulletin_date()
    stop_after = latest.date() if latest is not None else None
    max_pages = 2 if latest is None else 1
    bulletins = crawl_bulletins(stop_after_date=stop_after, max_pages=max_pages)
    downloaded = download_bulletins(bulletins, cache_dir=RAW_DIR)

    parsed_frames: list[pd.DataFrame] = []
    for bulletin, path in downloaded:
        parsed = parse_workbook(path, bulletin_date=bulletin.bulletin_date)
        if parsed.empty:
            LOGGER.warning("No parseable cotton progress rows in %s", path)
            continue
        parsed_frames.append(parsed)

    if not parsed_frames:
        return 0, len(bulletins), len(downloaded)

    new_rows = upsert_data(pd.concat(parsed_frames, ignore_index=True))
    return new_rows, len(bulletins), len(downloaded)


def render_conab_cotton_progress() -> None:
    st.subheader("Cotton Crop Progress (CONAB)")

    update_col, status_col = st.columns([1, 3])
    with update_col:
        update_clicked = st.button("Update CONAB data", type="primary", key="conab_update_button")

    if update_clicked:
        try:
            with st.spinner("Checking CONAB bulletins and parsing new workbooks..."):
                new_rows, bulletin_count, download_count = run_conab_update()
            load_conab_progress_data.clear()
            if new_rows:
                status_col.success(
                    f"CONAB update complete. Added {new_rows:,} new records from {download_count:,} workbook(s)."
                )
            elif download_count:
                status_col.warning(
                    f"CONAB workbooks were checked ({download_count:,}), but no new cotton rows were added."
                )
            else:
                status_col.warning(
                    f"No new CONAB workbooks were found on the checked index page(s) ({bulletin_count:,} links)."
                )
        except Exception as exc:
            LOGGER.exception("CONAB update failed")
            status_col.error(f"CONAB update failed: {exc}")

    df = load_conab_progress_data()
    cotton_df = cotton_rows(df) if not df.empty else df
    if cotton_df.empty:
        st.info("No CONAB cotton progress data is stored yet. Click Update CONAB data to fetch public CONAB workbooks.")
        return

    states = sorted(cotton_df["state"].dropna().unique().tolist())
    default_state_idx = states.index("MT") if "MT" in states else 0
    state = st.selectbox("State (UF)", states, index=default_state_idx, key="conab_state")

    seasons = available_seasons(cotton_df, state)
    if not seasons:
        st.info(f"No cotton progress rows are available for {state}.")
        return

    season = st.selectbox("Season", seasons, index=len(seasons) - 1, key="conab_season")
    st.plotly_chart(build_progress_chart(cotton_df, state, season), use_container_width=True)

    latest_date = pd.to_datetime(cotton_df["bulletin_date"]).max()
    st.caption(f"Source: CONAB. Stored rows: {len(cotton_df):,}. Latest bulletin date: {latest_date.date()}.")
