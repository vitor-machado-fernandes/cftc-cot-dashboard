from pathlib import Path

import pandas as pd
import streamlit as st

from usda_cotton_end_stocks_updater import (
    DATA_FILE,
    load_cotton_end_stocks,
    refresh_usda_cotton_end_stocks,
)


@st.cache_data
def _load_stocks_df_cached(
    path_str: str,
    modified_time_ns: int,
    file_size: int,
) -> pd.DataFrame:
    del modified_time_ns, file_size
    return load_cotton_end_stocks(path_str)


def load_stocks_df(path: str | Path = DATA_FILE) -> pd.DataFrame:
    workbook_path = Path(path)
    if not workbook_path.exists():
        return pd.DataFrame()

    stat = workbook_path.stat()
    return _load_stocks_df_cached(
        str(workbook_path),
        stat.st_mtime_ns,
        stat.st_size,
    )


def render_stocks_use():
    st.header("Stocks & Use")

    df = load_stocks_df()
    if df.empty:
        st.warning(
            "`usda_cotton_end_stocks.xlsx` was not found or does not contain data."
        )
        if st.button("Update data", type="primary"):
            with st.spinner("Checking USDA WASDE releases..."):
                result = refresh_usda_cotton_end_stocks()
            _load_stocks_df_cached.clear()
            if result["did_update"]:
                st.success(
                    f"Added {result['rows_added']} ending-stocks row(s) and "
                    f"{result.get('use_rows_added', 0)} use row(s)."
                )
                st.rerun()
            else:
                st.info("No new rows were added.")
        return

    latest = df.sort_values("ReleaseDate").iloc[-1]

    metric_1, metric_2, metric_3 = st.columns(3)
    metric_1.metric("Latest WASDE Release", latest["ReleaseDate"].strftime("%Y-%m-%d"))
    metric_2.metric("Market Year", str(latest["MarketYear"]))
    metric_3.metric("Ending Stocks", f"{latest['Value']:.2f} mil. bales")

    if st.button("Update data", type="primary"):
        with st.spinner("Checking USDA WASDE releases..."):
            try:
                result = refresh_usda_cotton_end_stocks()
            except Exception as exc:
                st.error(f"USDA WASDE update failed: {exc}")
                return

        _load_stocks_df_cached.clear()

        if result["did_update"]:
            st.success(
                f"Added {result['rows_added']} ending-stocks row(s) and "
                f"{result.get('use_rows_added', 0)} use row(s)."
            )
            if result.get("prices_updated"):
                st.success(f"Added {result['prices_updated']} cotton futures price(s).")
            if result["errors"]:
                st.warning("Some releases could not be parsed: " + "; ".join(result["errors"]))
            st.rerun()

        if result.get("prices_updated"):
            st.success(f"Added {result['prices_updated']} cotton futures price(s).")
            st.rerun()

        latest_remote = result.get("latest_remote")
        if latest_remote is not None and pd.notna(latest_remote):
            st.info(f"Already up to date through {latest_remote:%Y-%m-%d}.")
        else:
            st.info("No newer WASDE release was found.")

        if result["errors"]:
            st.warning("Some releases could not be parsed: " + "; ".join(result["errors"]))

    display_columns = ["WasdeNumber", "ReleaseDate", "MarketYear", "Value"]
    if "Stocks/Use" in df.columns:
        display_columns.append("Stocks/Use")
    display_columns.extend(["Contract", "Price", "Unit"])

    display_df = (
        df.sort_values("ReleaseDate", ascending=False)
        .head(12)
        [display_columns]
        .rename(
            columns={
                "WasdeNumber": "WASDE",
                "ReleaseDate": "Release Date",
                "MarketYear": "Market Year",
                "Value": "Ending Stocks",
            }
        )
    )
    st.dataframe(display_df, use_container_width=True, hide_index=True)
