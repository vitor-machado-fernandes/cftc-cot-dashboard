from pathlib import Path

import numpy as np
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


def _regression_line(
    chart_df: pd.DataFrame,
    model: str,
) -> tuple[pd.DataFrame, float]:
    model_df = chart_df.dropna(subset=["Stocks/Use", "Price"]).copy()
    is_log_model = model in ["Log price", "Piecewise log price"]
    is_piecewise = model in ["Piecewise OLS", "Piecewise log price"]
    breakpoint = 0.20

    if is_log_model:
        model_df = model_df[model_df["Price"] > 0].copy()

    if len(model_df) < 2:
        return pd.DataFrame(columns=["Stocks/Use", "Fitted Price"]), float("nan")

    x = model_df["Stocks/Use"].astype(float).to_numpy()
    y = model_df["Price"].astype(float).to_numpy()

    if is_log_model:
        y_fit_source = np.log(y)
    else:
        y_fit_source = y

    if is_piecewise:
        left_mask = x <= breakpoint
        right_mask = x > breakpoint
        if left_mask.sum() < 2 or right_mask.sum() < 2:
            return pd.DataFrame(columns=["Stocks/Use", "Fitted Price"]), float("nan")

        left_slope, left_intercept = np.polyfit(x[left_mask], y_fit_source[left_mask], 1)
        right_slope, right_intercept = np.polyfit(x[right_mask], y_fit_source[right_mask], 1)
        y_pred_source = np.where(
            left_mask,
            left_intercept + left_slope * x,
            right_intercept + right_slope * x,
        )
        ss_res = float(np.sum((y_fit_source - y_pred_source) ** 2))
        ss_tot = float(np.sum((y_fit_source - np.mean(y_fit_source)) ** 2))

        line_x_left = np.linspace(x.min(), min(breakpoint, x.max()), 50)
        line_x_right = np.linspace(max(breakpoint, x.min()), x.max(), 50)
        line_y_left = left_intercept + left_slope * line_x_left
        line_y_right = right_intercept + right_slope * line_x_right
        line_x = np.concatenate([line_x_left, line_x_right])
        line_y = np.concatenate([line_y_left, line_y_right])
        if is_log_model:
            line_y = np.exp(line_y)
    elif is_log_model:
        slope, intercept = np.polyfit(x, y_fit_source, 1)
        y_pred_source = intercept + slope * x
        ss_res = float(np.sum((y_fit_source - y_pred_source) ** 2))
        ss_tot = float(np.sum((y_fit_source - np.mean(y_fit_source)) ** 2))

        line_x = np.linspace(x.min(), x.max(), 100)
        line_y = np.exp(intercept + slope * line_x)
    else:
        slope, intercept = np.polyfit(x, y, 1)
        y_pred = intercept + slope * x
        ss_res = float(np.sum((y - y_pred) ** 2))
        ss_tot = float(np.sum((y - np.mean(y)) ** 2))

        line_x = np.linspace(x.min(), x.max(), 100)
        line_y = intercept + slope * line_x

    r_squared = 1 - ss_res / ss_tot if ss_tot else float("nan")
    return pd.DataFrame({"Stocks/Use": line_x, "Fitted Price": line_y}), r_squared


def _year_scale_scatter_with_regression(chart_df: pd.DataFrame, model: str) -> None:
    line_df, r_squared = _regression_line(chart_df, model)
    points = chart_df.copy()
    points["ReleaseDate"] = pd.to_datetime(points["ReleaseDate"], errors="coerce").dt.strftime(
        "%Y-%m-%d"
    )

    spec = {
        "height": 430,
        "data": {"values": points.to_dict("records")},
        "layer": [
            {
                "mark": {"type": "circle", "size": 70, "opacity": 0.78},
                "encoding": {
                    "x": {
                        "field": "Stocks/Use",
                        "type": "quantitative",
                        "title": "Stocks/Use",
                    },
                    "y": {
                        "field": "Price",
                        "type": "quantitative",
                        "title": "July Cotton Futures Price",
                    },
                    "color": {
                        "field": "ForecastYear",
                        "type": "quantitative",
                        "title": "Forecast Year",
                        "scale": {"scheme": "viridis"},
                    },
                    "tooltip": [
                        {"field": "ReleaseDate", "type": "nominal", "title": "Release"},
                        {"field": "MarketYear", "type": "nominal", "title": "Market Year"},
                        {"field": "Contract", "type": "nominal", "title": "Contract"},
                        {
                            "field": "Stocks/Use",
                            "type": "quantitative",
                            "format": ".3f",
                            "title": "Stocks/Use",
                        },
                        {
                            "field": "Price",
                            "type": "quantitative",
                            "format": ".2f",
                            "title": "Price",
                        },
                    ],
                },
            },
            {
                "data": {"values": line_df.to_dict("records")},
                "mark": {"type": "line", "color": "#111827", "strokeWidth": 3},
                "encoding": {
                    "x": {"field": "Stocks/Use", "type": "quantitative"},
                    "y": {"field": "Fitted Price", "type": "quantitative"},
                },
            },
        ],
    }

    st.vega_lite_chart(spec, use_container_width=True)
    if not np.isnan(r_squared):
        r2_label = "log-space R²" if model == "Log price" else "R²"
        st.caption(f"{model} regression {r2_label}: {r_squared:.3f}")


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

    if {"Stocks/Use", "Price"}.issubset(df.columns):
        chart_df = df.copy()
        chart_df["Stocks/Use"] = pd.to_numeric(chart_df["Stocks/Use"], errors="coerce")
        chart_df["Price"] = pd.to_numeric(chart_df["Price"], errors="coerce")
        chart_df = chart_df.dropna(subset=["Stocks/Use", "Price"]).sort_values("ReleaseDate")

        if not chart_df.empty:
            st.subheader("Stocks/Use vs July Cotton Futures")
            st.scatter_chart(
                chart_df,
                x="Stocks/Use",
                y="Price",
                color="MarketYear",
                use_container_width=True,
            )

            if "ForecastYear" in chart_df.columns:
                st.subheader("Stocks/Use vs July Cotton Futures - Year Scale")
                chart_df["ForecastYear"] = pd.to_numeric(
                    chart_df["ForecastYear"],
                    errors="coerce",
                )
                chart_df = chart_df.dropna(subset=["ForecastYear"])
                regression_model = st.radio(
                    "Regression model",
                    ["OLS", "Log price", "Piecewise OLS", "Piecewise log price"],
                    horizontal=True,
                    key="stocks_use_regression_model",
                )
                _year_scale_scatter_with_regression(chart_df, regression_model)

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
