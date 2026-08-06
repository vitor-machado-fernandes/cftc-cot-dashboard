from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st


REGION_STATES = {
    "Delta Region": ["AR", "TN", "MO", "MS", "AL", "LA"],
    "Southeast Region": ["GA", "FL", "SC", "NC", "VA"],
    "Southwest Region": ["TX", "OK", "KS", "NM"],
    "Far West Region": ["AZ", "CA"],
}


def _format_metric(value: float | int | None, suffix: str = "") -> str:
    if value is None or pd.isna(value):
        return "n/a"
    if isinstance(value, float):
        return f"{value:,.1f}{suffix}"
    return f"{value:,}{suffix}"


def build_seasonal_comparison(
    precip_df: pd.DataFrame,
    geography: str,
    comparison_years: int | None = 10,
) -> tuple[pd.DataFrame, dict]:
    if precip_df.empty:
        return pd.DataFrame(), {}

    working = precip_df.copy()
    working["date"] = pd.to_datetime(working["date"])
    working["year"] = working["date"].dt.year
    working["month_day"] = working["date"].dt.strftime("%m-%d")

    if geography == "National":
        rows: list[dict] = []
        for day, frame in working.groupby("date"):
            weights = frame["cotton_area_acres_est"].fillna(0.0)
            value = float(np.average(frame["ppt_mm"], weights=weights)) if weights.sum() > 0 else float(frame["ppt_mm"].mean())
            rows.append({"date": day, "ppt_mm": value})
        daily = pd.DataFrame(rows)
    elif geography in REGION_STATES:
        rows = []
        region_frame = working.loc[working["state"].isin(REGION_STATES[geography])].copy()
        for day, frame in region_frame.groupby("date"):
            weights = frame["cotton_area_acres_est"].fillna(0.0)
            value = float(np.average(frame["ppt_mm"], weights=weights)) if weights.sum() > 0 else float(frame["ppt_mm"].mean())
            rows.append({"date": day, "ppt_mm": value})
        daily = pd.DataFrame(rows)
    else:
        daily = working.loc[working["state"] == geography, ["date", "ppt_mm"]].copy()

    if daily.empty:
        return pd.DataFrame(), {}

    daily["year"] = daily["date"].dt.year
    current_year = int(daily["year"].max())
    current_rows = daily.loc[daily["year"] == current_year].copy()
    latest_date = current_rows["date"].max()
    cutoff_month_day = latest_date.strftime("%m-%d")

    comparison_pool = sorted([year for year in daily["year"].unique() if year < current_year], reverse=True)
    selected_history = comparison_pool if comparison_years is None else comparison_pool[:comparison_years]
    selected_years = sorted(selected_history + [current_year])

    filtered = daily.loc[daily["year"].isin(selected_years)].copy()
    filtered["month_day"] = filtered["date"].dt.strftime("%m-%d")
    filtered = filtered.loc[filtered["month_day"] <= cutoff_month_day].copy()
    if filtered.empty:
        return pd.DataFrame(), {}

    filtered["cumulative_ppt_mm"] = filtered.groupby("year")["ppt_mm"].cumsum()
    filtered["plot_date"] = pd.to_datetime("2000-" + filtered["month_day"], format="%Y-%m-%d")

    average_line = (
        filtered.loc[filtered["year"].isin(selected_history)]
        .groupby(["month_day", "plot_date"], as_index=False)["cumulative_ppt_mm"]
        .mean()
        .sort_values("plot_date")
    )
    if not average_line.empty:
        average_line["year_label"] = f"Average ({len(selected_history)} yrs)"

    display = filtered[["plot_date", "year", "cumulative_ppt_mm"]].copy()
    display["year_label"] = display["year"].astype(str)
    if not average_line.empty:
        display = pd.concat(
            [display, average_line[["plot_date", "cumulative_ppt_mm", "year_label"]]],
            ignore_index=True,
        )

    current_total = float(filtered.loc[filtered["year"] == current_year, "cumulative_ppt_mm"].iloc[-1])
    current_daily = filtered.loc[filtered["year"] == current_year].sort_values("plot_date")
    daily_change_24h = None
    if len(current_daily) >= 2:
        daily_change_24h = float(current_daily["cumulative_ppt_mm"].iloc[-1] - current_daily["cumulative_ppt_mm"].iloc[-2])
    average_total = float(average_line["cumulative_ppt_mm"].iloc[-1]) if not average_line.empty else None

    return display, {
        "current_year": current_year,
        "latest_date": latest_date.strftime("%Y-%m-%d"),
        "history_years": selected_history,
        "current_total_mm": current_total,
        "daily_change_24h_mm": daily_change_24h,
        "average_total_mm": average_total,
    }


def render_cotton_state_precipitation(
    state_precip: pd.DataFrame,
    state_precip_meta: dict,
    state_precip_progress: dict | None = None,
    *,
    key_prefix: str = "cotton_state_precip",
) -> None:
    st.subheader("Cotton-State Precipitation")

    if state_precip_progress:
        phase = state_precip_progress.get("phase", "unknown")
        current_day = int(state_precip_progress.get("current_day", 0))
        total_days = int(state_precip_progress.get("total_days", 0))
        current_date = state_precip_progress.get("date")
        progress_ratio = (current_day / total_days) if total_days else 0.0
        st.progress(progress_ratio, text=f"Precipitation build status: {phase} ({current_day}/{total_days})")
        st.caption(
            f"Current processing date: {current_date}. "
            f"States: {', '.join(state_precip_progress.get('states', []))}."
        )

    if state_precip.empty:
        st.info("No state-level cotton precipitation dataset found yet. It will be created by the daily weather refresh job.")
        return

    available_states = sorted(state_precip["state"].dropna().unique().tolist())
    state_precip_dates = pd.to_datetime(state_precip["date"])
    current_year_available = int(state_precip_dates.dt.year.max())
    available_history_years = sorted(
        [int(year) for year in state_precip_dates.dt.year.unique() if int(year) < current_year_available]
    )

    geography = st.selectbox(
        "Seasonal comparison geography",
        options=["National"] + list(REGION_STATES.keys()) + available_states,
        index=0,
        key=f"{key_prefix}_geography",
    )
    st.caption(
        "Regional definitions: Delta = AR, TN, MO, MS, AL, LA; "
        "Southeast = GA, FL, SC, NC, VA; "
        "Southwest = TX, OK, KS, NM; Far West = AZ, CA."
    )

    comparison_options: list[tuple[str, int | None]] = []
    for years in [10, 15, 20]:
        if len(available_history_years) >= years:
            comparison_options.append((f"Previous {years} years", years))
    if available_history_years:
        comparison_options.append((f"All available ({len(available_history_years)} years)", None))

    comparison_label_to_value = dict(comparison_options)
    comparison_label = st.selectbox(
        "Historical comparison window",
        options=list(comparison_label_to_value.keys()),
        index=0 if comparison_options else None,
        key=f"{key_prefix}_comparison_window",
    ) if comparison_options else None

    seasonal_df, seasonal_meta = build_seasonal_comparison(
        precip_df=state_precip,
        geography=geography,
        comparison_years=comparison_label_to_value.get(comparison_label, 10),
    )
    if seasonal_df.empty:
        st.warning("No seasonal precipitation comparison could be built for the selected geography yet.")
        return

    average_total = seasonal_meta.get("average_total_mm")
    current_total = seasonal_meta.get("current_total_mm")
    anomaly = (current_total - average_total) if average_total is not None else None
    daily_change_24h = seasonal_meta.get("daily_change_24h_mm")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("24h change", _format_metric(daily_change_24h, " mm"))
    col2.metric("Current-year cutoff", seasonal_meta["latest_date"])
    col3.metric(f"{seasonal_meta['current_year']} cumulative ppt", _format_metric(current_total, " mm"))
    col4.metric("Vs. average", _format_metric(anomaly, " mm"))

    average_label = f"Average ({len(seasonal_meta['history_years'])} yrs)"
    line_order = [str(seasonal_meta["current_year"])] + [str(year) for year in reversed(seasonal_meta["history_years"])]
    if average_label in seasonal_df["year_label"].values:
        line_order.append(average_label)

    seasonal_chart = px.line(
        seasonal_df,
        x="plot_date",
        y="cumulative_ppt_mm",
        color="year_label",
        category_orders={"year_label": line_order},
        color_discrete_map={str(seasonal_meta["current_year"]): "#0f172a", average_label: "#d97706"},
        labels={
            "plot_date": "Season",
            "cumulative_ppt_mm": "Cumulative precipitation (mm)",
            "year_label": "Series",
        },
        title=f"Cumulative cotton-area precipitation: {geography}",
    )
    seasonal_chart.update_xaxes(dtick="M1", tickformat="%b")
    seasonal_chart.update_layout(height=460, legend_title_text="Series")
    for trace in seasonal_chart.data:
        if trace.name == average_label:
            trace.update(line={"width": 4, "dash": "dash"})
        elif trace.name == str(seasonal_meta["current_year"]):
            trace.update(line={"width": 4})
        else:
            trace.update(line={"width": 1.5}, opacity=0.45)
    st.plotly_chart(seasonal_chart, use_container_width=True)
    st.caption(
        f"The chart shows {seasonal_meta['current_year']} against the previous "
        f"{len(seasonal_meta['history_years'])} available years, all cut off at {seasonal_meta['latest_date']}."
    )
    st.caption(
        f"Processed precipitation coverage currently loaded in the app runs from "
        f"{state_precip_dates.min().date()} through {state_precip_dates.max().date()}."
    )

    if state_precip_meta:
        st.caption(
            f"State precipitation dataset updated at {state_precip_meta.get('updated_at')}. "
            f"Coverage is {state_precip_meta.get('dataset_start')} through {state_precip_meta.get('dataset_end')}."
        )
