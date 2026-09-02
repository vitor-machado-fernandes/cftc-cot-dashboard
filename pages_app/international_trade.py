import json
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from comexstat_cotton_exports_updater import (
    DATA_FILENAME,
    DEFAULT_NCM_CODES,
    METADATA_FILENAME,
    WEEKLY_SNAPSHOT_FILENAME,
    load_brazil_cotton_exports,
    load_brazil_cotton_weekly_snapshots,
    refresh_brazil_cotton_exports,
)


BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / DATA_FILENAME
METADATA_PATH = BASE_DIR / METADATA_FILENAME
WEEKLY_DATA_PATH = BASE_DIR / WEEKLY_SNAPSHOT_FILENAME


@st.cache_data
def _load_brazil_exports_cached(path_str: str, mtime_ns: int | None) -> pd.DataFrame:
    del path_str, mtime_ns
    return load_brazil_cotton_exports(BASE_DIR)


def _load_brazil_exports() -> pd.DataFrame:
    mtime_ns = DATA_PATH.stat().st_mtime_ns if DATA_PATH.exists() else None
    return _load_brazil_exports_cached(str(DATA_PATH), mtime_ns)


def _load_monthly_metadata() -> dict:
    if not METADATA_PATH.exists():
        return {}
    try:
        return json.loads(METADATA_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _format_date(value: object) -> str | None:
    timestamp = pd.to_datetime(value, errors="coerce")
    if pd.isna(timestamp):
        return None
    return timestamp.strftime("%Y-%m-%d")


@st.cache_data
def _load_weekly_snapshots_cached(path_str: str, mtime_ns: int | None) -> pd.DataFrame:
    del path_str, mtime_ns
    return load_brazil_cotton_weekly_snapshots(BASE_DIR)


def _load_weekly_snapshots() -> pd.DataFrame:
    mtime_ns = WEEKLY_DATA_PATH.stat().st_mtime_ns if WEEKLY_DATA_PATH.exists() else None
    return _load_weekly_snapshots_cached(str(WEEKLY_DATA_PATH), mtime_ns)


def _marketing_year_label(year: int) -> str:
    return f"{year}/{str(year + 1)[-2:]}"


def _monthly_chart(chart_df: pd.DataFrame) -> None:
    monthly = (
        chart_df.groupby("date", as_index=False)["weight_tons"]
        .sum()
        .sort_values("date")
    )
    fig = px.bar(
        monthly,
        x="date",
        y="weight_tons",
        title="Brazil Monthly Cotton Exports",
        labels={"date": "", "weight_tons": "Tons"},
    )
    fig.update_traces(marker_color="#d97706", hovertemplate="%{x|%b %Y}<br>%{y:,.0f} tons<extra></extra>")
    fig.update_layout(
        height=430,
        margin=dict(l=8, r=8, t=56, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        xaxis=dict(showgrid=False),
        yaxis=dict(tickformat=",", gridcolor="#d9d9d9"),
    )
    st.plotly_chart(fig, use_container_width=True)


def _marketing_year_chart(chart_df: pd.DataFrame) -> None:
    season_df = chart_df.copy()
    season_df["year"] = season_df["date"].dt.year
    season_df["month"] = season_df["date"].dt.month
    season_df["marketing_year"] = season_df["year"].where(
        season_df["month"] >= 8,
        season_df["year"] - 1,
    )
    season_df["marketing_month"] = (season_df["month"] - 8) % 12 + 1
    season_df["marketing_year_label"] = season_df["marketing_year"].apply(_marketing_year_label)

    monthly = (
        season_df.groupby(["marketing_year", "marketing_year_label", "marketing_month"], as_index=False)[
            "weight_tons"
        ]
        .sum()
        .sort_values(["marketing_year", "marketing_month"])
    )
    latest_marketing_year = int(monthly["marketing_year"].max())
    plot_df = monthly[monthly["marketing_year"] >= latest_marketing_year - 4].copy()
    plot_df["month_label"] = plot_df["marketing_month"].map(
        dict(enumerate(["Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"], start=1))
    )

    fig = px.bar(
        plot_df,
        x="month_label",
        y="weight_tons",
        color="marketing_year_label",
        barmode="group",
        title="Brazil Cotton Exports by Marketing Year",
        labels={"month_label": "", "weight_tons": "Tons", "marketing_year_label": "Marketing Year"},
        category_orders={"month_label": ["Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"]},
    )
    fig.update_layout(
        height=430,
        margin=dict(l=8, r=8, t=56, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        yaxis=dict(tickformat=",", gridcolor="#d9d9d9"),
    )
    st.plotly_chart(fig, use_container_width=True)


def _weekly_partial_chart() -> None:
    weekly = _load_weekly_snapshots()
    st.subheader("Brazilian Cotton Exports - Weekly Partial")

    if weekly.empty:
        st.info(
            "No weekly preliminary snapshots have been captured yet. The automatic updater will begin storing them when the app can reach MDIC's weekly publication."
        )
        return

    weekly = weekly.copy()
    weekly["period_start"] = pd.to_datetime(weekly["period_start"], errors="coerce")
    weekly["week_of_month"] = pd.to_numeric(weekly["week_of_month"], errors="coerce")
    weekly["value_usd_million_cumulative"] = pd.to_numeric(
        weekly["value_usd_million_cumulative"],
        errors="coerce",
    )
    weekly = weekly.dropna(
        subset=["period_start", "week_of_month", "value_usd_million_cumulative"]
    ).sort_values(["period_start", "week_of_month"])

    if weekly.empty:
        st.info("Weekly preliminary snapshots exist locally, but none are valid for charting.")
        return

    weekly["period_label"] = weekly["period_start"].dt.strftime("%b %Y")
    period_options = (
        weekly[["period_start", "period_label"]]
        .drop_duplicates()
        .sort_values("period_start", ascending=False)
    )
    selected_period = st.selectbox(
        "Month and year",
        options=period_options["period_label"].tolist(),
        key="international_trade_brazil_weekly_period",
    )

    period_df = weekly[weekly["period_label"] == selected_period].copy()
    period_df = period_df.sort_values("week_of_month")
    period_df["weekly_value_usd_million"] = (
        period_df["value_usd_million_cumulative"].diff().fillna(
            period_df["value_usd_million_cumulative"]
        )
    )
    period_df["week_label"] = period_df["week_of_month"].astype(int).astype(str)

    fig = px.bar(
        period_df,
        x="week_label",
        y="weekly_value_usd_million",
        title=f"Brazil Raw Cotton Weekly Partial Exports - {selected_period}",
        labels={
            "week_label": "Week of Month",
            "weekly_value_usd_million": "US$ Million",
        },
    )
    fig.update_traces(
        marker_color="#2563eb",
        hovertemplate="Week %{x}<br>%{y:,.2f} US$ million<extra></extra>",
    )
    fig.update_layout(
        height=380,
        margin=dict(l=8, r=8, t=56, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        xaxis=dict(type="category", showgrid=False),
        yaxis=dict(tickformat=",.2f", gridcolor="#d9d9d9"),
    )
    st.plotly_chart(fig, use_container_width=True)

    latest_snapshot = pd.to_datetime(period_df["snapshot_date"], errors="coerce").max()
    latest_cumulative = period_df["value_usd_million_cumulative"].iloc[-1]
    snapshot_label = latest_snapshot.date() if pd.notna(latest_snapshot) else "N/A"
    st.caption(
        "Source: MDIC preliminary weekly trade publication. "
        f"Latest snapshot: {snapshot_label}. "
        f"Month-to-date raw cotton exports: {latest_cumulative:,.2f} US$ million. "
        "Bars are calculated from captured cumulative snapshots, so past weeks are available only after this tracker has observed them."
    )


def render_international_trade():
    st.header("International Trade")

    df = _load_brazil_exports()
    if df.empty:
        st.warning("No local ComexStat Brazilian cotton export data is available yet.")
        return

    df = df.copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["weight_tons"] = pd.to_numeric(df["weight_tons"], errors="coerce")
    df = df.dropna(subset=["date", "weight_tons"])

    st.subheader("Brazilian Cotton Exports")

    ncm_options = {
        code: f"{code} - {DEFAULT_NCM_CODES.get(code, code)}"
        for code in sorted(df["ncm_code"].dropna().astype(str).unique())
    }
    default_ncm = ["52010020"] if "52010020" in ncm_options else list(ncm_options)
    selected_ncm = st.multiselect(
        "NCM codes",
        options=list(ncm_options),
        default=default_ncm,
        format_func=lambda code: ncm_options.get(code, code),
        key="international_trade_brazil_exports_ncm",
    )

    min_date = df["date"].min().date()
    max_date = df["date"].max().date()
    selected_dates = st.date_input(
        "Date range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        key="international_trade_brazil_exports_date_range",
    )
    if not isinstance(selected_dates, tuple) or len(selected_dates) != 2:
        st.warning("Select a start and end date for the chart.")
        return
    start_date, end_date = selected_dates
    view = st.radio(
        "View",
        ["Monthly", "Marketing year"],
        horizontal=True,
        key="international_trade_brazil_exports_view",
    )

    chart_df = df[
        df["ncm_code"].astype(str).isin(selected_ncm)
        & (df["date"].dt.date >= start_date)
        & (df["date"].dt.date <= end_date)
    ].copy()

    if chart_df.empty:
        st.warning("No exports match the selected filters.")
        return

    if view == "Monthly":
        _monthly_chart(chart_df)
    else:
        _marketing_year_chart(chart_df)

    latest_month = chart_df["date"].max()
    total_tons = chart_df["weight_tons"].sum()
    metadata = _load_monthly_metadata()
    update_label = _format_date(metadata.get("comexstat_updated_date"))
    update_text = f" ComexStat latest publication: {update_label}." if update_label else ""
    st.caption(
        f"Source: ComexStat / MDIC. Latest month: {latest_month:%Y-%m}. "
        f"Selected-period total: {total_tons:,.0f} tons."
        f"{update_text}"
    )

    _weekly_partial_chart()
