import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from pages_app.CoT_position import COMMODITY_SHEETS, load_commodity_df, resolve_position_columns


LONG_4_COL_CANDIDATES = [
    "conc_net_le_4_tdr_long_all",
    "conc_net_le_4_ldr_long_all",
]
SHORT_4_COL_CANDIDATES = [
    "conc_net_le_4_tdr_short_all",
    "conc_net_le_4_ldr_short_all",
]
LONG_8_COL_CANDIDATES = [
    "conc_net_le_8_tdr_long_all",
    "conc_net_le_8_ldr_long_all",
]
SHORT_8_COL_CANDIDATES = [
    "conc_net_le_8_tdr_short_all",
    "conc_net_le_8_ldr_short_all",
]

TRADER_ORDER = [
    "PMPU",
    "Swap Dealer",
    "Managed Money",
    "Other",
    "Non-Reportables",
]

TRADER_COLORS = {
    "PMPU": "#7CC7FF",
    "Swap Dealer": "#6ED3B6",
    "Managed Money": "#FDBA74",
    "Other": "#C4B5FD",
    "Non-Reportables": "#FDA4AF",
}


def _clean_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"": None, "nan": None, "None": None}),
        errors="coerce",
    )


def _resolve_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = set(df.columns)
    return next((col for col in candidates if col in cols), None)


def _prepare_concentration_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if "report_date" not in out.columns and "report_date_as_yyyy_mm_dd" in out.columns:
        out["report_date"] = pd.to_datetime(out["report_date_as_yyyy_mm_dd"], errors="coerce")
    else:
        out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")

    out = out.dropna(subset=["report_date"]).copy()
    out["year"] = out["report_date"].dt.year

    long_4_col = _resolve_column(out, LONG_4_COL_CANDIDATES)
    short_4_col = _resolve_column(out, SHORT_4_COL_CANDIDATES)
    long_8_col = _resolve_column(out, LONG_8_COL_CANDIDATES)
    short_8_col = _resolve_column(out, SHORT_8_COL_CANDIDATES)

    out["long_concentration_4"] = _clean_numeric(out[long_4_col]) if long_4_col else pd.NA
    out["short_concentration_4"] = _clean_numeric(out[short_4_col]) if short_4_col else pd.NA
    out["long_concentration_8"] = _clean_numeric(out[long_8_col]) if long_8_col else pd.NA
    out["short_concentration_8"] = _clean_numeric(out[short_8_col]) if short_8_col else pd.NA

    return out[
        [
            "report_date",
            "year",
            "long_concentration_4",
            "long_concentration_8",
            "short_concentration_4",
            "short_concentration_8",
        ]
    ].sort_values("report_date")


def _prepare_net_pct_oi_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if "report_date" not in out.columns and "report_date_as_yyyy_mm_dd" in out.columns:
        out["report_date"] = pd.to_datetime(out["report_date_as_yyyy_mm_dd"], errors="coerce")
    else:
        out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")

    out = out.dropna(subset=["report_date"]).copy()
    out["year"] = out["report_date"].dt.year

    oi_col = _resolve_column(out, ["open_interest_all", "open_interest"])
    out["open_interest_all"] = _clean_numeric(out[oi_col]) if oi_col else pd.NA

    result_cols = ["report_date", "year", "open_interest_all"]

    for trader in TRADER_ORDER:
        long_col, short_col, _ = resolve_position_columns(out, trader, "All")
        trader_key = trader.lower().replace(" ", "_").replace("-", "_")

        long_vals = (
            _clean_numeric(out[long_col])
            if long_col
            else pd.Series(float("nan"), index=out.index, dtype="float64")
        )
        short_vals = (
            _clean_numeric(out[short_col])
            if short_col
            else pd.Series(float("nan"), index=out.index, dtype="float64")
        )
        net_vals = long_vals - short_vals

        out[f"{trader_key}_net"] = net_vals
        out[f"{trader_key}_pct_oi"] = (
            net_vals.div(out["open_interest_all"]).where(out["open_interest_all"] != 0) * 100
        )
        result_cols.extend([f"{trader_key}_net", f"{trader_key}_pct_oi"])

    return out[result_cols].sort_values("report_date")


def _get_percent_axis_config(df_plot: pd.DataFrame) -> tuple[str, str]:
    vals = pd.concat(
        [
            pd.to_numeric(df_plot["long_concentration_4"], errors="coerce"),
            pd.to_numeric(df_plot["long_concentration_8"], errors="coerce"),
            pd.to_numeric(df_plot["short_concentration_4"], errors="coerce"),
            pd.to_numeric(df_plot["short_concentration_8"], errors="coerce"),
        ],
        ignore_index=True,
    ).dropna()

    if vals.empty:
        return "Concentration", None

    if vals.max() <= 1.0:
        return "Concentration", ".0%"

    return "Concentration (%)", None


def _build_side_figure(
    df_plot: pd.DataFrame,
    title: str,
    col_4: str,
    col_8: str,
    line_color_4: str,
    fill_color_4: str,
    line_color_8: str,
    fill_color_8: str,
) -> go.Figure:
    yaxis_title, tickformat = _get_percent_axis_config(df_plot)

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot[col_4],
            mode="lines",
            name="Top 4",
            line=dict(color=line_color_4, width=2),
            fill="tozeroy",
            fillcolor=fill_color_4,
            hovertemplate="%{x|%Y-%m-%d}<br>Top 4: %{y:.1f}%<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot[col_8],
            mode="lines",
            name="Top 8",
            line=dict(color=line_color_8, width=2),
            fill="tozeroy",
            fillcolor=fill_color_8,
            hovertemplate="%{x|%Y-%m-%d}<br>Top 8: %{y:.1f}%<extra></extra>",
        )
    )

    fig.update_layout(
        title=title,
        height=420,
        margin=dict(l=20, r=20, t=60, b=20),
        hovermode="x unified",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        yaxis=dict(title=yaxis_title, tickformat=tickformat),
        xaxis=dict(title=""),
    )

    return fig


def _build_net_pct_oi_figure(
    df_plot: pd.DataFrame,
    commodity: str,
    report_type: str,
) -> go.Figure:
    fig = go.Figure()

    for trader in TRADER_ORDER:
        trader_key = trader.lower().replace(" ", "_").replace("-", "_")
        pct_col = f"{trader_key}_pct_oi"
        net_col = f"{trader_key}_net"

        if pct_col not in df_plot.columns:
            continue

        pct_vals = pd.to_numeric(df_plot[pct_col], errors="coerce")
        if pct_vals.dropna().empty:
            continue

        customdata = pd.DataFrame(
            {
                "net": pd.to_numeric(df_plot[net_col], errors="coerce"),
                "pct": pct_vals,
            }
        ).to_numpy()

        fig.add_trace(
            go.Scatter(
                x=df_plot["report_date"],
                y=pct_vals.clip(lower=0),
                mode="lines",
                name=trader,
                stackgroup="positive",
                line=dict(width=1.6, color=TRADER_COLORS[trader], shape="spline", smoothing=0.65),
                fillcolor=TRADER_COLORS[trader],
                customdata=customdata,
                hovertemplate=(
                    "%{x|%Y-%m-%d}<br>"
                    f"{trader}: "
                    "%{customdata[0]:,.0f} contracts"
                    "<br>% of OI: %{customdata[1]:.2f}%"
                    "<extra></extra>"
                ),
            )
        )
        fig.add_trace(
            go.Scatter(
                x=df_plot["report_date"],
                y=pct_vals.clip(upper=0),
                mode="lines",
                name=trader,
                stackgroup="negative",
                line=dict(width=1.6, color=TRADER_COLORS[trader], shape="spline", smoothing=0.65),
                fillcolor=TRADER_COLORS[trader],
                customdata=customdata,
                hovertemplate=(
                    "%{x|%Y-%m-%d}<br>"
                    f"{trader}: "
                    "%{customdata[0]:,.0f} contracts"
                    "<br>% of OI: %{customdata[1]:.2f}%"
                    "<extra></extra>"
                ),
                showlegend=False,
            )
        )

    fig.update_layout(
        title=f"{commodity} Net Position as % of Open Interest ({report_type})",
        height=430,
        margin=dict(l=20, r=20, t=60, b=20),
        hovermode="x unified",
        template="plotly_white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        paper_bgcolor="#FAF7F2",
        plot_bgcolor="#FFFDF9",
        yaxis=dict(
            title="% of Open Interest",
            ticksuffix="%",
            gridcolor="rgba(148, 163, 184, 0.18)",
            zeroline=False,
        ),
        xaxis=dict(title=""),
    )
    fig.update_xaxes(gridcolor="rgba(148, 163, 184, 0.10)")
    fig.add_hline(y=0, line_color="rgba(71, 85, 105, 0.55)", line_width=1.2)

    return fig


def render_concentration():
    st.title("Concentration")

    commodity = st.selectbox(
        "Commodity",
        list(COMMODITY_SHEETS.keys()),
        key="concentration_commodity",
    )

    report_type = st.radio(
        "Report Type",
        ["Futures Only", "Futures + Options"],
        horizontal=True,
        key="concentration_report_type",
    )

    futs_only = report_type == "Futures Only"
    raw_df = load_commodity_df(commodity, futs_only)
    df = _prepare_concentration_df(raw_df)
    df_net_pct = _prepare_net_pct_oi_df(raw_df)

    if df.empty:
        st.warning("No concentration data is available for this selection.")
        return

    concentration_cols = [
        "long_concentration_4",
        "long_concentration_8",
        "short_concentration_4",
        "short_concentration_8",
    ]

    if df[concentration_cols].dropna(how="all").empty:
        st.warning(
            "The concentration columns were not found for this commodity/report type yet."
        )
        return

    years = sorted(df["year"].dropna().unique().tolist())
    min_year, max_year = int(min(years)), int(max(years))

    selected_years = st.slider(
        "Years included",
        min_value=min_year,
        max_value=max_year,
        value=(max(min_year, max_year - 5), max_year),
        step=1,
        key="concentration_year_slider",
    )

    st.subheader("How big are the top players?")
    st.caption(
        "This view tracks concentration over time for the largest 4 and largest 8 traders on the long and short sides.")

    start_year, end_year = selected_years
    df_plot = df[(df["year"] >= start_year) & (df["year"] <= end_year)].copy()

    if df_plot.empty:
        st.warning("No data for the selected year range.")
        return

    df_net_pct_plot = df_net_pct[
        (df_net_pct["year"] >= start_year) & (df_net_pct["year"] <= end_year)
    ].copy()

    long_fig = _build_side_figure(
        df_plot=df_plot,
        title=f"{commodity} Long Concentration ({report_type})",
        col_4="long_concentration_4",
        col_8="long_concentration_8",
        line_color_4="#1D4ED8",
        fill_color_4="rgba(29, 78, 216, 0.30)",
        line_color_8="#93C5FD",
        fill_color_8="rgba(147, 197, 253, 0.45)",
    )
    short_fig = _build_side_figure(
        df_plot=df_plot,
        title=f"{commodity} Short Concentration ({report_type})",
        col_4="short_concentration_4",
        col_8="short_concentration_8",
        line_color_4="#B45309",
        fill_color_4="rgba(180, 83, 9, 0.28)",
        line_color_8="#FCD34D",
        fill_color_8="rgba(252, 211, 77, 0.45)",
    )

    left_chart, right_chart = st.columns(2)
    with left_chart:
        st.plotly_chart(long_fig, use_container_width=True)
    with right_chart:
        st.plotly_chart(short_fig, use_container_width=True)

    latest = df_plot.dropna(
        subset=["long_concentration_4", "long_concentration_8", "short_concentration_4", "short_concentration_8"],
        how="all",
    ).tail(1)
    if not latest.empty:
        row = latest.iloc[0]
        metric_1, metric_2, metric_3, metric_4 = st.columns(4)
        metric_1.metric("Latest Long Top 4", f"{row['long_concentration_4']:.1f}%")
        metric_2.metric("Latest Long Top 8", f"{row['long_concentration_8']:.1f}%")
        metric_3.metric("Latest Short Top 4", f"{row['short_concentration_4']:.1f}%")
        metric_4.metric("Latest Short Top 8", f"{row['short_concentration_8']:.1f}%")    
    
    net_pct_cols = [f"{trader.lower().replace(' ', '_').replace('-', '_')}_pct_oi" for trader in TRADER_ORDER]
    if not df_net_pct_plot.empty and not df_net_pct_plot[net_pct_cols].dropna(how="all").empty:
        st.subheader("How big is each category in net terms?")
        st.caption(
            "Net position is calculated as long minus short for each trader category, then divided by total open interest. Spread positions are excluded."
        )
        net_pct_fig = _build_net_pct_oi_figure(
            df_plot=df_net_pct_plot,
            commodity=commodity,
            report_type=report_type,
        )
        st.plotly_chart(net_pct_fig, use_container_width=True)

