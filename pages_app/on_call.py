from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


ON_CALL_PARQUET = Path("Cotton_OnCall.parquet")


@st.cache_data
def _load_on_call_df_cached(
    path_str: str,
    modified_time_ns: int,
    file_size: int,
) -> pd.DataFrame:
    del modified_time_ns, file_size
    parquet_path = Path(path_str)
    df = pd.read_parquet(parquet_path)
    for col in ["report_date", "release_date", "futures_contract_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def load_on_call_df(path: str | Path = ON_CALL_PARQUET) -> pd.DataFrame:
    parquet_path = Path(path)
    if not parquet_path.exists():
        return pd.DataFrame()

    stat = parquet_path.stat()
    return _load_on_call_df_cached(
        str(parquet_path),
        stat.st_mtime_ns,
        stat.st_size,
    )


def _prepare_on_call_total_series(
    df: pd.DataFrame,
    source_col: str,
    output_col: str,
) -> pd.DataFrame:
    out = df.copy()
    out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")
    out = out.dropna(subset=["report_date"]).copy()

    if "is_total" in out.columns and out["is_total"].any():
        out = out[out["is_total"]].copy()
        out[output_col] = pd.to_numeric(out[source_col], errors="coerce")
    else:
        out[source_col] = pd.to_numeric(out[source_col], errors="coerce")
        out = (
            out.groupby("report_date", as_index=False)[source_col]
            .sum(min_count=1)
            .rename(columns={source_col: output_col})
        )

    out["year"] = out["report_date"].dt.year
    out["week_of_year"] = ((out["report_date"].dt.dayofyear - 1) // 7) + 1

    return out[["report_date", "year", "week_of_year", output_col]].sort_values("report_date")


def _prepare_unfixed_call_sales_df(df: pd.DataFrame) -> pd.DataFrame:
    return _prepare_on_call_total_series(
        df=df,
        source_col="unfixed_call_sales",
        output_col="unfixed_call_sales_total",
    )


def _prepare_unfixed_call_purchases_df(df: pd.DataFrame) -> pd.DataFrame:
    return _prepare_on_call_total_series(
        df=df,
        source_col="unfixed_call_purchases",
        output_col="unfixed_call_purchases_total",
    )


def _prepare_latest_contract_month_table(df: pd.DataFrame) -> tuple[pd.Timestamp | None, pd.DataFrame]:
    out = df.copy()
    out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")
    out["futures_contract_date"] = pd.to_datetime(out["futures_contract_date"], errors="coerce")
    out = out.dropna(subset=["report_date"]).copy()

    if "is_total" in out.columns:
        out = out[~out["is_total"]].copy()

    latest_report_date = out["report_date"].max() if not out.empty else None
    if pd.isna(latest_report_date):
        return None, pd.DataFrame()

    latest = out[out["report_date"] == latest_report_date].copy()
    latest["unfixed_call_sales"] = pd.to_numeric(latest["unfixed_call_sales"], errors="coerce")
    latest["sales_change_from_previous_week"] = pd.to_numeric(
        latest["sales_change_from_previous_week"], errors="coerce"
    )
    latest["unfixed_call_purchases"] = pd.to_numeric(
        latest["unfixed_call_purchases"], errors="coerce"
    )
    latest["purchases_change_from_previous_week"] = pd.to_numeric(
        latest["purchases_change_from_previous_week"], errors="coerce"
    )

    latest = latest.sort_values(["futures_contract_date", "contract_month"], na_position="last")

    display_df = latest[
        [
            "contract_month",
            "unfixed_call_sales",
            "sales_change_from_previous_week",
            "unfixed_call_purchases",
            "purchases_change_from_previous_week",
        ]
    ].rename(
        columns={
            "contract_month": "Contract Month",
            "unfixed_call_sales": "Sales",
            "sales_change_from_previous_week": "WoW Change Sales",
            "unfixed_call_purchases": "Purchases",
            "purchases_change_from_previous_week": "WoW Change Purchases",
        }
    )

    return latest_report_date, display_df


def _get_year_color_map(selected_years: list[int], palette: list[str]) -> dict[int, str]:
    return {year: palette[idx % len(palette)] for idx, year in enumerate(sorted(selected_years, reverse=True))}


def _build_on_call_seasonal_chart(
    df_plot: pd.DataFrame,
    selected_years: list[int],
    value_col: str,
    title: str,
    hover_label: str,
    palette: list[str],
) -> go.Figure:
    fig = go.Figure()
    color_map = _get_year_color_map(selected_years, palette)
    current_year = max(selected_years) if selected_years else None

    for year in sorted(selected_years, reverse=True):
        tmp = df_plot[df_plot["year"] == year].dropna(subset=[value_col]).copy()
        if tmp.empty:
            continue

        is_current_year = year == current_year

        fig.add_trace(
            go.Scatter(
                x=tmp["week_of_year"],
                y=tmp[value_col],
                mode="lines",
                name=str(year),
                line=dict(
                    color="#111111" if is_current_year else color_map[year],
                    width=3.4 if is_current_year else 2,
                ),
                opacity=1.0 if is_current_year else 0.62,
                hovertemplate=(
                    f"{year}<br>"
                    "Week %{x}<br>"
                    f"{hover_label}: "
                    "%{y:,.0f}<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        title=title,
        height=520,
        margin=dict(l=10, r=10, t=50, b=10),
        template="plotly_white",
        hovermode="x unified",
        legend=dict(
            title="",
            orientation="v",
            yanchor="top",
            y=0.98,
            xanchor="left",
            x=0.01,
            bgcolor="rgba(255,255,255,0.65)",
        ),
        xaxis=dict(
            title="",
            tickmode="array",
            tickvals=[1, 5, 9, 14, 18, 22, 27, 31, 36, 40, 44, 49],
            ticktext=["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            range=[1, 53],
            gridcolor="rgba(148, 163, 184, 0.16)",
        ),
        yaxis=dict(
            title="Contracts",
            tickformat=",d",
            gridcolor="rgba(148, 163, 184, 0.16)",
        ),
    )

    return fig


def render_on_call():
    st.title("On-Call")

    df = load_on_call_df()
    if df.empty:
        st.warning(
            "No local cotton on-call parquet was found yet. Build `Cotton_OnCall.parquet` first, then we can layer charts on top."
        )
        return

    detail_df = df[~df["is_total"]].copy() if "is_total" in df.columns else df.copy()
    latest_report_date = detail_df["report_date"].max() if "report_date" in detail_df.columns else pd.NaT

    st.write(
        """
        The Cotton On-Call report is a CFTC report **specific to cotton**. It tracks unfixed on-call sales and purchases (from a trading house's point of view) in the physical cotton market that are priced against ICE New York cotton futures.

        It is released on Thursdays and reflects positions as of the prior Friday close. In practice, it shows how much cotton remains to be priced later "on call" against specific futures months.
        """
    )

    metric_1, _ = st.columns([1, 4])
    metric_1.metric(
        "Latest Report",
        latest_report_date.strftime("%Y-%m-%d") if pd.notna(latest_report_date) else "N/A",
    )

    sales_df = _prepare_unfixed_call_sales_df(df)
    purchases_df = _prepare_unfixed_call_purchases_df(df)
    available_years = sorted(sales_df["year"].dropna().unique().tolist())

    if not available_years:
        st.warning("No unfixed call sales history is available yet.")
        return

    default_years = available_years[-6:] if len(available_years) >= 6 else available_years

    selected_years = st.multiselect(
        "Years",
        options=available_years,
        default=default_years,
        key="on_call_sales_years",
    )

    if not selected_years:
        st.info("Select at least one year to draw the seasonal chart.")
        return

    plot_df = sales_df[sales_df["year"].isin(selected_years)].copy()
    sales_chart = _build_on_call_seasonal_chart(
        df_plot=plot_df,
        selected_years=selected_years,
        value_col="unfixed_call_sales_total",
        title="Weekly Unfixed Call Position Sales",
        hover_label="Unfixed Call Sales",
        palette=[
            "#7F2704",
            "#A63603",
            "#D94801",
            "#F16913",
            "#FD8D3C",
            "#FDAE6B",
            "#FDD0A2",
            "#FFE6CC",
        ],
    )
    st.plotly_chart(sales_chart, use_container_width=True)

    purchases_plot_df = purchases_df[purchases_df["year"].isin(selected_years)].copy()
    purchases_chart = _build_on_call_seasonal_chart(
        df_plot=purchases_plot_df,
        selected_years=selected_years,
        value_col="unfixed_call_purchases_total",
        title="Weekly Unfixed Call Position Purchases",
        hover_label="Unfixed Call Purchases",
        palette=[
            "#0B3C5D",
            "#145DA0",
            "#1D70A2",
            "#2E86AB",
            "#5FA8D3",
            "#8ECAE6",
            "#BDE0FE",
            "#DCEEF8",
        ],
    )
    st.plotly_chart(purchases_chart, use_container_width=True)

    latest_snapshot_date, latest_snapshot_df = _prepare_latest_contract_month_table(df)
    if not latest_snapshot_df.empty:
        st.subheader("Latest Contract-Month Snapshot")
        st.caption(
            f"Latest report as of {latest_snapshot_date:%Y-%m-%d}: unfixed sales and purchases by contract month, plus week-over-week change."
        )

        formatted_df = latest_snapshot_df.copy()
        for col in [
            "Sales",
            "WoW Change Sales",
            "Purchases",
            "WoW Change Purchases",
        ]:
            formatted_df[col] = formatted_df[col].map(
                lambda x: f"{x:+,.0f}" if "WoW Change" in col and pd.notna(x) else (f"{x:,.0f}" if pd.notna(x) else "")
            )

        left, center, right = st.columns([1, 2.2, 1])
        with center:
            st.dataframe(formatted_df, use_container_width=True, hide_index=True)
