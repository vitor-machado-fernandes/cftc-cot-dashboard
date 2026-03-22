# pages_app/n_traders.py

import pandas as pd
import streamlit as st
from plotly.subplots import make_subplots
import plotly.graph_objects as go

from pages_app.CoT_position import load_commodity_df, COMMODITY_SHEETS


# ------------------------------------------------
# Helpers
# ------------------------------------------------
def _clean_numeric(s):
    return pd.to_numeric(
        s.astype(str)
         .str.replace(",", "", regex=False)
         .str.strip()
         .replace({"": None, "nan": None, "None": None}),
        errors="coerce"
    )


def _prepare_ntraders_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    out["report_date"] = pd.to_datetime(out["report_date"], errors="coerce")
    out = out.dropna(subset=["report_date"]).copy()
    out["year"] = out["report_date"].dt.year

    needed_cols = [
        "traders_tot_all",
        "traders_tot_rept_long_all",
        "traders_tot_rept_short_all",
        "traders_prod_merc_long_all",
        "traders_prod_merc_short_all",
        "traders_swap_long_all",
        "traders_swap_short_all",
        "traders_swap_spread_all",
        "traders_m_money_long_all",
        "traders_m_money_short_all",
        "traders_m_money_spread_all",
        "traders_other_rept_long_all",
        "traders_other_rept_short",
        "traders_other_rept_spread",
    ]

    for col in needed_cols:
        if col in out.columns:
            out[col] = _clean_numeric(out[col])

    return out


def _build_total_traders_figure(df_plot: pd.DataFrame):
    COLORS = {
        "total": "#0F2647",
        "long": "#3769B4",
        "short": "#96BEE6",
    }

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_tot_all"],
            mode="lines",
            name="Total",
            line=dict(color=COLORS["total"], width=3),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Total: %{y:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_tot_rept_long_all"],
            mode="lines",
            name="Long",
            line=dict(color=COLORS["long"], width=2),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Long: %{y:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_tot_rept_short_all"],
            mode="lines",
            name="Short",
            line=dict(color=COLORS["short"], width=2),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Short: %{y:,.0f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="Total Number of Traders",
        height=360,
        margin=dict(l=0, r=0, t=50, b=80),
        hovermode="x unified",
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            title_text="",
        ),
        yaxis=dict(title="Traders"),
        xaxis=dict(title=""),
    )

    return fig


def _build_spread_traders_figure(df_plot: pd.DataFrame):
    COLORS = {
        "spread_swap": "#6B7280",
        "spread_mm": "#9CA3AF",
        "spread_other": "#D1D5DB",
    }

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_swap_spread_all"],
            mode="lines",
            name="Swap Dealers",
            line=dict(color=COLORS["spread_swap"], width=2),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Swap Dealers: %{y:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_m_money_spread_all"],
            mode="lines",
            name="Managed Money",
            line=dict(color=COLORS["spread_mm"], width=2),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Managed Money: %{y:,.0f}<extra></extra>",
        )
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_other_rept_spread"],
            mode="lines",
            name="Other Reportables",
            line=dict(color=COLORS["spread_other"], width=2),
            hovertemplate="Date %{x|%Y-%m-%d}<br>Other Reportables: %{y:,.0f}<extra></extra>",
        )
    )

    fig.update_layout(
        title="Number of Spread Traders",
        height=360,
        margin=dict(l=0, r=0, t=50, b=80),
        hovermode="x unified",
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            title_text="",
        ),
        yaxis=dict(title="Traders"),
        xaxis=dict(title=""),
    )

    return fig


def _build_faceted_figure(df_plot: pd.DataFrame):
    fig = make_subplots(
        rows=2,
        cols=2,
        shared_xaxes=True,
        vertical_spacing=0.14,
        horizontal_spacing=0.10,
        subplot_titles=[
            "PMPU",
            "Swap Dealers",
            "Managed Money",
            "Other Reportables",
        ],
    )

    # Professional-looking palette
    COLORS = {
        "total": "#0F2647",        # dark blue
        "long": "#3769B4",         # ligher blue
        "short": "#96BEE6",        # very light blue
        "spread_swap": "#6B7280",  # grey
        "spread_mm": "#9CA3AF",    # lighter grey
        "spread_other": "#D1D5DB", # very light grey
    }

    # ---------- Row 1, Col 1 : PMPU ----------
    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_prod_merc_long_all"],
            mode="lines",
            name="Long",
            line=dict(color=COLORS["long"], width=2),
            legendgroup="long",
            showlegend=True,
        ),
        row=1, col=1
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_prod_merc_short_all"],
            mode="lines",
            name="Short",
            line=dict(color=COLORS["short"], width=2),
            legendgroup="short",
            showlegend=True,
        ),
        row=1, col=1
    )

    # ---------- Row 1, Col 2 : Swap ----------
    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_swap_long_all"],
            mode="lines",
            name="Long",
            line=dict(color=COLORS["long"], width=2),
            legendgroup="long",
            showlegend=False,
        ),
        row=1, col=2
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_swap_short_all"],
            mode="lines",
            name="Short",
            line=dict(color=COLORS["short"], width=2),
            legendgroup="short",
            showlegend=False,
        ),
        row=1, col=2
    )

    # ---------- Row 2, Col 1 : Managed Money ----------
    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_m_money_long_all"],
            mode="lines",
            name="Long",
            line=dict(color=COLORS["long"], width=2),
            legendgroup="long",
            showlegend=False,
        ),
        row=2, col=1
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_m_money_short_all"],
            mode="lines",
            name="Short",
            line=dict(color=COLORS["short"], width=2),
            legendgroup="short",
            showlegend=False,
        ),
        row=2, col=1
    )

    # ---------- Row 2, Col 2 : Other Reportables ----------
    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_other_rept_long_all"],
            mode="lines",
            name="Long",
            line=dict(color=COLORS["long"], width=2),
            legendgroup="long",
            showlegend=False,
        ),
        row=2, col=2
    )

    fig.add_trace(
        go.Scatter(
            x=df_plot["report_date"],
            y=df_plot["traders_other_rept_short"],
            mode="lines",
            name="Short",
            line=dict(color=COLORS["short"], width=2),
            legendgroup="short",
            showlegend=False,
        ),
        row=2, col=2
    )


    # Independent y-axes
    fig.update_yaxes(matches=None)

    fig.update_layout(
        height=700,
        margin=dict(l=20, r=20, t=90, b=40),
        showlegend=True,
        hovermode="x unified",
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5,
            title_text="",
        ),
    )

    # Clean axis labels
    for r in [1, 2]:
        for c in [1, 2]:
            fig.update_yaxes(title_text="Traders", row=r, col=c)
            fig.update_xaxes(title_text="", row=r, col=c)
            fig.update_xaxes(showticklabels=True, row=r, col=c)

    return fig


# ------------------------------------------------
# Main render
# ------------------------------------------------
def render_n_traders():
    st.title("Number of Traders")

    commodity = st.selectbox(
        "Commodity",
        list(COMMODITY_SHEETS.keys()),
        key="ntraders_commodity",
    )

    st.write("""
    Threshold, in contracts, for a reportable position:  
             - **Corn**: 250 lots  
             - **Soybeans/SBO/SBM**: 200 lots  
             - **Wheat**: 150 lots  
             - **Cotton**: 1000 lots  

    """)

    report_type = st.radio(
        "Report Type",
        ["Futures Only", "Futures + Options"],
        horizontal=True,
        key="ntraders_report_type",
    )

    futs_only = report_type == "Futures Only"
    df = load_commodity_df(commodity, futs_only)
    df = _prepare_ntraders_df(df)

    if df.empty:
        st.warning("No data available.")
        return

    years = sorted(df["year"].dropna().unique().tolist())
    min_year, max_year = int(min(years)), int(max(years))

    selected_years = st.slider(
        "Years included",
        min_value=min_year,
        max_value=max_year,
        value=(max(min_year, max_year - 5), max_year),
        step=1,
        key="ntraders_year_slider",
    )

    start_year, end_year = selected_years
    df_plot = df[(df["year"] >= start_year) & (df["year"] <= end_year)].copy()

    if df_plot.empty:
        st.warning("No data for the selected year range.")
        return

    top_left, top_right = st.columns(2)

    with top_left:
        fig_total = _build_total_traders_figure(df_plot)
        st.plotly_chart(fig_total, use_container_width=True)

    with top_right:
        fig_spread = _build_spread_traders_figure(df_plot)
        st.plotly_chart(fig_spread, use_container_width=True)

    st.subheader("Number of players by type")
    
    fig = _build_faceted_figure(df_plot)
    st.plotly_chart(fig, use_container_width=True)
