import streamlit as st
import pandas as pd
import plotly.express as px

from pathlib import Path
from CoT_position import load_commodity_df, COMMODITY_SHEETS


def _prepare_trader_participation(df: pd.DataFrame, futs_only: bool):

    if futs_only:

        cols = {
            "Total": "traders_tot_old",
            "Producer/Merchant": "traders_prod_merc_long_old",
            "Swap Dealers": "traders_swap_long_old",
            "Managed Money": "traders_m_money_long_old",
            "Other Reportables": "traders_other_rept_long_old",
        }

    else:

        cols = {
            "Total": "traders_tot_all",
            "Producer/Merchant": "traders_prod_merc_long_all",
            "Swap Dealers": "traders_swap_long_all",
            "Managed Money": "traders_m_money_long_all",
            "Other Reportables": "traders_other_rept_long_all",
        }

    data = df[["report_date"] + list(cols.values())].rename(columns=cols)

    long_df = data.melt(
        id_vars="report_date",
        var_name="Category",
        value_name="Traders"
    )

    return long_df


def render_n_traders():

    st.title("Number of Traders")

    # --- Commodity selector ---
    commodity = st.selectbox(
        "Commodity",
        list(COMMODITY_SHEETS.keys()),
        key="ntraders_commodity"
    )

    # --- Report type ---
    report_type = st.radio(
        "Report Type",
        ["Futures Only", "Futures + Options"],
        horizontal=True,
        key="ntraders_report"
    )

    futs_only = report_type == "Futures Only"

    df = load_commodity_df(commodity, futs_only)

    chart_df = _prepare_trader_participation(df, futs_only)

    fig = px.line(
        chart_df,
        x="report_date",
        y="Traders",
        color="Category"
    )

    fig.update_layout(
        height=550,
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis_title="",
        yaxis_title="Number of Traders",
        legend_title=""
    )

    fig.update_traces(line=dict(width=2))

    st.plotly_chart(fig, use_container_width=True)
