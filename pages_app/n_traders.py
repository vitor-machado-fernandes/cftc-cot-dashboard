# n_traders.py

import pandas as pd
import plotly.express as px
import streamlit as st


def _prepare_trader_participation(df: pd.DataFrame, report_type: str) -> pd.DataFrame:
    """
    Prepare long-format dataframe for trader participation over time.
    
    report_type options:
    - 'Futures Only'
    - 'Futures + Options'
    - 'Both'
    """

    data = df.copy()

    # Adjust this if your date column has a different name
    data["date"] = pd.to_datetime(data["Report_Date_as_YYYY-MM-DD"])

    fut_only_cols = {
        "Producer/Merchant": "traders_prod_merc_long_old",
        "Swap Dealers": "traders_swap_long_old",
        "Managed Money": "traders_m_money_long_old",
        "Other Reportables": "traders_other_rept_long_old",
    }

    fut_opt_cols = {
        "Producer/Merchant": "traders_prod_merc_long_all",
        "Swap Dealers": "traders_swap_long_all",
        "Managed Money": "traders_m_money_long_all",
        "Other Reportables": "traders_other_rept_long_all",
    }

    frames = []

    if report_type in ["Futures Only", "Both"]:
        tmp = data[["date"] + list(fut_only_cols.values())].rename(columns=fut_only_cols)
        tmp = tmp.melt(id_vars="date", var_name="Category", value_name="Traders")
        tmp["Report"] = "Futures Only"
        frames.append(tmp)

    if report_type in ["Futures + Options", "Both"]:
        tmp = data[["date"] + list(fut_opt_cols.values())].rename(columns=fut_opt_cols)
        tmp = tmp.melt(id_vars="date", var_name="Category", value_name="Traders")
        tmp["Report"] = "Futures + Options"
        frames.append(tmp)

    out = pd.concat(frames, ignore_index=True)
    out = out.dropna(subset=["Traders"])
    return out


def render_n_traders(df: pd.DataFrame):
    st.subheader("Number of Traders")

    report_type = st.selectbox(
        "Report Type",
        ["Futures Only", "Futures + Options", "Both"],
        index=0,
        key="ntraders_report_type"
    )

    chart_df = _prepare_trader_participation(df, report_type)

    if chart_df.empty:
        st.warning("No data available for this selection.")
        return

    if report_type == "Both":
        fig = px.line(
            chart_df,
            x="date",
            y="Traders",
            color="Category",
            line_dash="Report",
            labels={"date": "", "Traders": "Number of Traders"},
        )
    else:
        fig = px.line(
            chart_df,
            x="date",
            y="Traders",
            color="Category",
            labels={"date": "", "Traders": "Number of Traders"},
        )

    fig.update_layout(
        height=500,
        margin=dict(l=20, r=20, t=20, b=20),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="Number of Traders",
    )

    fig.update_traces(mode="lines", line=dict(width=2))

    st.plotly_chart(fig, use_container_width=True)
