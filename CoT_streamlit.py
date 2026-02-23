import os
import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

st.markdown(
    """
    <style>
        .block-container {
            max-width: 1000px;
            padding-left: 2rem;
            padding-right: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------
st.set_page_config(layout="wide")
st.title("CFTC Commitment of Traders Dashboard")


# ---------------------------------------------------------------------
# Load Data
# ---------------------------------------------------------------------
@st.cache_data
def load_data():
    return pd.read_excel("CFTC_Disaggregated_COT_Ags.xlsx",
                         sheet_name="Cotton")
df = load_data()

df["date"] = pd.to_datetime(df["report_date_as_yyyy_mm_dd"])
df["year"] = df["date"].dt.year
df["seasonal_date"] = pd.to_datetime(
    "2000-" + df["date"].dt.strftime("%m-%d")
)



# ---------------------------------------------------------------------
# Sidebar Controls
# ---------------------------------------------------------------------
st.sidebar.header("Chart Controls")

available_years = sorted(df["year"].unique())
current_year = max(available_years)

years_selected = st.sidebar.multiselect(
    "Years",
    available_years,
    default=available_years[-6:]
)

trader_map = {
    "PMPU": {
        "long": "prod_merc_positions_long",
        "short": "prod_merc_positions_short",
        "spread": None
    },
    "Swap Dealer": {
        "long": "swap_positions_long_all",
        "short": "swap__positions_short_all",
        "spread": "swap__positions_spread_all"
    },
    "Money Managers": {
        "long": "m_money_positions_long_all",
        "short": "m_money_positions_short_all",
        "spread": "m_money_positions_spread"
    },
    "Other Reportables": {
        "long": "other_rept_positions_long",
        "short": "other_rept_positions_short",
        "spread": "other_rept_positions_spread"
    },
    "Nonreportables": {
        "long": "nonrept_positions_long_all",
        "short": "nonrept_positions_short_all",
        "spread": None
    }
}


trader_choice = st.sidebar.selectbox(
    "Trader Type",
    list(trader_map.keys())
)

crop_map = {
    "All": "",
    "Old Crop": "_1",
    "Other Crop": "_2"
}

crop_choice = st.sidebar.selectbox(
    "Crop",
    list(crop_map.keys())
)


# ---------------------------------------------------------------------
# Plotting Position Size
# ---------------------------------------------------------------------
from plotly.subplots import make_subplots
import plotly.graph_objects as go

cols = trader_map[trader_choice]
df_plot = df[df["year"].isin(years_selected)].copy()

# helper to get numeric series (or None)
def series_from_col(colname):
    if colname is None or colname not in df_plot.columns:
        return None
    return pd.to_numeric(df_plot[colname], errors="coerce")

# Compute the 4 “sides”
df_plot["pos_long"] = series_from_col(cols["long"])
df_plot["pos_short"] = series_from_col(cols["short"])
df_plot["pos_net"] = df_plot["pos_long"] - df_plot["pos_short"]

if cols["spread"] is None:
    df_plot["pos_spread"] = pd.NA
else:
    df_plot["pos_spread"] = series_from_col(cols["spread"])

# Make 2x2 subplots
fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=("Long", "Short", "Net", "Spreading"),
    shared_xaxes=True,
    shared_yaxes=False,
    vertical_spacing = 0.1
)

panels = {
    "pos_long": (1, 1),
    "pos_short": (1, 2),
    "pos_net": (2, 1),
    "pos_spread": (2, 2),
}

current_year = max(years_selected)

palette = px.colors.qualitative.Plotly
years_sorted = sorted(years_selected)
color_map = {y: palette[i % len(palette)] for i, y in enumerate(years_sorted)}

for y in sorted(years_selected):
    tmp = df_plot[df_plot["year"] == y].sort_values("seasonal_date")
    is_current = (y == current_year)

    width = 4 if is_current else 1.5
    opacity = 1.0 if is_current else 0.6

    for col, (r, c) in panels.items():
        # skip spreading if unavailable
        if col == "pos_spread" and (cols["spread"] is None):
            continue

        fig.add_trace(
            go.Scatter(
                x=tmp["seasonal_date"],
                y=tmp[col],
                mode="lines",
                name=str(y),
                legendgroup=str(y),
                showlegend=(r == 1 and c == 1),  # legend only once
                line=dict(width=width, color=color_map[y]),
                opacity=opacity,
            ),
            row=r, col=c
        )

# Axes formatting
fig.update_xaxes(dtick="M1", tickformat="%b", showticklabels=True)
fig.update_layout(
    height=700,
    title=f"Cotton – {trader_choice} ({crop_choice})",
    legend_title_text="Year",
    margin=dict(l=40, r=40, t=60, b=40),
)

# If spreading isn't available (PMPU/Nonreportables), add a note
if cols["spread"] is None:
    fig.add_annotation(
        text="Spreading not available for this trader type",
        xref="paper", yref="paper",
        x=0.9225, y=0.15,  # roughly bottom-right panel
        showarrow=False,
        font=dict(size=12)
    )

st.plotly_chart(fig, use_container_width=True)