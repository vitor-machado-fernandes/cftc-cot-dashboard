import streamlit as st
import pandas as pd
from pathlib import Path
import plotly.express as px

def render_open_interest():
    st.title("Open Interest")
    st.write("Insert tab description here")

# ----------------------------
# Config
# ----------------------------
DATA_DIR = Path(".")
FUTS_FILE = DATA_DIR / "CoT_Disagg_FutsOnly.xlsx"
FNO_FILE = DATA_DIR / "CoT_Disagg_FnO.xlsx"

COMMODITY_SHEETS = {
    "Corn": "Corn",
    "Cotton": "Cotton",
    "Soybeans": "Soybeans",
    "SBO": "SBO",
    "SBM": "SBM",
}

OI_COLS = [
    "report_date_as_yyyy_mm_dd",
    "open_interest_all",
    "open_interest_old",
    "open_interest_other",
]

REPORT_LABELS = {
    "futs": "Futures Only",
    "opt": "Options Only",
    "fno": "Futures + Options",
}

ROW_LABELS = {
    "all": "All",
    "old": "Old",
    "other": "Other",
}

# ----------------------------
# Loaders
# ----------------------------
@st.cache_data
def load_oi_sheet(filepath, sheet_name):
    df = pd.read_excel(filepath, sheet_name=sheet_name, engine="openpyxl")
    df = df.copy()

    # keep only needed columns
    keep = [c for c in OI_COLS if c in df.columns]
    df = df[keep]

    # parse date
    df["report_date_as_yyyy_mm_dd"] = pd.to_datetime(
        df["report_date_as_yyyy_mm_dd"], errors="coerce"
    )

    # numeric columns may come as strings
    for col in ["open_interest_all", "open_interest_old", "open_interest_other"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=["report_date_as_yyyy_mm_dd"]).sort_values("report_date_as_yyyy_mm_dd")
    df["year"] = df["report_date_as_yyyy_mm_dd"].dt.year
    df["month"] = df["report_date_as_yyyy_mm_dd"].dt.month
    df["month_label"] = df["report_date_as_yyyy_mm_dd"].dt.strftime("%b")

    return df


def make_options_only(df_fno, df_futs):
    """
    Options only = (Futures + Options) - (Futures only)
    Merge by report date.
    """
    cols = ["report_date_as_yyyy_mm_dd", "open_interest_all", "open_interest_old", "open_interest_other"]

    x = df_fno[cols].copy()
    y = df_futs[cols].copy()

    df = x.merge(
        y,
        on="report_date_as_yyyy_mm_dd",
        how="inner",
        suffixes=("_fno", "_futs")
    )

    out = pd.DataFrame()
    out["report_date_as_yyyy_mm_dd"] = df["report_date_as_yyyy_mm_dd"]

    for bucket in ["all", "old", "other"]:
        out[f"open_interest_{bucket}"] = (
            df[f"open_interest_{bucket}_fno"] - df[f"open_interest_{bucket}_futs"]
        )

    out["year"] = out["report_date_as_yyyy_mm_dd"].dt.year
    out["month"] = out["report_date_as_yyyy_mm_dd"].dt.month
    out["month_label"] = out["report_date_as_yyyy_mm_dd"].dt.strftime("%b")

    return out.sort_values("report_date_as_yyyy_mm_dd")

def prepare_seasonal(df, value_col, years):
    out = df[df["year"].isin(years)].copy()
    out = out[["report_date_as_yyyy_mm_dd", "year", "month", "month_label", value_col]].dropna()

# day of year for seasonal alignment
    out["doy"] = out["report_date_as_yyyy_mm_dd"].dt.dayofyear

    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    out["month_label"] = pd.Categorical(out["month_label"], categories=month_order, ordered=True)

    return out

def seasonal_oi_chart(df, value_col, years, title):
    plot_df = prepare_seasonal(df, value_col, years)

    if plot_df.empty:
        fig = px.line(title=title)
        fig.update_layout(
            height=260,
            margin=dict(l=10, r=10, t=40, b=10),
        )
        fig.add_annotation(
            text="No data available",
            x=0.5, y=0.5,
            xref="paper", yref="paper",
            showarrow=False
        )
        return fig

    # Put selected years in ascending order for cleaner legend
    plot_df["year"] = plot_df["year"].astype(str)

    fig = px.line(
        plot_df,
        x="doy",
        y=value_col,
        color="year",
        line_group="year",
        markers=False,
        title=title,
    )

    fig.update_layout(
        height=260,
        margin=dict(l=10, r=10, t=40, b=10),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
    )

    fig.update_xaxes(
    tickmode="array",
    tickvals=[1,32,60,91,121,152,182,213,244,274,305,335],
    ticktext=["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"])

    fig.update_xaxes(
        categoryorder="array",
        categoryarray=["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    )

    fig.update_yaxes(tickformat="~s")

    return fig

def render_oi_matrix(df_futs, df_opt, df_fno, years):
    report_dfs = {
        "futs": df_futs,
        "opt": df_opt,
        "fno": df_fno,
    }

    row_keys = ["all", "old", "other"]
    col_keys = ["futs", "opt", "fno"]

    for row_key in row_keys:
        cols = st.columns(3)

        for i, col_key in enumerate(col_keys):
            value_col = f"open_interest_{row_key}"
            title = f"{REPORT_LABELS[col_key]} | {ROW_LABELS[row_key]}"

            fig = seasonal_oi_chart(
                df=report_dfs[col_key],
                value_col=value_col,
                years=years,
                title=title,
            )

            with cols[i]:
                st.plotly_chart(fig, use_container_width=True)

# ----------------------------
# Main render
# ----------------------------
def render_open_interest():
    st.title("Open Interest")

    # ---- Controls
    commodity = st.selectbox("Commodity", list(COMMODITY_SHEETS.keys()), index=1)

    current_year = pd.Timestamp.today().year
    all_years = list(range(current_year, 2008, -1))
    default_years = [y for y in range(current_year, current_year - 6, -1)]

    years = st.multiselect(
        "Years",
        options=all_years,
        default=default_years
    )

    if not years:
        st.warning("Please select at least one year.")
        return

    sheet = COMMODITY_SHEETS[commodity]

    # ---- Load data
    df_futs = load_oi_sheet(FUTS_FILE, sheet)
    df_fno = load_oi_sheet(FNO_FILE, sheet)
    df_opt = make_options_only(df_fno, df_futs)

    render_oi_matrix(df_futs, df_opt, df_fno, years)