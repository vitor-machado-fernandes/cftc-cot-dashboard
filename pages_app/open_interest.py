import streamlit as st
import pandas as pd
from pathlib import Path
import plotly.express as px
from plotly.colors import qualitative
import plotly.graph_objects as go
from plotly.subplots import make_subplots



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


def build_lower_facet_df(df_futs, df_opt, df_fno, years):
    frames = []

    report_map = {
        "Futures Only": df_futs,
        "Options Only": df_opt,
        "Futures + Options": df_fno,
    }

    bucket_map = {
        "Old": "open_interest_old",
        "Other": "open_interest_other",
    }

    for report_label, df in report_map.items():
        for bucket_label, value_col in bucket_map.items():
            tmp = prepare_seasonal(df, value_col, years).copy()
            tmp["report_type"] = report_label
            tmp["bucket"] = bucket_label
            tmp["value"] = tmp[value_col]
            frames.append(tmp[["doy", "year", "value", "report_type", "bucket"]])

    if not frames:
        return pd.DataFrame(columns=["doy", "year", "value", "report_type", "bucket"])

    out = pd.concat(frames, ignore_index=True)
    out["year"] = out["year"].astype(str)
    return out


def get_year_color_map(years):
    years_sorted = sorted([int(y) for y in years])
    years_str = [str(y) for y in years_sorted]

    # oldest -> lightest, newest prior year -> darkest
    blue_palette = [
        "#dbeafe",  # very light
        "#bfdbfe",
        "#93c5fd",
        "#60a5fa",
        "#2563eb",  # dark blue
    ]

    color_map = {}

    prior_years = years_str[:-1]
    current_year = years_str[-1]

    if len(prior_years) <= len(blue_palette):
        palette = blue_palette[-len(prior_years):]
    else:
        # fallback if user selects many years
        palette = [blue_palette[i % len(blue_palette)] for i in range(len(prior_years))]

    for y, c in zip(prior_years, palette):
        color_map[y] = c

    color_map[current_year] = "black"

    return color_map


def seasonal_oi_chart(df, value_col, years, title, color_map=None, show_legend=True, height=260):
    plot_df = prepare_seasonal(df, value_col, years)

    if plot_df.empty:
        fig = px.line(title=title)
        fig.update_layout(
            height=height,
            margin=dict(l=10, r=10, t=40, b=10),
            showlegend=show_legend,
        )
        fig.add_annotation(
            text="No data available",
            x=0.5, y=0.5,
            xref="paper", yref="paper",
            showarrow=False
        )
        return fig

    plot_df["year"] = plot_df["year"].astype(str)
    current_year = str(max(years))

    if color_map is None:
        color_map = get_year_color_map(years)

    fig = px.line(
        plot_df,
        x="doy",
        y=value_col,
        color="year",
        line_group="year",
        markers=False,
        title=title,
        color_discrete_map=color_map,
    )

    for trace in fig.data:
        if trace.name == current_year:
            trace.line.width = 3
            trace.line.color = "black"

    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        legend_title_text="",
        xaxis_title="",
        yaxis_title="",
        showlegend=show_legend,
    )

    fig.update_xaxes(
        tickmode="array",
        tickvals=[1,32,60,91,121,152,182,213,244,274,305,335],
        ticktext=["Jan","Feb","Mar","Apr","May","Jun",
                  "Jul","Aug","Sep","Oct","Nov","Dec"]
    )

    fig.update_yaxes(tickformat="~s")

    return fig


def render_oi_top_section(df_futs, df_opt, df_fno, years):
    color_map = get_year_color_map(years)

    # Row 1: Futures Only | All
    fig1 = seasonal_oi_chart(
        df=df_futs,
        value_col="open_interest_all",
        years=years,
        title="Futures Only | All",
        color_map=color_map,
        show_legend=True,
        height=320,
    )
    st.plotly_chart(fig1, use_container_width=True)

    # Row 2: Options Only | All   +   Futures + Options | All
    col1, col2 = st.columns(2)

    fig2 = seasonal_oi_chart(
        df=df_opt,
        value_col="open_interest_all",
        years=years,
        title="Options Only | All",
        color_map=color_map,
        show_legend=False,
        height=300,
    )

    fig3 = seasonal_oi_chart(
        df=df_fno,
        value_col="open_interest_all",
        years=years,
        title="Futures + Options | All",
        color_map=color_map,
        show_legend=False,
        height=300,
    )

    with col1:
        st.plotly_chart(fig2, use_container_width=True)

    with col2:
        st.plotly_chart(fig3, use_container_width=True)


def render_oi_lower_facet(df_futs, df_opt, df_fno, years):
    plot_df = build_lower_facet_df(df_futs, df_opt, df_fno, years)
    color_map = get_year_color_map(years)

    ymin = plot_df["value"].min()
    ymax = plot_df["value"].max()

    fig = make_subplots(
        rows=2,
        cols=3,
        subplot_titles=[
            "Futures Only | Old", "Options Only | Old", "Futures + Options | Old",
            "Futures Only | Other", "Options Only | Other", "Futures + Options | Other",
        ],
        shared_xaxes=True,
        horizontal_spacing=0.05,
        vertical_spacing=0.10,
    )

    row_map = {"Old": 1, "Other": 2}
    col_map = {"Futures Only": 1, "Options Only": 2, "Futures + Options": 3}

    added_years = set()

    current_year = str(max(years))
    
    for (bucket, report, year), grp in plot_df.groupby(["bucket", "report_type", "year"]):
        showlegend = year not in added_years

        fig.add_trace(
            go.Scatter(
                x=grp["doy"],
                y=grp["value"],
                mode="lines",
                name=str(year),
                legendgroup=str(year),
                showlegend=showlegend,
                line=dict(
                    color="black" if str(year) == current_year else color_map[str(year)],
                    width=3 if str(year) == current_year else 1.5,
                ),
            ),
            row=row_map[bucket],
            col=col_map[report],)

        added_years.add(year)

    fig.update_xaxes(
        tickmode="array",
        tickvals=[1,32,60,91,121,152,182,213,244,274,305,335],
        ticktext=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    )

    fig.update_yaxes(
    tickformat="~s",
    range=[ymin, ymax])

    for r in [1, 2]:
        fig.update_yaxes(showticklabels=False, row=r, col=2)
        fig.update_yaxes(showticklabels=False, row=r, col=3)

    fig.update_layout(
        height=400,
        margin=dict(l=10, r=10, t=60, b=10),
        legend_title_text="",
    )

    fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")

    st.plotly_chart(fig, use_container_width=True)


def render_oi_matrix(df_futs, df_opt, df_fno, years):
    render_oi_lower_facet(df_futs, df_opt, df_fno, years)


# ----------------------------
# Main render
# ----------------------------
def render_open_interest():
    st.title("Open Interest Tool")
    st.write(
    """The Commimment of Traders report releases open interest (O.I.) data in two formats:  
    - **Futures Only**  
    - **Futures + Options**  
    The options O.I. is the estimated sum of the deltas of all outstanding options. 
    From these two reports it is possible to extract the 'Option Only' open interest.  
    """)

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

    render_oi_top_section(df_futs, df_opt, df_fno, years)

    st.markdown("---")

    st.write("""
    The report also discloses the crop year of the positions:  
    - **Old** = Old Crop  
    - **Other** = New Crop(s)  
    - **All** = old + new crop
             """)

    render_oi_matrix(df_futs, df_opt, df_fno, years)