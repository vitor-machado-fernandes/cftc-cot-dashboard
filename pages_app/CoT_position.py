import pandas as pd
import streamlit as st
from pathlib import Path
import plotly.graph_objects as go


COMMODITY_SHEETS = {
    "Corn": "Corn",
    "Cotton": "Cotton",
    "Soybeans": "Soybeans",
    "SBM": "SBM",
    "SBO": "SBO",
    "Wheat - SRW": "Wheat",
}

TRADER_MAP = {
    "PMPU": "prod_merc",
    "Swap Dealer": "swap",
    "Managed Money": "m_money",
    "Other": "other_rept",
    "Non-Reportables": "nonrept"
}

CROP_SUFFIX = {
    "All": "_all",
    "Old": "_old",
    "Other": "_other"
}

def resolve_position_columns(df, trader, crop):
    cols = set(df.columns)

    if trader == "PMPU":
        if crop == "All":
            long_candidates  = ["prod_merc_positions_long", "prod_merc_positions_long_all"]
            short_candidates = ["prod_merc_positions_short", "prod_merc_positions_short_all"]
            spread_candidates = []
        elif crop == "Old":
            long_candidates  = ["prod_merc_positions_long_old", "prod_merc_positions_long_1"]
            short_candidates = ["prod_merc_positions_short_old", "prod_merc_positions_short_1"]
            spread_candidates = []
        else:  # Other
            long_candidates  = ["prod_merc_positions_long_other", "prod_merc_positions_long_2"]
            short_candidates = ["prod_merc_positions_short_other", "prod_merc_positions_short_2"]
            spread_candidates = []

    elif trader == "Swap Dealer":
        if crop == "All":
            long_candidates  = ["swap_positions_long_all", "swap__positions_long_all", "swap_positions_long"]
            short_candidates = ["swap__positions_short_all", "swap_positions_short_all", "swap_positions_short"]
            spread_candidates = ["swap__positions_spread_all", "swap_positions_spread_all", "swap_positions_spread"]
        elif crop == "Old":
            long_candidates  = ["swap_positions_long_old", "swap_positions_long_1"]
            short_candidates = ["swap__positions_short_old", "swap_positions_short_old", "swap_positions_short_1"]
            spread_candidates = ["swap__positions_spread_old", "swap_positions_spread_old", "swap_positions_spread_1"]
        else:
            long_candidates  = ["swap_positions_long_other", "swap_positions_long_2"]
            short_candidates = ["swap__positions_short_other", "swap_positions_short_other", "swap_positions_short_2"]
            spread_candidates = ["swap__positions_spread_other", "swap_positions_spread_other", "swap_positions_spread_2"]

    elif trader == "Managed Money":
        if crop == "All":
            long_candidates  = ["m_money_positions_long_all", "m_money_positions_long"]
            short_candidates = ["m_money_positions_short_all", "m_money_positions_short"]
            spread_candidates = ["m_money_positions_spread", "m_money_positions_spread_all"]
        elif crop == "Old":
            long_candidates  = ["m_money_positions_long_old", "m_money_positions_long_1"]
            short_candidates = ["m_money_positions_short_old", "m_money_positions_short_1"]
            spread_candidates = ["m_money_positions_spread_old", "m_money_positions_spread_1"]
        else:
            long_candidates  = ["m_money_positions_long_other", "m_money_positions_long_2"]
            short_candidates = ["m_money_positions_short_other", "m_money_positions_short_2"]
            spread_candidates = ["m_money_positions_spread_other", "m_money_positions_spread_2"]

    elif trader == "Other":
        if crop == "All":
            long_candidates  = ["other_rept_positions_long", "other_rept_positions_long_all"]
            short_candidates = ["other_rept_positions_short", "other_rept_positions_short_all"]
            spread_candidates = ["other_rept_positions_spread", "other_rept_positions_spread_all"]
        elif crop == "Old":
            long_candidates  = ["other_rept_positions_long_old", "other_rept_positions_long_1"]
            short_candidates = ["other_rept_positions_short_old", "other_rept_positions_short_1"]
            spread_candidates = ["other_rept_positions_spread_old", "other_rept_positions_spread_1"]
        else:
            long_candidates  = ["other_rept_positions_long_other", "other_rept_positions_long_2"]
            short_candidates = ["other_rept_positions_short_other", "other_rept_positions_short_2"]
            spread_candidates = ["other_rept_positions_spread_other", "other_rept_positions_spread_2"]

    elif trader == "Non-Reportables":
        if crop == "All":
            long_candidates  = ["nonrept_positions_long_all", "nonrept_positions_long"]
            short_candidates = ["nonrept_positions_short_all", "nonrept_positions_short"]
            spread_candidates = []
        elif crop == "Old":
            long_candidates  = ["nonrept_positions_long_old", "nonrept_positions_long_1"]
            short_candidates = ["nonrept_positions_short_old", "nonrept_positions_short_1"]
            spread_candidates = []
        else:
            long_candidates  = ["nonrept_positions_long_other", "nonrept_positions_long_2"]
            short_candidates = ["nonrept_positions_short_other", "nonrept_positions_short_2"]
            spread_candidates = []

    long_col = next((c for c in long_candidates if c in cols), None)
    short_col = next((c for c in short_candidates if c in cols), None)
    spread_col = next((c for c in spread_candidates if c in cols), None)

    return long_col, short_col, spread_col


@st.cache_data
def load_commodity_df(commodity, futs_only=True):
    path = Path("CoT_Disagg_FutsOnly.xlsx") if futs_only else Path("CoT_Disagg_FnO.xlsx")
    sheet = COMMODITY_SHEETS[commodity]

    df = pd.read_excel(path, sheet_name=sheet)
    df["report_date"] = pd.to_datetime(df["report_date_as_yyyy_mm_dd"], errors="coerce")
    return df


def build_position_df(df, long_col, short_col, spread_col):
    out = df.copy()

    out["report_date"] = pd.to_datetime(out["report_date_as_yyyy_mm_dd"], errors="coerce")
    out = out.dropna(subset=["report_date"]).copy()
    out["year"] = out["report_date"].dt.year

    def clean_numeric(series):
        return pd.to_numeric(
            series.astype(str)
                  .str.replace(",", "", regex=False)
                  .str.strip()
                  .replace({"": None, "nan": None, "None": None}),
            errors="coerce"
        )

    out["long"] = clean_numeric(out[long_col]) if long_col in out.columns else pd.NA
    out["short"] = clean_numeric(out[short_col]) if short_col in out.columns else pd.NA

    if spread_col is not None and spread_col in out.columns:
        out["spread"] = clean_numeric(out[spread_col])
    else:
        out["spread"] = pd.NA

    out["net"] = out["long"] - out["short"]
    out["week_of_year"] = ((out["report_date"].dt.dayofyear - 1) // 7) + 1

    return out[["report_date", "year", "week_of_year", "long", "short", "spread", "net"]].sort_values("report_date")


def get_available_years(pos_df):
    years = sorted(pos_df["year"].dropna().unique().tolist())
    return [int(y) for y in years]


def get_recent_5_completed_years(pos_df):
    years = get_available_years(pos_df)
    if not years:
        return []

    current_year = max(years)
    hist_years = [y for y in years if y < current_year]

    return hist_years[-5:]


def build_reference_band(pos_df, value_col):
    ref_years = get_recent_5_completed_years(pos_df)

    if not ref_years:
        return None

    ref = pos_df[pos_df["year"].isin(ref_years)].copy()
    ref = ref.dropna(subset=[value_col])

    if ref.empty:
        return None

    band = (
        ref.groupby("week_of_year")[value_col]
        .agg(ref_min="min", ref_max="max", ref_avg="mean")
        .reset_index()
        .sort_values("week_of_year")
    )

    return band


def filter_years(pos_df, selected_years):
    if not selected_years:
        return pos_df.iloc[0:0].copy()

    return pos_df[pos_df["year"].isin(selected_years)].copy()


def build_net_table(df, crop):
    trader_list = ["PMPU", "Swap Dealer", "Managed Money", "Other", "Non-Reportables"]

    out = None

    for trader in trader_list:

        long_col, short_col, _ = resolve_position_columns(df, trader, crop)

        if long_col is None or short_col is None:
            continue

        tmp = df.copy()

        tmp["date"] = pd.to_datetime(tmp["report_date_as_yyyy_mm_dd"], errors="coerce")

        long_vals = pd.to_numeric(tmp[long_col], errors="coerce")
        short_vals = pd.to_numeric(tmp[short_col], errors="coerce")

        tmp[trader] = long_vals - short_vals

        tmp = tmp[["date", trader]]

        if out is None:
            out = tmp
        else:
            out = out.merge(tmp, on="date", how="outer")

    return out.sort_values("date")




def render_position():
    
    st.header("CFTC Positions")

    # ---- TOP: commodity only ----
    commodity = st.selectbox(
        "Commodity",
        list(COMMODITY_SHEETS.keys()),
        key="position_commodity"
    )

    df = load_commodity_df(commodity)

    st.write("""
    Generally, the data in the COT reports is from Tuesday and released Friday.           

    The COT reports are based on position data supplied by reporting firms (FCMs, clearing members, foreign brokers and exchanges). 
    While the position data is supplied by reporting firms, the actual trader category or classification is based on the predominant business purpose
     self-reported by traders on the CFTC Form 40 and is subject to review by CFTC staff for reasonableness.

    The CFTC will classify each reportable trader as one of the below categories:  
    - **Producer/Merchant/Processor/User (PMPU)**: A “producer/merchant/processor/user” is an entity that predominantly engages in the production, processing,
     packing or handling of a physical commodity and uses the futures markets to manage or hedge risks associated with those activities.
    - **Swap Dealer**: A “swap dealer” is an entity that deals primarily in swaps for a commodity and uses the futures markets to manage or hedge the risk associated with those swaps transactions. The swap
    dealer’s counterparties may be speculative traders, like hedge funds, or traditional commercial clients that are managing risk arising from their dealings in the physical commodity.
    - **Money Manager**: A “money manager” is a registered commodity trading advisor (CTA); a registered commodity pool operator (CPO); or an unregistered fund
    identified by CFTC. These traders are engaged in managing and conducting organized futures trading on behalf of clients.
    - **Other Reportables**: Every other reportable trader that is not placed into one of the other three categories is placed into the “other reportables” category. 
    
    Quick mental model for CoT:  
    Category------------Real economic meaning  
    PMPU:---------------Physical hedgers (producers, merchants, processors)  
    Swap Dealers:-----OTC intermediaries hedging swaps  
    Managed Money:-Hedge funds / CTAs  
    Other:---------------Misc. reportables  
    Non-reportables:-Retail / small         
    """)

    # ---- TABLE ----
    net_table = build_net_table(df, crop="All")   # or keep your preferred crop logic

    st.markdown("### <center>Net Positions by Trader Type</center>", unsafe_allow_html=True)

    display_df = net_table.sort_values("date", ascending=False).copy()
    display_df["date"] = display_df["date"].dt.strftime("%Y-%m-%d")

    num_cols = display_df.columns.drop("date")
    display_df[num_cols] = display_df[num_cols].applymap(
        lambda x: f"{int(x):,}" if pd.notnull(x) else ""
    )

    left, center, right = st.columns([1, 3, 1])
    with center:
        st.dataframe(
            display_df,
            height=200,
            use_container_width=True
        )

    # ---- BELOW TABLE: remaining controls ----
    col1, col2 = st.columns(2)

    trader = col1.selectbox(
        "Trader Type",
        ["PMPU", "Swap Dealer", "Managed Money", "Other", "Non-Reportables"],
        key="position_trader"
    )

    crop = col2.selectbox(
        "Crop Year",
        ["All", "Old", "Other"],
        key="position_crop"
    )
    st.subheader("Position Visualized")


    long_col, short_col, spread_col = resolve_position_columns(df, trader, crop)
    pos_df = build_position_df(df, long_col, short_col, spread_col)

    available_years = get_available_years(pos_df)

    selected_years = st.multiselect(
        "Years",
        available_years,
        default=available_years[-3:] if len(available_years) >= 3 else available_years,
        key="position_years"
    )

    col_a, col_b = st.columns(2)

    show_avg = col_a.checkbox("5Y average", value=False, key="position_show_avg")
    show_band = col_b.checkbox("5Y min/max", value=False, key="position_show_band")

    # ...rest of chart code...
    
    plot_df = filter_years(pos_df, selected_years)


    net_band = build_reference_band(pos_df, "net") if (show_avg or show_band) else None
    long_band = build_reference_band(pos_df, "long") if (show_avg or show_band) else None
    short_band = build_reference_band(pos_df, "short") if (show_avg or show_band) else None
    spread_band = build_reference_band(pos_df, "spread") if (show_avg or show_band) else None

    color_map = get_year_color_map(selected_years)

    long_short_range = get_combined_range(plot_df, ["long", "short"])

    fig_net = plot_seasonal_series(
        plot_df=plot_df,
        selected_years=selected_years,
        value_col="net",
        title="Net Position",
        show_avg=show_avg,
        show_band=show_band,
        ref_band=net_band,
        color_map=color_map,
        show_legend=True
    )

    fig_long = plot_seasonal_series(
        plot_df=plot_df,
        selected_years=selected_years,
        value_col="long",
        title="Long Position",
        show_avg=show_avg,
        show_band=show_band,
        ref_band=long_band,
        color_map=color_map,
        yaxis_range=long_short_range,
        show_legend=False
    )

    fig_short = plot_seasonal_series(
        plot_df=plot_df,
        selected_years=selected_years,
        value_col="short",
        title="Short Position",
        show_avg=show_avg,
        show_band=show_band,
        ref_band=short_band,
        color_map=color_map,
        yaxis_range=long_short_range,
        show_legend=False
    )

    st.plotly_chart(fig_net, use_container_width=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.plotly_chart(fig_long, use_container_width=True)

    with col2:
        st.plotly_chart(fig_short, use_container_width=True)

    with col3:
        if plot_df["spread"].dropna().empty:
            st.info("No spread data for this trader type.")
        else:
            fig_spread = plot_seasonal_series(
                plot_df=plot_df,
                selected_years=selected_years,
                value_col="spread",
                title="Spread Position",
                show_avg=show_avg,
                show_band=show_band,
                ref_band=spread_band,
                color_map=color_map,
                show_legend=False
            )
            st.plotly_chart(fig_spread, use_container_width=True)


def get_year_color_map(selected_years):
    palette = [
        "#1f77b4", "#d62728", "#2ca02c", "#ff7f0e", "#9467bd",
        "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"
    ]
    return {y: palette[i % len(palette)] for i, y in enumerate(sorted(selected_years))}




def plot_seasonal_series(
    plot_df,
    selected_years,
    value_col,
    title,
    show_avg=False,
    show_band=False,
    ref_band=None,
    color_map=None,
    yaxis_range=None,
    show_legend=True
):
    fig = go.Figure()

    if color_map is None:
        color_map = get_year_color_map(selected_years)

    if show_band and ref_band is not None and not ref_band.empty:
        fig.add_trace(go.Scatter(
            x=ref_band["week_of_year"],
            y=ref_band["ref_max"],
            mode="lines",
            line=dict(width=0),
            showlegend=False,
            hoverinfo="skip"
        ))
        fig.add_trace(go.Scatter(
            x=ref_band["week_of_year"],
            y=ref_band["ref_min"],
            mode="lines",
            line=dict(width=0),
            fill="tonexty",
            fillcolor="rgba(160,160,160,0.25)",
            name="5Y Min/Max",
            showlegend=show_legend,
            hoverinfo="skip"
        ))

    if show_avg and ref_band is not None and not ref_band.empty:
        fig.add_trace(go.Scatter(
            x=ref_band["week_of_year"],
            y=ref_band["ref_avg"],
            mode="lines",
            name="5Y Avg",
            showlegend=show_legend,
            line=dict(color="black", dash="dash", width=2)
        ))

    for year in sorted(selected_years):
        tmp = plot_df[plot_df["year"] == year].dropna(subset=[value_col]).copy()
        if tmp.empty:
            continue

        fig.add_trace(go.Scatter(
            x=tmp["week_of_year"],
            y=tmp[value_col],
            mode="lines",
            name=str(year),
            showlegend=show_legend,
            line=dict(color=color_map[year], width=2)
        ))

    fig.update_layout(
        title=title,
        height=260,
        margin=dict(l=5, r=5, t=35, b=10),
        xaxis=dict(
            title="",
            tickmode="array",
            tickvals=[1, 5, 9, 14, 18, 22, 27, 31, 36, 40, 44, 49],
            ticktext=["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            range=[1, 53]
        ),
        yaxis=dict(
            title="Contracts",
            range=yaxis_range
        ),
        legend_title=""
    )

    return fig

def get_combined_range(plot_df, cols):
    vals = []
    for c in cols:
        if c in plot_df.columns:
            vals.extend(pd.to_numeric(plot_df[c], errors="coerce").dropna().tolist())

    if not vals:
        return None

    vmin = min(vals)
    vmax = max(vals)

    if vmin == vmax:
        pad = max(abs(vmin) * 0.05, 1)
    else:
        pad = (vmax - vmin) * 0.08

    return [vmin - pad, vmax + pad]
