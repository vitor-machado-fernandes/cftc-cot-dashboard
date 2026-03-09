import pandas as pd
import streamlit as st
from pathlib import Path


COMMODITY_SHEETS = {
    "Corn": "Corn",
    "Cotton": "Cotton",
    "Soybeans": "Soybeans",
    "SBM": "Soybean Meal",
    "SBO": "Soybean Oil"
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
    out["month_day"] = out["report_date"].dt.strftime("%m-%d")

    return out[["report_date", "year", "month_day", "long", "short", "spread", "net"]].sort_values("report_date")


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
        ref.groupby("month_day")[value_col]
        .agg(ref_min="min", ref_max="max", ref_avg="mean")
        .reset_index()
    )

    return band


def filter_years(pos_df, selected_years):
    if not selected_years:
        return pos_df.iloc[0:0].copy()

    return pos_df[pos_df["year"].isin(selected_years)].copy()


def render_position():
    
    st.header("CFTC Positions")

    col1, col2, col3 = st.columns(3)

    commodity = col1.selectbox(
        "Commodity",
        ["Corn", "Cotton", "Soybeans", "SBM", "SBO"],
        key="position_commodity"
    )

    trader = col2.selectbox(
        "Trader Type",
        ["PMPU", "Swap Dealer", "Managed Money", "Other", "Non-Reportables"],
        key="position_trader"
    )

    crop = col3.selectbox(
        "Crop Year",
        ["All", "Old", "Other"],
        key="position_crop"
    )

    df = load_commodity_df(commodity)

    long_col, short_col, spread_col = resolve_position_columns(df, trader, crop)

    pos_df = build_position_df(df, long_col, short_col, spread_col)

    available_years = get_available_years(pos_df)

    selected_years = st.multiselect(
        "Years",
        available_years,
        default=available_years[-3:] if len(available_years) >= 3 else available_years,
        key="position_years")
    
    col_a, col_b = st.columns(2)

    show_avg = col_a.checkbox(
        "5Y average",
        value=False,
        key="position_show_avg")
    
    show_band = col_b.checkbox(
        "5Y min/max",
        value=False,
        key="position_show_band")
    
    plot_df = filter_years(pos_df, selected_years)

    net_band = build_reference_band(pos_df, "net") if (show_avg or show_band) else None
    long_band = build_reference_band(pos_df, "long") if (show_avg or show_band) else None
    short_band = build_reference_band(pos_df, "short") if (show_avg or show_band) else None
    spread_band = build_reference_band(pos_df, "spread") if (show_avg or show_band) else None

    st.write("Selected years:", selected_years)
    st.write("Recent 5 completed years used for reference:", get_recent_5_completed_years(pos_df))

    st.write("Filtered plot df:")
    st.dataframe(plot_df.head(10))


    if net_band is not None:
        st.write("Net reference band preview:")
        st.dataframe(net_band.head(10))


