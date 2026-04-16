from __future__ import annotations

import os
import tomllib
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from usda_crop_progress_condition_updater import fetch_crop_progress_condition_history


BASE_DIR = Path(__file__).resolve().parents[1]
DATA_PATH = BASE_DIR / "usda_crop_progress_condition.parquet"
SECRETS_PATH = BASE_DIR / ".streamlit" / "secrets.toml"
CONDITION_ORDER = ["VERY POOR", "POOR", "FAIR", "GOOD", "EXCELLENT"]
CONDITION_COLORS = {
    "VERY POOR": "#7f1d1d",
    "POOR": "#dc2626",
    "FAIR": "#f59e0b",
    "GOOD": "#65a30d",
    "EXCELLENT": "#15803d",
}
PROGRESS_COLORS = {
    "PLANTED": "#b5655a",
    "EMERGED": "#64c86b",
    "SQUARING": "#8b3f6d",
    "SETTING BOLLS": "#e28df2",
    "BOLLS OPENING": "#8ad7df",
    "SILKING": "#c084fc",
    "DOUGH": "#f59e0b",
    "DENTED": "#f97316",
    "MATURE": "#6b7280",
    "HEADED": "#7c3aed",
    "BLOOMING": "#ec4899",
    "SETTING PODS": "#a855f7",
    "DROPPING LEAVES": "#64748b",
    "HARVESTED": "#6f79c8",
}
STATE_OPTIONS = [
    "National",
    "Alabama",
    "Alaska",
    "Arizona",
    "Arkansas",
    "California",
    "Colorado",
    "Connecticut",
    "Delaware",
    "Florida",
    "Georgia",
    "Hawaii",
    "Idaho",
    "Illinois",
    "Indiana",
    "Iowa",
    "Kansas",
    "Kentucky",
    "Louisiana",
    "Maine",
    "Massachusetts",
    "Maryland",
    "Michigan",
    "Minnesota",
    "Mississippi",
    "Missouri",
    "Montana",
    "Nebraska",
    "Nevada",
    "New Hampshire",
    "New Jersey",
    "New Mexico",
    "New York",
    "North Carolina",
    "North Dakota",
    "Ohio",
    "Oklahoma",
    "Oregon",
    "Pennsylvania",
    "Rhode Island",
    "South Carolina",
    "South Dakota",
    "Tennessee",
    "Texas",
    "Utah",
    "Vermont",
    "Virginia",
    "Washington",
    "West Virginia",
    "Wisconsin",
    "Wyoming",
]


@st.cache_data
def load_crop_progress_condition_data(path_str: str, mtime_ns: int | None) -> pd.DataFrame:
    path = Path(path_str)
    if not path.exists():
        return pd.DataFrame()

    df = pd.read_parquet(path)
    if df.empty:
        return df

    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["week_of_year"] = pd.to_numeric(df["week_of_year"], errors="coerce")
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    if "agg_level_desc" not in df.columns:
        df["agg_level_desc"] = "NATIONAL"
    if "state_name" not in df.columns:
        df["state_name"] = ""
    if "location_label" not in df.columns:
        df["location_label"] = "National"
    return df.dropna(subset=["report_date"]).copy()


def get_usda_api_key() -> str | None:
    if hasattr(st, "secrets"):
        for key_name in ["USDA_QUICKSTATS_API_KEY", "QUICKSTATS_API_KEY", "NASS_API_KEY"]:
            value = st.secrets.get(key_name)
            if value:
                return value

    for key_name in ["USDA_QUICKSTATS_API_KEY", "QUICKSTATS_API_KEY", "NASS_API_KEY"]:
        value = os.getenv(key_name)
        if value:
            return value

    if SECRETS_PATH.exists():
        with SECRETS_PATH.open("rb") as fh:
            local_secrets = tomllib.load(fh)
        for key_name in ["USDA_QUICKSTATS_API_KEY", "QUICKSTATS_API_KEY", "NASS_API_KEY"]:
            value = local_secrets.get(key_name)
            if value:
                return value

    return None


@st.cache_data(show_spinner=False)
def load_state_crop_progress_condition_data(
    crop: str,
    state_name: str,
    api_key: str | None,
    start_year: int,
) -> pd.DataFrame:
    if not api_key:
        return pd.DataFrame()

    state_candidates = [state_name]
    if state_name.upper() not in state_candidates:
        state_candidates.append(state_name.upper())

    df = pd.DataFrame()
    for state_candidate in state_candidates:
        df = fetch_crop_progress_condition_history(
            api_key=api_key,
            crop=crop,
            agg_level_desc="STATE",
            state_name=state_candidate,
            start_year=start_year,
        )
        if not df.empty:
            break

    if df.empty:
        return df

    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["week_of_year"] = pd.to_numeric(df["week_of_year"], errors="coerce")
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df.dropna(subset=["report_date"]).copy()


def get_year_color_map(years: list[int]) -> dict[int, str]:
    palette = [
        "#1f77b4",
        "#d62728",
        "#2ca02c",
        "#ff7f0e",
        "#9467bd",
        "#8c564b",
        "#e377c2",
        "#7f7f7f",
        "#bcbd22",
        "#17becf",
    ]
    return {year: palette[idx % len(palette)] for idx, year in enumerate(sorted(years))}


def month_tick_config() -> dict:
    return dict(
        tickmode="array",
        tickvals=[1, 5, 9, 14, 18, 22, 27, 31, 36, 40, 44, 49],
        ticktext=["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        range=[1, 53],
    )


def apply_report_layout(fig: go.Figure, height: int, y_title: str, y_range: list[int] | None = None) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=8, t=44, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
            **month_tick_config(),
        ),
        yaxis=dict(
            title=y_title,
            range=y_range,
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
            tickformat=".0f",
        ),
    )
    return fig


def latest_year_dates(df: pd.DataFrame) -> tuple[int | None, list[pd.Timestamp]]:
    if df.empty:
        return None, []

    latest_year = int(df["year"].dropna().max())
    dates = sorted(df.loc[df["year"] == latest_year, "report_date"].dropna().unique())
    return latest_year, [pd.Timestamp(d) for d in dates]


def build_broken_line_arrays(
    series_df: pd.DataFrame,
    x_col: str = "week_of_year",
    y_col: str = "value",
    date_col: str | None = "report_date",
    gap_days: int = 35,
    gap_weeks: int = 4,
) -> tuple[list, list, list]:
    if series_df.empty:
        return [], [], []

    sort_cols = [date_col] if date_col else [x_col]
    ordered = series_df.sort_values(sort_cols).copy()
    x_vals: list = []
    y_vals: list = []
    custom_vals: list = []
    prev_date = None
    prev_x = None

    for _, row in ordered.iterrows():
        current_x = row[x_col]
        current_date = pd.Timestamp(row[date_col]) if date_col else None

        should_break = False
        if current_date is not None and prev_date is not None and (current_date - prev_date).days > gap_days:
            should_break = True
        if prev_x is not None and pd.notna(current_x) and (current_x - prev_x) > gap_weeks:
            should_break = True

        if should_break:
            x_vals.append(None)
            y_vals.append(None)
            custom_vals.append(None)

        x_vals.append(current_x)
        y_vals.append(row[y_col])
        custom_vals.append(current_date.strftime("%Y-%m-%d") if current_date is not None else None)
        prev_date = current_date
        prev_x = current_x

    return x_vals, y_vals, custom_vals


def infer_current_season_start(progress_df: pd.DataFrame, selected_date: pd.Timestamp) -> pd.Timestamp | None:
    selected_year = int(selected_date.year)
    year_df = progress_df[
        (progress_df["year"] == selected_year)
        & (progress_df["report_date"] <= selected_date)
    ].copy()
    if year_df.empty:
        return None

    active_progress = year_df[~year_df["stage"].isin(["HARVESTED"])]
    if not active_progress.empty:
        return active_progress["report_date"].min()

    return year_df["report_date"].min()


def derive_progress_stage(df: pd.DataFrame) -> pd.Series:
    stage = df["unit_desc"].fillna("").str.upper().str.strip()
    stage = stage.str.replace(r"^PCT\s+", "", regex=True)
    stage = stage.mask(stage.eq(""), df["short_desc"].str.upper().str.extract(r"PCT\s+(.+)$", expand=False))
    return stage.fillna("").str.replace(r"\s+", " ", regex=True).str.strip()


def progress_subset(df: pd.DataFrame, crop: str) -> pd.DataFrame:
    out = df[
        (df["crop"] == crop)
        & (df["statisticcat_desc"] == "PROGRESS")
    ].copy()
    out["stage"] = derive_progress_stage(out)
    out = out[~out["stage"].isin(["", "TOTAL"])]
    return out


def condition_subset(df: pd.DataFrame, crop: str) -> pd.DataFrame:
    out = df[
        (df["crop"] == crop)
        & (df["statisticcat_desc"] == "CONDITION")
    ].copy()
    out["condition"] = (
        out["unit_desc"]
        .fillna("")
        .str.upper()
        .str.strip()
        .str.replace(r"^PCT\s+", "", regex=True)
    )
    out = out[out["condition"].isin(CONDITION_ORDER)]
    return out


def build_progress_lines(df: pd.DataFrame, selected_date: pd.Timestamp) -> go.Figure:
    current_year = int(selected_date.year)
    season_start = infer_current_season_start(df, selected_date)
    plot_df = df[(df["year"] == current_year) & (df["report_date"] <= selected_date)].copy()
    if season_start is not None:
        plot_df = plot_df[plot_df["report_date"] >= season_start]
    plot_df = plot_df.dropna(subset=["value", "week_of_year"])
    hist_df = df[df["year"] < current_year].dropna(subset=["value", "week_of_year"]).copy()

    fig = go.Figure()
    current_stages = set(plot_df["stage"].dropna().unique().tolist())
    hist_stages = set(hist_df["stage"].dropna().unique().tolist())
    stages = sorted(current_stages | hist_stages)
    for stage in stages:
        stage_color = PROGRESS_COLORS.get(stage, "#4b5563")
        hist_stage = hist_df[hist_df["stage"] == stage]
        if not hist_stage.empty:
            hist_avg = (
                hist_stage.groupby("week_of_year", as_index=False)["value"]
                .mean()
                .sort_values("week_of_year")
            )
            hist_x, hist_y, _ = build_broken_line_arrays(hist_avg, date_col=None)
            fig.add_trace(
                go.Scatter(
                    x=hist_x,
                    y=hist_y,
                    mode="lines",
                    name=f"{stage.title()} 5Y Avg",
                    line=dict(width=1.5, dash="dot", color=stage_color),
                    opacity=0.55,
                    connectgaps=False,
                    hovertemplate=f"{stage.title()} 5Y avg<br>Week %{{x}}<br>%{{y:.0f}}%<extra></extra>",
                    showlegend=False,
                )
            )

        stage_df = plot_df[plot_df["stage"] == stage].sort_values("report_date")
        if stage_df.empty:
            continue
        x_vals, y_vals, custom_vals = build_broken_line_arrays(stage_df)
        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_vals,
                mode="lines",
                name=stage.title(),
                line=dict(width=3, color=stage_color),
                customdata=custom_vals,
                connectgaps=False,
                hovertemplate="%{fullData.name}<br>Week %{x}<br>%{y:.0f}%<br>%{customdata}<extra></extra>",
            )
        )
        last_row = stage_df.iloc[-1]
        fig.add_annotation(
            x=float(last_row["week_of_year"]),
            y=float(last_row["value"]),
            text=stage.title(),
            showarrow=False,
            xanchor="left",
            yanchor="middle",
            xshift=8,
            font=dict(color=stage_color, size=12),
        )

    fig.add_vline(x=int(selected_date.isocalendar().week), line_dash="dash", line_color="black")
    fig.update_layout(title=f"Progress ({current_year})")
    apply_report_layout(fig, height=360, y_title="% Progress", y_range=[0, 100])
    fig.update_layout(showlegend=False)
    return fig


def build_planting_and_harvest_chart(df: pd.DataFrame, selected_date: pd.Timestamp) -> go.Figure:
    stage_df = df[df["stage"].str.upper().isin(["PLANTED", "HARVESTED"])].copy()
    stage_df = stage_df.dropna(subset=["value", "week_of_year", "year"])
    if stage_df.empty:
        return go.Figure()

    selected_year = int(selected_date.year)
    all_years = sorted(int(y) for y in stage_df["year"].dropna().unique())
    prior_years = [y for y in all_years if y < selected_year]
    chosen_years = (prior_years[-5:] + [selected_year]) if selected_year in all_years else prior_years[-6:]
    color_map = get_year_color_map(chosen_years)
    dash_map = {"PLANTED": "solid", "HARVESTED": "dash"}

    fig = go.Figure()
    for year in chosen_years:
        year_df = stage_df[stage_df["year"] == year].copy()
        if year == selected_year:
            year_df = year_df[year_df["report_date"] <= selected_date]

        if year_df.empty:
            continue

        for stage_name in ["PLANTED", "HARVESTED"]:
            stage_year_df = year_df[year_df["stage"].str.upper() == stage_name].sort_values("report_date")
            if stage_year_df.empty:
                continue
            x_vals, y_vals, custom_vals = build_broken_line_arrays(stage_year_df)

            fig.add_trace(
                go.Scatter(
                    x=x_vals,
                    y=y_vals,
                    mode="lines",
                    name=f"{year} {stage_name.title()}",
                    line=dict(
                        width=3 if year == selected_year else 2,
                        color=color_map[year],
                        dash=dash_map[stage_name],
                    ),
                    customdata=custom_vals,
                    connectgaps=False,
                    hovertemplate="%{fullData.name}<br>Week %{x}<br>%{y:.0f}%<br>%{customdata}<extra></extra>",
                )
            )

    fig.add_vline(x=int(selected_date.isocalendar().week), line_dash="dash", line_color="black")
    fig.update_layout(title="Planting and Harvest Progress vs Other Years")
    apply_report_layout(fig, height=340, y_title="% of Crop", y_range=[0, 100])
    return fig


def build_condition_stacked_bar(
    df: pd.DataFrame,
    selected_date: pd.Timestamp,
    season_start: pd.Timestamp | None = None,
) -> go.Figure:
    current_year = int(selected_date.year)
    plot_df = df[(df["year"] == current_year) & (df["report_date"] <= selected_date)].copy()
    if season_start is not None:
        plot_df = plot_df[plot_df["report_date"] >= season_start]
    plot_df = plot_df.dropna(subset=["value", "week_of_year"])

    weekly = (
        plot_df.groupby(["week_of_year", "condition"], as_index=False)["value"]
        .sum()
        .pivot(index="week_of_year", columns="condition", values="value")
        .fillna(0.0)
        .reindex(columns=CONDITION_ORDER, fill_value=0.0)
        .sort_index()
    )

    fig = go.Figure()
    for condition in CONDITION_ORDER:
        fig.add_trace(
            go.Bar(
                x=weekly.index.tolist(),
                y=weekly[condition].tolist(),
                name=condition.title(),
                marker_color=CONDITION_COLORS[condition],
                hovertemplate=f"{condition.title()}<br>Week %{{x}}<br>%{{y:.0f}}%<extra></extra>",
            )
        )

    fig.update_layout(
        title=f"Condition Breakdown ({current_year})",
        barmode="stack",
        bargap=0.12,
    )
    apply_report_layout(fig, height=300, y_title="% Condition", y_range=[0, 100])
    return fig


def build_good_excellent_chart(df: pd.DataFrame, selected_date: pd.Timestamp) -> go.Figure:
    plot_df = df[df["condition"].isin(["GOOD", "EXCELLENT"])].copy()
    plot_df = plot_df.dropna(subset=["value", "week_of_year", "year"])
    if plot_df.empty:
        return go.Figure()

    plot_df = (
        plot_df.groupby(["year", "week_of_year", "report_date"], as_index=False)["value"]
        .sum()
        .sort_values(["year", "report_date"])
    )

    selected_year = int(selected_date.year)
    all_years = sorted(int(y) for y in plot_df["year"].dropna().unique())
    prior_years = [y for y in all_years if y < selected_year]
    chosen_years = (prior_years[-5:] + [selected_year]) if selected_year in all_years else prior_years[-6:]
    color_map = get_year_color_map(chosen_years)

    fig = go.Figure()
    for year in chosen_years:
        year_df = plot_df[plot_df["year"] == year].sort_values("report_date")
        if year == selected_year:
            year_df = year_df[year_df["report_date"] <= selected_date]

        if year_df.empty:
            continue

        x_vals, y_vals, custom_vals = build_broken_line_arrays(year_df)

        fig.add_trace(
            go.Scatter(
                x=x_vals,
                y=y_vals,
                mode="lines",
                name=str(year),
                line=dict(width=3 if year == selected_year else 2, color=color_map[year]),
                customdata=custom_vals,
                connectgaps=False,
                hovertemplate="%{fullData.name}<br>Week %{x}<br>%{y:.0f}%<br>%{customdata}<extra></extra>",
            )
        )
        if year == selected_year:
            last_row = year_df.iloc[-1]
            fig.add_annotation(
                x=float(last_row["week_of_year"]),
                y=float(last_row["value"]),
                text=str(year),
                showarrow=False,
                xanchor="left",
                yanchor="middle",
                xshift=8,
                font=dict(color=color_map[year], size=13),
            )

    fig.add_vline(x=int(selected_date.isocalendar().week), line_dash="dash", line_color="black")
    fig.update_layout(title="Good + Excellent vs Other Years")
    apply_report_layout(fig, height=280, y_title="% Good + Excellent", y_range=[0, 100])
    return fig


def render_crop_progress_condition():
    st.header("Crop Progress & Condition")
    st.write(
        """
        Weekly crop progress and crop condition data come from USDA NASS Quick Stats and are cached locally.
        This page focuses on national weekly series for the selected crop.
        """
    )

    file_mtime_ns = DATA_PATH.stat().st_mtime_ns if DATA_PATH.exists() else None
    df = load_crop_progress_condition_data(str(DATA_PATH), file_mtime_ns)
    if df.empty:
        st.warning(
            "No USDA crop progress/condition data is available locally yet. "
            "Set `USDA_QUICKSTATS_API_KEY` in the Streamlit environment and reload the app."
        )
        return

    available_crops = sorted(df["crop"].dropna().unique().tolist())
    crop_index = 0
    if "crop_progress_condition_crop" in st.session_state:
        existing_crop = st.session_state["crop_progress_condition_crop"]
        if existing_crop in available_crops:
            crop_index = available_crops.index(existing_crop)
    state_index = 0
    if "crop_progress_condition_state" in st.session_state:
        existing_state = st.session_state["crop_progress_condition_state"]
        if existing_state in STATE_OPTIONS:
            state_index = STATE_OPTIONS.index(existing_state)

    controls = st.columns(4)
    with controls[0]:
        crop = st.selectbox("Crop", available_crops, index=crop_index, key="crop_progress_condition_crop")
    with controls[1]:
        state_name = st.selectbox("Location", STATE_OPTIONS, index=state_index, key="crop_progress_condition_state")

    if state_name == "National":
        active_df = df[df["location_label"] == "National"].copy()
    else:
        usda_api_key = get_usda_api_key()
        state_start_year = max(pd.Timestamp.today().year - 7, 2018)

        with st.spinner(f"Loading USDA {state_name} crop progress data..."):
            active_df = load_state_crop_progress_condition_data(
                crop,
                state_name,
                usda_api_key,
                state_start_year,
            )

        if active_df.empty:
            st.warning(
                f"No USDA crop progress/condition data was returned for {crop} in {state_name}. "
                "This can happen when the crop is not reported at that state level, the API key is unavailable to this Streamlit session, "
                "or USDA does not return the request cleanly on the first try."
            )
            return

    progress_df = progress_subset(active_df, crop)
    condition_df = condition_subset(active_df, crop)
    progress_year, progress_dates = latest_year_dates(progress_df)
    condition_year, condition_dates = latest_year_dates(condition_df)
    with controls[2]:
        progress_date = None
        if progress_dates:
            progress_date = st.selectbox(
                "Progress report date",
                progress_dates,
                index=len(progress_dates) - 1,
                format_func=lambda d: pd.Timestamp(d).strftime("%Y-%m-%d"),
                key="crop_progress_report_date",
            )
    with controls[3]:
        condition_date = None
        if condition_dates:
            condition_date = st.selectbox(
                "Condition report date",
                condition_dates,
                index=len(condition_dates) - 1,
                format_func=lambda d: pd.Timestamp(d).strftime("%Y-%m-%d"),
                key="crop_condition_report_date",
            )

    if not progress_dates:
        st.info("No progress series found for this crop.")
    if not condition_dates:
        st.info("No condition series found for this crop.")
    if not progress_dates or not condition_dates:
        return

    progress_date = pd.Timestamp(progress_date)
    condition_date = pd.Timestamp(condition_date)
    condition_season_start = infer_current_season_start(progress_df, condition_date)

    st.plotly_chart(build_good_excellent_chart(condition_df, condition_date), use_container_width=True)
    st.plotly_chart(
        build_condition_stacked_bar(condition_df, condition_date, season_start=condition_season_start),
        use_container_width=True,
    )
    st.plotly_chart(build_progress_lines(progress_df, progress_date), use_container_width=True)
    planting_harvest_fig = build_planting_and_harvest_chart(progress_df, progress_date)
    if len(planting_harvest_fig.data) == 0:
        st.info("No planted or harvested progress series is available for the selected crop/location.")
    else:
        st.plotly_chart(planting_harvest_fig, use_container_width=True)
    st.caption(
        f"Using {state_name} {crop} progress data through {progress_date.date()} and "
        f"condition data through {condition_date.date()}."
    )
