"""Plotly chart construction for CONAB cotton crop progress."""

from __future__ import annotations

from datetime import date

import pandas as pd
import plotly.graph_objects as go


def _season_start(season: str) -> date:
    first = int(str(season).split("/")[0])
    year = 2000 + first if first < 70 else 1900 + first
    return date(year, 12, 1)


def _season_sort_key(season: str) -> int:
    try:
        return int(str(season).split("/")[0])
    except (TypeError, ValueError):
        return -1


def _location_label(state: str) -> str:
    return "National" if state == "BR" else state


def _with_week_of_season(df: pd.DataFrame) -> pd.DataFrame:
    data = df.copy()
    data["bulletin_date"] = pd.to_datetime(data["bulletin_date"])
    data["season_start"] = data["season"].map(lambda season: pd.Timestamp(_season_start(season)))
    data["week_of_season"] = ((data["bulletin_date"] - data["season_start"]).dt.days // 7) + 1
    return data[(data["week_of_season"] >= 1) & (data["week_of_season"] <= 53)]


def cotton_rows(df: pd.DataFrame) -> pd.DataFrame:
    crop = df["crop"].fillna("").astype(str).str.lower()
    return df[crop.str.contains("algod|cotton", regex=True)].copy()


def available_seasons(df: pd.DataFrame, state: str | None = None) -> list[str]:
    data = cotton_rows(df)
    if state:
        data = data[data["state"] == state]
    return sorted(data["season"].dropna().unique().tolist(), key=_season_sort_key)


def _stage_curve(df: pd.DataFrame, pct_col: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["week_of_season", pct_col, "bulletin_date"])
    observations = (
        df.dropna(subset=[pct_col])
        .sort_values(["week_of_season", "bulletin_date"])
    )
    if observations.empty:
        return pd.DataFrame(columns=["week_of_season", pct_col, "bulletin_date"])
    max_rows = observations.groupby("week_of_season")[pct_col].idxmax()
    curve = observations.loc[max_rows, ["week_of_season", pct_col, "bulletin_date"]].sort_values("week_of_season")
    curve[pct_col] = curve[pct_col].cummax()
    return curve


def _stage_average(df: pd.DataFrame, seasons: list[str], pct_col: str) -> pd.DataFrame:
    if not seasons:
        return pd.DataFrame(columns=["week_of_season", pct_col])
    season_curves = (
        df[df["season"].isin(seasons)]
        .dropna(subset=[pct_col])
        .groupby(["season", "week_of_season"], as_index=False)[pct_col]
        .max()
        .sort_values(["season", "week_of_season"])
    )
    season_curves[pct_col] = season_curves.groupby("season")[pct_col].cummax()
    return season_curves.groupby("week_of_season", as_index=False)[pct_col].mean().sort_values("week_of_season")


def _trim_after_complete(curve: pd.DataFrame, y_col: str, threshold: float = 99.5) -> pd.DataFrame:
    if curve.empty:
        return curve
    trimmed = curve.dropna(subset=[y_col]).sort_values("week_of_season").copy()
    complete_rows = trimmed.index[trimmed[y_col] >= threshold].tolist()
    if complete_rows:
        first_complete_pos = trimmed.index.get_loc(complete_rows[0])
        trimmed = trimmed.iloc[: first_complete_pos + 1]
    return trimmed


def _broken_line_values(curve: pd.DataFrame, y_col: str, max_gap_weeks: int = 6) -> tuple[list, list]:
    if curve.empty:
        return [], []
    ordered = curve.dropna(subset=["week_of_season", y_col]).sort_values("week_of_season")
    x_vals: list = []
    y_vals: list = []
    previous_week = None
    for row in ordered.itertuples(index=False):
        week = getattr(row, "week_of_season")
        value = getattr(row, y_col)
        if previous_week is not None and week - previous_week > max_gap_weeks:
            x_vals.append(None)
            y_vals.append(None)
        x_vals.append(week)
        y_vals.append(value)
        previous_week = week
    return x_vals, y_vals


def _add_end_label(
    fig: go.Figure,
    curve: pd.DataFrame,
    y_col: str,
    text: str,
    color: str,
    *,
    yshift: int = 0,
) -> None:
    if curve.empty:
        return
    last_row = curve.dropna(subset=[y_col]).sort_values("week_of_season").tail(1)
    if last_row.empty:
        return
    fig.add_annotation(
        x=float(last_row["week_of_season"].iloc[0]),
        y=float(last_row[y_col].iloc[0]),
        text=text,
        showarrow=False,
        xanchor="left",
        yanchor="middle",
        xshift=8,
        yshift=yshift,
        font=dict(color=color, size=12),
    )


def build_progress_chart(df: pd.DataFrame, state: str, current_season: str) -> go.Figure:
    """Build a USDA-style CONAB cotton progress chart."""

    fig = go.Figure()
    data = cotton_rows(df)
    if data.empty:
        return fig.update_layout(title=f"Cotton Crop Progress - {_location_label(state)}")

    data = _with_week_of_season(data[data["state"] == state])
    if data.empty:
        return fig.update_layout(title=f"Cotton Crop Progress - {_location_label(state)}")

    prior_seasons = [
        season
        for season in available_seasons(data, state)
        if _season_sort_key(season) < _season_sort_key(current_season)
    ]
    avg_seasons = prior_seasons[-5:]
    current_df = data[data["season"] == current_season].copy()

    stage_config = [
        {
            "label": "Planting",
            "pct_col": "planting_pct",
            "current_color": "#b5655a",
            "avg_color": "#b5655a",
        },
        {
            "label": "Harvest",
            "pct_col": "harvest_pct",
            "current_color": "#3A8F62",
            "avg_color": "#3A8F62",
        },
    ]

    latest_current_week = None
    for config in stage_config:
        label = config["label"]
        pct_col = config["pct_col"]
        current_color = config["current_color"]
        avg_color = config["avg_color"]

        avg_curve = _trim_after_complete(_stage_average(data, avg_seasons, pct_col), pct_col)
        if not avg_curve.empty:
            avg_x, avg_y = _broken_line_values(avg_curve, pct_col)
            fig.add_trace(
                go.Scatter(
                    x=avg_x,
                    y=avg_y,
                    mode="lines",
                    name=f"{label} 5Y Avg",
                    line=dict(color=avg_color, dash="dot", width=1.6),
                    opacity=0.55,
                    connectgaps=False,
                    hovertemplate=f"{label} 5Y avg<br>Week %{{x}}<br>%{{y:.1f}}%<extra></extra>",
                    showlegend=False,
                )
            )
            _add_end_label(fig, avg_curve, pct_col, f"{label} 5Y Avg", avg_color, yshift=11)

        current_curve = _trim_after_complete(_stage_curve(current_df, pct_col), pct_col)
        if current_curve.empty:
            continue
        current_x, current_y = _broken_line_values(current_curve, pct_col)
        latest_current_week = max(
            latest_current_week or 0,
            int(current_curve["week_of_season"].max()),
        )
        fig.add_trace(
            go.Scatter(
                x=current_x,
                y=current_y,
                mode="lines",
                name=label,
                line=dict(color=current_color, width=3),
                connectgaps=False,
                customdata=pd.to_datetime(current_curve["bulletin_date"]).dt.strftime("%Y-%m-%d"),
                hovertemplate=f"{label}<br>Week %{{x}}<br>%{{y:.1f}}%<br>%{{customdata}}<extra></extra>",
            )
        )
        _add_end_label(fig, current_curve, pct_col, label, current_color, yshift=-11)

    if latest_current_week is not None:
        fig.add_vline(x=latest_current_week, line_dash="dash", line_color="black", line_width=2)

    month_starts = pd.date_range("2020-12-01", periods=12, freq="MS")
    tick_vals = (((month_starts - month_starts[0]).days // 7) + 1).tolist()
    tick_text = [month.strftime("%b") for month in month_starts]

    fig.update_layout(
        title=f"Cotton Crop Progress - {_location_label(state)}",
        height=390,
        margin=dict(l=8, r=86, t=44, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        showlegend=False,
        xaxis=dict(
            title="",
            tickmode="array",
            tickvals=tick_vals,
            ticktext=tick_text,
            range=[1, 53],
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
        ),
        yaxis=dict(
            title="% Progress",
            range=[0, 105],
            tickmode="array",
            tickvals=[0, 20, 40, 60, 80, 100],
            ticksuffix="%",
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
        ),
        hovermode="x unified",
    )
    return fig
