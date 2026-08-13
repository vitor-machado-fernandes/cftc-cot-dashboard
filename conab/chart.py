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


def build_progress_chart(df: pd.DataFrame, state: str, current_season: str) -> go.Figure:
    """Build the CONAB cotton progress chart."""

    fig = go.Figure()
    data = cotton_rows(df)
    if data.empty:
        return fig.update_layout(title=f"Cotton Crop Progress - {state}")

    data = _with_week_of_season(data[data["state"] == state])
    if data.empty:
        return fig.update_layout(title=f"Cotton Crop Progress - {state}")

    prior_seasons = [
        season
        for season in available_seasons(data, state)
        if _season_sort_key(season) < _season_sort_key(current_season)
    ]

    planting_avg_seasons = prior_seasons[-3:]
    if planting_avg_seasons:
        planting_avg = (
            data[data["season"].isin(planting_avg_seasons)]
            .dropna(subset=["planting_pct"])
            .groupby("week_of_season", as_index=False)["planting_pct"]
            .mean()
        )
        fig.add_trace(
            go.Scatter(
                x=planting_avg["week_of_season"],
                y=planting_avg["planting_pct"],
                mode="lines",
                name="Planting 3-year avg",
                line=dict(color="#79BDE8", dash="dash", width=2),
            )
        )

    current = data[data["season"] == current_season].dropna(subset=["planting_pct"])
    if not current.empty:
        current = current.sort_values("week_of_season")
        fig.add_trace(
            go.Scatter(
                x=current["week_of_season"],
                y=current["planting_pct"],
                mode="lines",
                name=f"Planting current ({current_season})",
                line=dict(color="#0B4F8A", width=3),
            )
        )

    if prior_seasons:
        harvest_avg = (
            data[data["season"] == prior_seasons[-1]]
            .dropna(subset=["harvest_pct"])
            .sort_values("week_of_season")
        )
        fig.add_trace(
            go.Scatter(
                x=harvest_avg["week_of_season"],
                y=harvest_avg["harvest_pct"],
                mode="lines",
                name="Harvest 1-year avg",
                line=dict(color="#3A8F62", dash="dash", width=2),
            )
        )

    month_starts = pd.date_range("2020-12-01", periods=12, freq="MS")
    tick_vals = (((month_starts - month_starts[0]).days // 7) + 1).tolist()
    tick_text = [month.strftime("%b") for month in month_starts]

    fig.update_layout(
        title=f"Cotton Crop Progress - {state}",
        height=500,
        margin=dict(l=8, r=8, t=80, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
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
            range=[0, 100],
            ticksuffix="%",
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
        ),
        hovermode="x unified",
    )
    return fig
