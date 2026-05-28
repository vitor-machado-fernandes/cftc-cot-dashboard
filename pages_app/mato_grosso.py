from __future__ import annotations

import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


BASE_DIR = Path(__file__).resolve().parents[1]
PLANTING_PROGRESS_PATH = BASE_DIR / "IMEA" / "planting_prog_mt.json"
HARVEST_PROGRESS_PATH = BASE_DIR / "IMEA" / "harvest_prog_mt.json"
HARVEST_PROGRESS_FALLBACK_PATH = BASE_DIR / "IMEA" / "harvest_prog_mt.txt"
STAGE_CONFIG = {
    "Planting": {"current_color": "#2563eb", "avg_color": "#93c5fd"},
    "Harvest": {"current_color": "#16a34a", "avg_color": "#86efac"},
}


def _get_imea_credentials() -> tuple[str | None, str | None]:
    email = None
    password = None

    if hasattr(st, "secrets"):
        imea_secrets = st.secrets.get("imea", {})
        email = imea_secrets.get("email") or st.secrets.get("IMEA_EMAIL")
        password = imea_secrets.get("password") or st.secrets.get("IMEA_PASSWORD")

    return email or os.getenv("IMEA_EMAIL"), password or os.getenv("IMEA_PASSWORD")


@st.cache_data
def load_progress_file(path_str: str, mtime_ns: int | None, stage: str) -> pd.DataFrame:
    del mtime_ns
    path = Path(path_str)
    if not path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_json(path)
    except ValueError:
        return pd.DataFrame()

    if df.empty:
        return df

    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce")
    df = df.dropna(subset=["data", "valor", "safraDescricao"]).copy()
    df["stage"] = stage
    df["location"] = df["regiaoNome"].fillna(df["estadoNome"])
    df["safra_start_year"] = df["safraDescricao"].str[:2].astype(int) + 2000
    return df.sort_values(["stage", "safra_start_year", "location", "data"])


def load_crop_progress() -> tuple[pd.DataFrame, bool]:
    planting_mtime_ns = PLANTING_PROGRESS_PATH.stat().st_mtime_ns if PLANTING_PROGRESS_PATH.exists() else None
    planting_df = load_progress_file(str(PLANTING_PROGRESS_PATH), planting_mtime_ns, "Planting")

    harvest_path = HARVEST_PROGRESS_PATH if HARVEST_PROGRESS_PATH.exists() else HARVEST_PROGRESS_FALLBACK_PATH
    harvest_mtime_ns = harvest_path.stat().st_mtime_ns if harvest_path.exists() else None
    harvest_df = load_progress_file(str(harvest_path), harvest_mtime_ns, "Harvest")
    harvest_file_invalid = harvest_path.exists() and harvest_df.empty

    frames = [df for df in [planting_df, harvest_df] if not df.empty]
    if not frames:
        return pd.DataFrame(), harvest_file_invalid
    return pd.concat(frames, ignore_index=True), harvest_file_invalid


def _current_season_start_year(df: pd.DataFrame) -> int:
    return int(df["safra_start_year"].max())


def _aligned_current_season_date(date_value: pd.Timestamp, current_start_year: int) -> pd.Timestamp:
    aligned_year = current_start_year if date_value.month == 12 else current_start_year + 1
    return pd.Timestamp(year=aligned_year, month=int(date_value.month), day=int(date_value.day))


def _prior_average_curve(
    prior: pd.DataFrame,
    current_start_year: int,
) -> pd.DataFrame:
    if prior.empty:
        return pd.DataFrame(columns=["aligned_date", "three_year_avg"])

    aligned_prior = prior.copy()
    aligned_prior["aligned_date"] = aligned_prior["data"].apply(
        lambda value: _aligned_current_season_date(value, current_start_year)
    )
    current_dates = sorted(aligned_prior["aligned_date"].dropna().unique())
    avg_rows = []
    for aligned_date in current_dates:
        values = []
        for _, season_df in prior.groupby("safraDescricao"):
            season_df = season_df.copy()
            season_df["aligned_date"] = season_df["data"].apply(
                lambda value: _aligned_current_season_date(value, current_start_year)
            )
            season_df = season_df.sort_values("aligned_date")
            x = season_df["aligned_date"].map(pd.Timestamp.toordinal).to_numpy()
            y = season_df["valor"].to_numpy(dtype=float)
            target = pd.Timestamp(aligned_date).toordinal()
            if len(x) == 0 or target < x.min() or target > x.max():
                continue
            point_series = pd.Series(y, index=x).groupby(level=0).mean().sort_index()
            values.append(float(np.interp(target, point_series.index.to_numpy(), point_series.to_numpy())))
        if values:
            avg_rows.append({"aligned_date": pd.Timestamp(aligned_date), "three_year_avg": sum(values) / len(values)})

    return pd.DataFrame(avg_rows)


def build_crop_progress_chart(df: pd.DataFrame, location: str) -> go.Figure:
    plot_df = df[df["location"] == location].copy()
    if plot_df.empty:
        return go.Figure()

    current_start_year = _current_season_start_year(plot_df)
    fig = go.Figure()

    for stage, config in STAGE_CONFIG.items():
        stage_df = plot_df[plot_df["stage"] == stage].copy()
        if stage_df.empty:
            continue

        current = stage_df[stage_df["safra_start_year"] == current_start_year].copy()
        prior = stage_df[
            (stage_df["safra_start_year"] < current_start_year)
            & (stage_df["safra_start_year"] >= current_start_year - 3)
        ].copy()

        avg = _prior_average_curve(prior, current_start_year).sort_values("aligned_date")
        if not avg.empty:
            prior_season_count = int(prior["safraDescricao"].nunique())
            avg_label = f"{stage} {prior_season_count}-year avg"
            fig.add_trace(
                go.Scatter(
                    x=avg["aligned_date"],
                    y=avg["three_year_avg"],
                    mode="lines+markers",
                    name=avg_label,
                    line=dict(color=config["avg_color"], width=3, dash="dash"),
                    marker=dict(size=7),
                    hovertemplate=f"{avg_label}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
                )
            )

        latest_prior = pd.DataFrame()
        if avg.empty and current.empty and not prior.empty:
            latest_prior_year = int(prior["safra_start_year"].max())
            latest_prior = prior[prior["safra_start_year"] == latest_prior_year].copy()
            latest_prior["aligned_date"] = latest_prior["data"].apply(
                lambda value: _aligned_current_season_date(value, current_start_year)
            )
            latest_prior = latest_prior.sort_values("aligned_date")
            latest_safra = str(latest_prior["safraDescricao"].iloc[-1])
            fig.add_trace(
                go.Scatter(
                    x=latest_prior["aligned_date"],
                    y=latest_prior["valor"],
                    mode="lines+markers",
                    name=f"{stage} prior ({latest_safra})",
                    line=dict(color=config["avg_color"], width=3, dash="dash"),
                    marker=dict(size=7),
                    hovertemplate=f"{stage} {latest_safra}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
                )
            )

        if current.empty:
            continue

        current["aligned_date"] = current["data"].apply(
            lambda value: _aligned_current_season_date(value, current_start_year)
        )
        current = current.sort_values("aligned_date")
        current_safra = str(current["safraDescricao"].iloc[-1])
        fig.add_trace(
            go.Scatter(
                x=current["aligned_date"],
                y=current["valor"],
                mode="lines+markers",
                name=f"{stage} current ({current_safra})",
                line=dict(color=config["current_color"], width=4),
                marker=dict(size=7),
                hovertemplate=f"{stage} {current_safra}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
            )
        )

    x_start = pd.Timestamp(year=current_start_year, month=12, day=1)
    x_end = pd.Timestamp(year=current_start_year + 1, month=11, day=30)

    fig.update_layout(
        title=f"Cotton Crop Progress - {location}",
        height=420,
        margin=dict(l=8, r=8, t=56, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
        xaxis=dict(
            title="",
            tickformat="%b",
            dtick="M1",
            range=[x_start, x_end],
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
    )
    return fig


def render_mato_grosso():
    st.header("Mato Grosso")

    email, password = _get_imea_credentials()
    if email and password:
        st.caption("IMEA credentials are configured.")
    else:
        st.info(
            "Add IMEA credentials to `.streamlit/secrets.toml` before enabling automatic data updates."
        )

    st.subheader("Crop Progress")

    progress_df, harvest_file_invalid = load_crop_progress()
    if progress_df.empty:
        st.warning("No IMEA crop progress data is available locally yet.")
        return
    if harvest_file_invalid:
        st.warning("The saved harvest file is not valid JSON yet; it looks like an IMEA error response.")

    locations = ["Mato Grosso"] + sorted(
        location for location in progress_df["location"].dropna().unique().tolist() if location != "Mato Grosso"
    )
    location = st.selectbox("Location", locations, index=0, key="mato_grosso_crop_progress_location")

    st.plotly_chart(build_crop_progress_chart(progress_df, location), use_container_width=True)

    latest_date = progress_df.loc[progress_df["location"] == location, "data"].max()
    if pd.notna(latest_date):
        st.caption(f"Source: IMEA. Latest {location} crop progress observation: {latest_date.date()}.")
