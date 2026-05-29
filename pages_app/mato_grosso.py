from __future__ import annotations

import json
import math
import os
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import urllib3
from urllib3.exceptions import InsecureRequestWarning


BASE_DIR = Path(__file__).resolve().parents[1]
PLANTING_PROGRESS_PATH = BASE_DIR / "IMEA" / "planting_prog_mt.json"
HARVEST_PROGRESS_PATH = BASE_DIR / "IMEA" / "harvest_prog_mt.json"
HARVEST_PROGRESS_FALLBACK_PATH = BASE_DIR / "IMEA" / "harvest_prog_mt.txt"
FARMER_SELLING_PATH = BASE_DIR / "IMEA" / "farmer_selling_mt.json"
IMEA_API_BASE_URL = "https://api-imeadigital.imea.com.br/api"
IMEA_SERIES = {
    "Planting": {
        "indicator_id": "705576963633053696",
        "json_path": PLANTING_PROGRESS_PATH,
        "excel_path": BASE_DIR / "IMEA" / "planting_prog_mt.xlsx",
    },
    "Harvest": {
        "indicator_id": "703492383711166464",
        "json_path": HARVEST_PROGRESS_PATH,
        "excel_path": BASE_DIR / "IMEA" / "harvest_prog_mt.xlsx",
    },
    "Farmer selling": {
        "indicator_id": "703126874901708800",
        "json_path": FARMER_SELLING_PATH,
        "excel_path": BASE_DIR / "IMEA" / "farmer_selling_mt.xlsx",
    },
}
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


def _read_json_records(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig") as fh:
        data = json.load(fh)
    return data if isinstance(data, list) else []


def _write_json_records(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(records, fh, ensure_ascii=False, indent=2)


def _login_imea() -> str:
    email, password = _get_imea_credentials()
    if not email or not password:
        raise RuntimeError("IMEA credentials are missing from Streamlit secrets.")

    response = _imea_request(
        "post",
        f"{IMEA_API_BASE_URL}/Account/login",
        headers={"accept": "application/json", "Content-Type": "application/json-patch+json"},
        json={"email": email, "password": password},
        timeout=30,
    )
    response.raise_for_status()
    payload = response.json()
    token = payload.get("access_token")
    if not token:
        raise RuntimeError("IMEA login succeeded but did not return an access_token.")
    return token


def _imea_get_json(path: str, token: str) -> object:
    response = _imea_request(
        "get",
        f"{IMEA_API_BASE_URL}{path}",
        headers={"accept": "application/json", "Authorization": f"Bearer {token}"},
        timeout=45,
    )
    response.raise_for_status()
    return response.json()


def _imea_request(method: str, url: str, **kwargs) -> requests.Response:
    try:
        return requests.request(method, url, **kwargs)
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(InsecureRequestWarning)
        return requests.request(method, url, verify=False, **kwargs)


def _fetch_available_dates(indicator_id: str, token: str) -> list[str]:
    payload = _imea_get_json(f"/SerieHistorica/publica/datas/{indicator_id}", token)
    dates = payload.get("datas", []) if isinstance(payload, dict) else payload
    return sorted(str(date)[:10] for date in dates)


def _fetch_series_date(indicator_id: str, date: str, token: str) -> list[dict]:
    payload = _imea_get_json(f"/SerieHistorica/publica/{indicator_id}/{date}", token)
    if isinstance(payload, list):
        return [record for record in payload if isinstance(record, dict)]
    if isinstance(payload, dict):
        return [payload]
    return []


def _record_date(record: dict) -> str | None:
    value = record.get("data")
    return str(value)[:10] if value else None


def _merge_records(existing: list[dict], additions: list[dict]) -> list[dict]:
    merged = {str(record.get("Id") or f"{record.get('indicadorFinalId')}-{record.get('data')}-{idx}"): record for idx, record in enumerate(existing)}
    for idx, record in enumerate(additions, start=len(merged)):
        key = str(record.get("Id") or f"{record.get('indicadorFinalId')}-{record.get('data')}-{idx}")
        merged[key] = record
    return sorted(
        merged.values(),
        key=lambda record: (
            _record_date(record) or "",
            str(record.get("safraDescricao") or ""),
            str(record.get("tipoLocalidadeDescricao") or ""),
            str(record.get("regiaoNome") or ""),
        ),
    )


def _write_progress_excel(records: list[dict], excel_path: Path) -> None:
    if not records:
        return

    raw = pd.json_normalize(records)
    raw["data"] = pd.to_datetime(raw["data"], errors="coerce")
    raw["valor"] = pd.to_numeric(raw["valor"], errors="coerce")
    raw["variacao"] = pd.to_numeric(raw["variacao"], errors="coerce")
    raw = raw.sort_values(["data", "safraDescricao", "tipoLocalidadeDescricao", "regiaoNome"], na_position="first")

    clean = pd.DataFrame(
        {
            "date": raw["data"],
            "safra": raw["safraDescricao"],
            "indicator_id": raw["indicadorFinalId"],
            "indicator": raw["indicadorFinalNome"],
            "chain": raw["cadeiaNome"],
            "state": raw["estadoNome"],
            "state_code": raw["estadoSigla"],
            "location_type": raw["tipoLocalidadeDescricao"],
            "region": raw["regiaoNome"],
            "region_code": raw["regiaoSigla"],
            "value_pct": raw["valor"],
            "variation": raw["variacao"],
            "unit": raw["unidadeSigla"],
        }
    )
    clean["location"] = clean["region"].fillna(clean["state"])
    clean = clean[
        [
            "date",
            "safra",
            "indicator_id",
            "indicator",
            "chain",
            "state",
            "state_code",
            "location_type",
            "location",
            "region",
            "region_code",
            "value_pct",
            "variation",
            "unit",
        ]
    ]

    pivot_src = clean.copy()
    pivot_src["location"] = pivot_src["location_type"].fillna("") + " - " + pivot_src["location"].fillna("")
    pivot = pivot_src.pivot_table(
        index=["date", "safra"],
        columns="location",
        values="value_pct",
        aggfunc="first",
    ).reset_index()
    pivot.columns.name = None

    seasonal_statewide = pd.DataFrame()
    statewide = clean[clean["location_type"] == "Estado"].sort_values(["safra", "date"]).copy()
    if not statewide.empty:
        statewide["season_index"] = statewide.groupby("safra").cumcount() + 1
        seasonal_parts = [pd.DataFrame({"season_index": sorted(statewide["season_index"].unique())})]
        for safra, safra_df in statewide.groupby("safra", sort=True):
            safra_df = safra_df[["season_index", "date", "value_pct"]].copy()
            safra_df = safra_df.rename(
                columns={
                    "date": f"{safra}_date",
                    "value_pct": f"{safra}_value_pct",
                }
            )
            seasonal_parts.append(safra_df)
        seasonal_statewide = seasonal_parts[0]
        for part in seasonal_parts[1:]:
            seasonal_statewide = seasonal_statewide.merge(part, on="season_index", how="left")

    sheets = {"Clean": clean, "Pivot_by_Location": pivot}
    if not seasonal_statewide.empty:
        sheets["Seasonal_Statewide"] = seasonal_statewide
    sheets["Raw_API"] = raw

    try:
        writer = pd.ExcelWriter(excel_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd")
    except ImportError:
        try:
            writer = pd.ExcelWriter(excel_path, engine="openpyxl")
        except ImportError:
            _write_simple_xlsx(excel_path, sheets)
            return

    with writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def _xlsx_col_name(n: int) -> str:
    value = ""
    while n:
        n, remainder = divmod(n - 1, 26)
        value = chr(65 + remainder) + value
    return value


def _xlsx_cell(row: int, col: int, value: object) -> str:
    ref = f"{_xlsx_col_name(col)}{row}"
    if value is None or (isinstance(value, float) and math.isnan(value)) or pd.isna(value):
        return f'<c r="{ref}"/>'
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    if isinstance(value, pd.Timestamp):
        value = value.strftime("%Y-%m-%d")
    return f'<c r="{ref}" t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def _xlsx_sheet(df: pd.DataFrame) -> str:
    export_df = df.copy()
    for col in export_df.columns:
        if pd.api.types.is_datetime64_any_dtype(export_df[col]):
            export_df[col] = export_df[col].dt.strftime("%Y-%m-%d")

    rows = [
        '<row r="1">'
        + "".join(_xlsx_cell(1, idx + 1, col) for idx, col in enumerate(export_df.columns))
        + "</row>"
    ]
    for row_idx, row in enumerate(export_df.itertuples(index=False, name=None), start=2):
        rows.append(
            f'<row r="{row_idx}">'
            + "".join(_xlsx_cell(row_idx, col_idx + 1, value) for col_idx, value in enumerate(row))
            + "</row>"
        )
    last_cell = f"{_xlsx_col_name(max(len(export_df.columns), 1))}{max(len(export_df) + 1, 1)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="A1:{last_cell}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
        'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        "<sheetData>"
        + "".join(rows)
        + "</sheetData></worksheet>"
    )


def _write_simple_xlsx(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    sheet_items = list(sheets.items())
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(1, len(sheet_items) + 1)
        )
        + "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
        + "".join(
            f'<sheet name="{escape(name)}" sheetId="{idx}" r:id="rId{idx}"/>'
            for idx, (name, _) in enumerate(sheet_items, start=1)
        )
        + "</sheets></workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
            for idx in range(1, len(sheet_items) + 1)
        )
        + "</Relationships>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        for idx, (_, df) in enumerate(sheet_items, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", _xlsx_sheet(df))


def refresh_imea_crop_progress() -> dict:
    token = _login_imea()
    result = {"updated": False, "messages": []}

    for stage, config in IMEA_SERIES.items():
        json_path = Path(config["json_path"])
        excel_path = Path(config["excel_path"])
        indicator_id = str(config["indicator_id"])

        existing = _read_json_records(json_path)
        existing_dates = {_record_date(record) for record in existing}
        existing_dates.discard(None)
        available_dates = _fetch_available_dates(indicator_id, token)
        missing_dates = [date for date in available_dates if date not in existing_dates]

        additions = []
        for date in missing_dates:
            additions.extend(_fetch_series_date(indicator_id, date, token))

        if additions:
            updated_records = _merge_records(existing, additions)
            _write_json_records(json_path, updated_records)
            _write_progress_excel(updated_records, excel_path)
            result["updated"] = True
            result["messages"].append(f"{stage}: added {len(missing_dates)} dates / {len(additions)} rows.")
        else:
            if existing:
                _write_progress_excel(existing, excel_path)
            result["messages"].append(f"{stage}: already up to date.")

    if hasattr(load_progress_file, "clear"):
        load_progress_file.clear()
    if hasattr(load_farmer_selling, "clear"):
        load_farmer_selling.clear()
    return result


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


@st.cache_data
def load_farmer_selling(path_str: str, mtime_ns: int | None) -> pd.DataFrame:
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
    df["location"] = df["regiaoNome"].fillna(df["estadoNome"])
    df["safra_start_year"] = df["safraDescricao"].str[:2].astype(int) + 2000
    return df.sort_values(["safra_start_year", "location", "data"])


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


def build_farmer_selling_chart(df: pd.DataFrame, location: str) -> go.Figure:
    plot_df = df[df["location"] == location].copy()
    if plot_df.empty:
        return go.Figure()

    current_start_year = int(plot_df["safra_start_year"].max())
    plot_df = plot_df[
        (plot_df["data"].dt.year >= plot_df["safra_start_year"])
        & (plot_df["data"].dt.year <= plot_df["safra_start_year"] + 1)
    ].copy()
    current = plot_df[plot_df["safra_start_year"] == current_start_year].sort_values("data").copy()
    prior = plot_df[plot_df["safra_start_year"] < current_start_year].copy()
    complete_prior_years = (
        prior.groupby("safra_start_year")["data"]
        .agg(
            first_year=lambda values: int(values.min().year),
            last_year=lambda values: int(values.max().year),
            count="count",
        )
        .reset_index()
    )
    comparable_years = complete_prior_years.loc[
        (complete_prior_years["first_year"] == complete_prior_years["safra_start_year"])
        & (complete_prior_years["last_year"] >= complete_prior_years["safra_start_year"] + 1)
        & (complete_prior_years["count"] >= 12),
        "safra_start_year",
    ].tolist()
    comparable_years = comparable_years[-3:]
    prior = prior[prior["safra_start_year"].isin(comparable_years)].copy()
    if current.empty:
        return go.Figure()

    axis_start_year = 2000

    def selling_axis_date(row: pd.Series) -> pd.Timestamp:
        year_offset = int(row["data"].year) - int(row["safra_start_year"])
        year_offset = max(0, min(1, year_offset))
        return pd.Timestamp(year=axis_start_year + year_offset, month=int(row["data"].month), day=1)

    current["season_date"] = current.apply(selling_axis_date, axis=1)
    current = current.sort_values("data").groupby(["safraDescricao", "safra_start_year", "season_date"], as_index=False).tail(1)
    prior = prior.sort_values(["safra_start_year", "data"]).copy()
    prior["season_date"] = prior.apply(selling_axis_date, axis=1)
    prior = prior.sort_values("data").groupby(["safraDescricao", "safra_start_year", "season_date"], as_index=False).tail(1)
    avg = (
        prior.groupby("season_date", as_index=False)["valor"].mean().rename(columns={"valor": "three_year_avg"})
        if not prior.empty
        else pd.DataFrame(columns=["season_date", "three_year_avg"])
    )

    fig = go.Figure()
    if not avg.empty:
        prior_count = int(prior["safraDescricao"].nunique())
        avg_label = f"Last {prior_count}-safra avg"
        fig.add_trace(
            go.Scatter(
                x=avg["season_date"],
                y=avg["three_year_avg"],
                mode="lines+markers",
                name=avg_label,
                line=dict(color="#9ca3af", width=3, dash="dash"),
                marker=dict(size=7),
                hovertemplate=f"{avg_label}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
            )
        )

    prior_colors = ["#64748b", "#475569", "#334155"]
    prior_years = sorted(prior["safra_start_year"].dropna().unique().tolist())
    for idx, year in enumerate(prior_years):
        year_df = prior[prior["safra_start_year"] == year].sort_values("season_date")
        if year_df.empty:
            continue
        safra_label = str(year_df["safraDescricao"].iloc[-1])
        fig.add_trace(
            go.Scatter(
                x=year_df["season_date"],
                y=year_df["valor"],
                mode="lines+markers",
                name=safra_label,
                line=dict(color=prior_colors[idx % len(prior_colors)], width=2),
                marker=dict(size=6),
                hovertemplate=f"{safra_label}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
            )
        )

    current_safra = str(current["safraDescricao"].iloc[-1])
    fig.add_trace(
        go.Scatter(
            x=current["season_date"],
            y=current["valor"],
            mode="lines+markers",
            name=f"Current ({current_safra})",
            line=dict(color="#1d4ed8", width=4),
            marker=dict(size=7),
            hovertemplate=f"{current_safra}<br>%{{x|%b %d}}<br>%{{y:.1f}}%<extra></extra>",
        )
    )

    tick_vals = [pd.Timestamp(year=axis_start_year + offset, month=month, day=1) for offset in [0, 1] for month in range(1, 13)]
    tick_text = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"] * 2
    x_start = pd.Timestamp(year=axis_start_year, month=1, day=1)
    x_end = pd.Timestamp(year=axis_start_year + 1, month=12, day=31)
    fig.update_layout(
        title=f"Mato Grosso Farmer Selling - {location}",
        height=420,
        margin=dict(l=8, r=8, t=56, b=8),
        paper_bgcolor="#f4f2ed",
        plot_bgcolor="#f8f7f4",
        legend_title="",
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="left", x=0),
        xaxis=dict(
            title="",
            tickmode="array",
            tickvals=tick_vals,
            ticktext=tick_text,
            range=[x_start, x_end],
            showgrid=True,
            gridcolor="#d9d9d9",
            zeroline=False,
        ),
        yaxis=dict(
            title="% Sold",
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
    if not email or not password:
        st.info(
            "Add IMEA credentials to `.streamlit/secrets.toml` before enabling automatic data updates."
        )

    st.subheader("Crop Progress")

    if st.button("Check IMEA updates", key="mato_grosso_update_button"):
        with st.spinner("Checking IMEA crop progress updates..."):
            try:
                update_result = refresh_imea_crop_progress()
                if update_result["updated"]:
                    st.success("IMEA crop progress data updated.")
                else:
                    st.info("IMEA crop progress data is already up to date.")
                for message in update_result["messages"]:
                    st.caption(message)
            except Exception as exc:
                st.error(f"IMEA update failed: {exc}")

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

    st.subheader("Farmer Selling")

    selling_mtime_ns = FARMER_SELLING_PATH.stat().st_mtime_ns if FARMER_SELLING_PATH.exists() else None
    selling_df = load_farmer_selling(str(FARMER_SELLING_PATH), selling_mtime_ns)
    if selling_df.empty:
        st.warning("No IMEA farmer selling data is available locally yet.")
        return

    selling_locations = ["Mato Grosso"] + sorted(
        selling_location
        for selling_location in selling_df["location"].dropna().unique().tolist()
        if selling_location != "Mato Grosso"
    )
    selling_index = selling_locations.index(location) if location in selling_locations else 0
    selling_location = st.selectbox(
        "Farmer selling location",
        selling_locations,
        index=selling_index,
        key="mato_grosso_farmer_selling_location",
    )
    st.plotly_chart(build_farmer_selling_chart(selling_df, selling_location), use_container_width=True)

    latest_selling_date = selling_df.loc[selling_df["location"] == selling_location, "data"].max()
    if pd.notna(latest_selling_date):
        st.caption(f"Source: IMEA. Latest {selling_location} farmer selling observation: {latest_selling_date.date()}.")
