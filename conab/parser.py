"""Parse CONAB crop progress workbooks into normalized rows."""

from __future__ import annotations

from datetime import date, datetime
import logging
from pathlib import Path
import re
import unicodedata

import pandas as pd

from .scraper import parse_brazilian_date

LOGGER = logging.getLogger(__name__)

STATE_UFS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG",
    "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO",
}
STATE_NAME_TO_UF = {
    "acre": "AC",
    "alagoas": "AL",
    "amapa": "AP",
    "amazonas": "AM",
    "bahia": "BA",
    "ceara": "CE",
    "distrito federal": "DF",
    "espirito santo": "ES",
    "goias": "GO",
    "maranhao": "MA",
    "mato grosso": "MT",
    "mato grosso do sul": "MS",
    "minas gerais": "MG",
    "para": "PA",
    "paraiba": "PB",
    "parana": "PR",
    "pernambuco": "PE",
    "piaui": "PI",
    "rio de janeiro": "RJ",
    "rio grande do norte": "RN",
    "rio grande do sul": "RS",
    "rondonia": "RO",
    "roraima": "RR",
    "santa catarina": "SC",
    "sao paulo": "SP",
    "sergipe": "SE",
    "tocantins": "TO",
}
NATIONAL_STATE_LABELS = {
    "brasil",
    "brazil",
    "total",
    "total brasil",
    "total brazil",
    "nacional",
    "national",
}


def normalize_text(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().lower()


def _parse_pct(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.replace("%", "").replace(",", ".").strip()
        if not value or value in {"-", "--"}:
            return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if 0 <= number <= 1:
        number *= 100
    return number if 0 <= number <= 100 else None


def _season_from_text(text: str) -> str | None:
    match = re.search(r"(\d{2,4})\s*/\s*(\d{2,4})", text)
    if not match:
        return None
    first, second = match.groups()
    return f"{int(first) % 100:02d}/{int(second) % 100:02d}"


def _season_from_date(bulletin_date: date | None) -> str | None:
    if bulletin_date is None:
        return None
    start = bulletin_date.year if bulletin_date.month >= 12 else bulletin_date.year - 1
    return f"{start % 100:02d}/{(start + 1) % 100:02d}"


def _contains_crop(value: object, crop_terms: tuple[str, ...]) -> bool:
    normalized = normalize_text(value)
    return any(term in normalized for term in crop_terms)


def _state_to_uf(value: object) -> str | None:
    normalized = normalize_text(value)
    if normalized in NATIONAL_STATE_LABELS or re.fullmatch(r"\d+\s+estados?", normalized):
        return "BR"
    if normalized in STATE_NAME_TO_UF:
        return STATE_NAME_TO_UF[normalized]
    state = re.sub(r"[^A-Z]", "", str(value).strip().upper())
    return state if len(state) == 2 and state in STATE_UFS else None


def _date_from_cell(value: object, *, reference_date: date | None = None) -> date | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (date, datetime, pd.Timestamp)):
        return pd.to_datetime(value).date()

    text = str(value).strip()
    short_date = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})", text)
    if short_date and reference_date is not None:
        day, month = (int(part) for part in short_date.groups())
        year = reference_date.year - 1 if month > reference_date.month + 1 else reference_date.year
        try:
            return date(year, month, day)
        except ValueError:
            return None

    parsed_br_date = parse_brazilian_date(text)
    if parsed_br_date is not None:
        return parsed_br_date
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.notna(parsed):
        return parsed.date()
    return None


def _find_header_row(raw: pd.DataFrame) -> int | None:
    for idx in range(min(len(raw), 40)):
        row = [normalize_text(v) for v in raw.iloc[idx].tolist()]
        joined = " ".join(row)
        has_state = any(v in {"uf", "estado"} or "unidade federacao" in v for v in row)
        has_progress = any("plantio" in v or "colheita" in v for v in row)
        has_crop = any("cultura" in v or "produto" in v or "cultivar" in v for v in row)
        if has_state and (has_progress or has_crop or "safra" in joined):
            return idx
    return None


def _column_lookup(columns: list[object]) -> dict[str, str | None]:
    normalized = {str(col): normalize_text(col) for col in columns}

    def pick(*needles: str) -> str | None:
        for original, clean in normalized.items():
            if any(needle in clean for needle in needles):
                return original
        return None

    return {
        "crop": pick("cultura", "produto"),
        "state": pick("uf", "estado"),
        "season": pick("safra"),
        "planting": pick("plantio", "semeadura"),
        "harvest": pick("colheita"),
    }


def _parse_tabular_sheet(raw: pd.DataFrame, *, bulletin_date: date | None, crop_terms: tuple[str, ...]) -> pd.DataFrame:
    header_row = _find_header_row(raw)
    if header_row is None:
        return pd.DataFrame()

    frame = raw.iloc[header_row + 1:].copy()
    frame.columns = [str(c).strip() for c in raw.iloc[header_row].tolist()]
    frame = frame.dropna(how="all")
    lookup = _column_lookup(list(frame.columns))
    if lookup["state"] is None or (lookup["planting"] is None and lookup["harvest"] is None):
        return pd.DataFrame()

    default_season = _season_from_text(" ".join(raw.astype(str).head(12).stack().tolist()))
    default_season = default_season or _season_from_date(bulletin_date)

    rows: list[dict[str, object]] = []
    for _, row in frame.iterrows():
        crop = row.get(lookup["crop"]) if lookup["crop"] else "Algodao"
        if lookup["crop"] and not _contains_crop(crop, crop_terms):
            continue
        state = _state_to_uf(row.get(lookup["state"], ""))
        if state is None:
            continue
        season = _season_from_text(str(row.get(lookup["season"], ""))) if lookup["season"] else None
        season = season or default_season
        if season is None:
            continue
        rows.append(
            {
                "bulletin_date": bulletin_date,
                "crop": str(crop).strip() or "Algodao",
                "state": state,
                "season": season,
                "planting_pct": _parse_pct(row.get(lookup["planting"])) if lookup["planting"] else None,
                "harvest_pct": _parse_pct(row.get(lookup["harvest"])) if lookup["harvest"] else None,
            }
        )
    return pd.DataFrame(rows)


def _parse_matrix_sheet(raw: pd.DataFrame, *, bulletin_date: date | None, crop_terms: tuple[str, ...], sheet_name: str) -> pd.DataFrame:
    sheet_text = " ".join(raw.astype(str).head(15).stack().tolist())
    default_season = _season_from_text(sheet_text) or _season_from_date(bulletin_date)
    if default_season is None:
        return pd.DataFrame()

    operation = None
    clean_sheet = normalize_text(sheet_name + " " + sheet_text)
    if "plantio" in clean_sheet or "semeadura" in clean_sheet:
        operation = "planting_pct"
    elif "colheita" in clean_sheet:
        operation = "harvest_pct"

    rows: list[dict[str, object]] = []
    for row_idx in range(len(raw)):
        row_values = raw.iloc[row_idx].tolist()
        if not any(_contains_crop(value, crop_terms) for value in row_values):
            continue
        crop_name = next(str(value).strip() for value in row_values if _contains_crop(value, crop_terms))

        for next_idx in range(row_idx, min(row_idx + 12, len(raw))):
            values = raw.iloc[next_idx].tolist()
            state_positions = [
                (col_idx, str(value).strip().upper())
                for col_idx, value in enumerate(values)
                if str(value).strip().upper() in STATE_UFS
            ]
            for col_idx, state in state_positions:
                pct = None
                for offset in range(1, 5):
                    if col_idx + offset < len(values):
                        pct = _parse_pct(values[col_idx + offset])
                        if pct is not None:
                            break
                if pct is None:
                    continue
                rows.append(
                    {
                        "bulletin_date": bulletin_date,
                        "crop": crop_name,
                        "state": state,
                        "season": default_season,
                        "planting_pct": pct if operation == "planting_pct" else None,
                        "harvest_pct": pct if operation == "harvest_pct" else None,
                    }
                )
    return pd.DataFrame(rows)


def _parse_progress_blocks(
    raw: pd.DataFrame,
    *,
    bulletin_date: date | None,
    crop_terms: tuple[str, ...],
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    row_count = len(raw)
    for row_idx in range(row_count):
        row_values = raw.iloc[row_idx].tolist()
        crop_cells = [value for value in row_values if _contains_crop(value, crop_terms)]
        if not crop_cells:
            continue

        crop_name = str(crop_cells[0]).strip()
        crop_season = _season_from_text(crop_name)
        operation = None
        data_start = None
        date_row = None
        state_col = None
        for scan_idx in range(row_idx + 1, min(row_idx + 8, row_count)):
            scan_values = raw.iloc[scan_idx].tolist()
            scan_text = " ".join(normalize_text(value) for value in scan_values)
            if "plantio" in scan_text or "semeadura" in scan_text:
                operation = "planting_pct"
            elif "colheita" in scan_text:
                operation = "harvest_pct"
            state_cols = [col_idx for col_idx, value in enumerate(scan_values) if normalize_text(value) == "estado"]
            if state_cols:
                state_col = state_cols[0]
                date_row = scan_idx + 2 if scan_idx + 2 < row_count else None
                data_start = scan_idx + 3 if scan_idx + 3 < row_count else None
                break
        if operation is None or date_row is None or data_start is None or state_col is None:
            continue

        dates_by_col = {
            col_idx: parsed_date
            for col_idx, value in enumerate(raw.iloc[date_row].tolist())
            if (parsed_date := _date_from_cell(value, reference_date=bulletin_date)) is not None
        }
        if not dates_by_col and bulletin_date is not None:
            dates_by_col = {col_idx: bulletin_date for col_idx in range(1, raw.shape[1])}

        for data_idx in range(data_start, row_count):
            state_cell = raw.iat[data_idx, state_col] if state_col < raw.shape[1] else None
            if pd.isna(state_cell) or not str(state_cell).strip():
                break
            if any(_contains_crop(value, crop_terms) for value in raw.iloc[data_idx].tolist()):
                break
            state = _state_to_uf(state_cell)
            if state is None:
                continue
            for col_idx, obs_date in dates_by_col.items():
                if col_idx >= raw.shape[1]:
                    continue
                pct = _parse_pct(raw.iat[data_idx, col_idx])
                if pct is None:
                    continue
                season = crop_season
                if obs_date is not None:
                    season = _season_from_date(obs_date)
                season = season or crop_season or _season_from_date(bulletin_date)
                if season is None:
                    continue
                rows.append(
                    {
                        "bulletin_date": obs_date or bulletin_date,
                        "crop": crop_name,
                        "state": state,
                        "season": season,
                        "planting_pct": pct if operation == "planting_pct" else None,
                        "harvest_pct": pct if operation == "harvest_pct" else None,
                    }
                )
    return pd.DataFrame(rows)


def parse_workbook(
    path: str | Path,
    *,
    bulletin_date: date | None = None,
    crop_terms: tuple[str, ...] = ("algodao", "cotton"),
) -> pd.DataFrame:
    """Parse a downloaded CONAB workbook into normalized long format."""

    path = Path(path)
    if bulletin_date is None:
        bulletin_date = parse_brazilian_date(path.name)

    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, engine="openpyxl")
    except Exception as exc:
        LOGGER.warning("Could not read workbook %s: %s", path, exc)
        return pd.DataFrame()

    parsed: list[pd.DataFrame] = []
    normalized_terms = tuple(normalize_text(term) for term in crop_terms)
    for sheet_name, raw in sheets.items():
        try:
            parsed.extend(
                [
                    _parse_tabular_sheet(raw, bulletin_date=bulletin_date, crop_terms=normalized_terms),
                    _parse_matrix_sheet(raw, bulletin_date=bulletin_date, crop_terms=normalized_terms, sheet_name=sheet_name),
                    _parse_progress_blocks(raw, bulletin_date=bulletin_date, crop_terms=normalized_terms),
                ]
            )
        except Exception as exc:
            LOGGER.warning("Skipping sheet %s in %s: %s", sheet_name, path, exc)

    if not parsed:
        return pd.DataFrame()

    df = pd.concat(parsed, ignore_index=True)
    if df.empty:
        return df

    df["bulletin_date"] = pd.to_datetime(df["bulletin_date"]).dt.date
    df["crop"] = df["crop"].fillna("Algodao").astype(str).str.strip()
    df = df.dropna(subset=["bulletin_date", "state", "season"])
    df = df.drop_duplicates(["bulletin_date", "crop", "state", "season"], keep="last")
    return df[["bulletin_date", "crop", "state", "season", "planting_pct", "harvest_pct"]]
