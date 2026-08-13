"""Parse CONAB crop progress workbooks into normalized rows."""

from __future__ import annotations

from datetime import date
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
        state = str(row.get(lookup["state"], "")).strip().upper()
        state = re.sub(r"[^A-Z]", "", state)[:2]
        if state not in STATE_UFS:
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
