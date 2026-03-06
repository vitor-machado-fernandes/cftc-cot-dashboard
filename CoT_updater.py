#CoT Updater: this file makes sure the data is always up to date.

from __future__ import annotations

import requests
import pandas as pd
from datetime import timedelta
from pathlib import Path

DATE_COL = "report_date_as_yyyy_mm_dd"
BASE_URL = "https://publicreporting.cftc.gov/resource/{dataset_id}.json"
LIMIT = 50000

# Configure your workbooks and sheet -> market mappings here
FILES = {
    "CoT_Disagg_FutsOnly.xlsx": {
        "dataset_id": "6dca-aqww",
        "sheets": {
            "Corn": "CORN - CHICAGO BOARD OF TRADE",
            "Cotton": "COTTON NO. 2 - ICE FUTURES U.S.",
            "Soybeans": "SOYBEANS - CHICAGO BOARD OF TRADE",
            "SBO": "SOYBEAN OIL - CHICAGO BOARD OF TRADE",
            "SBM": "SOYBEAN MEAL - CHICAGO BOARD OF TRADE",
        },
    },
    "CoT_Disagg_FnO.xlsx": {
        "dataset_id": "kh3c-gbw2",
        "sheets": {
            "Corn": "CORN - CHICAGO BOARD OF TRADE",
            "Cotton": "COTTON NO. 2 - ICE FUTURES U.S.",
            "Soybeans": "SOYBEANS - CHICAGO BOARD OF TRADE",
            "SBO": "SOYBEAN OIL - CHICAGO BOARD OF TRADE",
            "SBM": "SOYBEAN MEAL - CHICAGO BOARD OF TRADE",
        },
    },
}

# Sentinel (choose a stable sheet that always exists)
SENTINEL_FILE = "CoT_Disagg_FutsOnly.xlsx"
SENTINEL_SHEET = "Cotton"


def _latest_date_in_df(df: pd.DataFrame) -> pd.Timestamp | None:
    if df is None or df.empty or DATE_COL not in df.columns:
        return None
    s = pd.to_datetime(df[DATE_COL], errors="coerce")
    if s.isna().all():
        return None
    return s.max().normalize()


def _fetch_latest_cftc_date(market_name: str, dataset_id: str) -> pd.Timestamp | None:
    url = BASE_URL.format(dataset_id=dataset_id)
    params = {
        "market_and_exchange_names": market_name,
        "$select": DATE_COL,
        "$order": f"{DATE_COL} DESC",
        "$limit": 1,
    }
    r = requests.get(url, params=params, timeout=60)
    r.raise_for_status()
    data = r.json()
    if not data:
        return None
    return pd.to_datetime(data[0].get(DATE_COL), errors="coerce").normalize()


def _fetch_new_rows(market_name: str, start_date_yyyy_mm_dd: str, dataset_id: str) -> pd.DataFrame:
    url = BASE_URL.format(dataset_id=dataset_id)
    params = {
        "market_and_exchange_names": market_name,
        "$where": f"{DATE_COL} >= '{start_date_yyyy_mm_dd}'",
        "$order": f"{DATE_COL} ASC",
        "$limit": LIMIT,
    }
    r = requests.get(url, params=params, timeout=60)
    r.raise_for_status()
    return pd.DataFrame(r.json())


def _append_and_sort(df_existing: pd.DataFrame, df_new: pd.DataFrame) -> pd.DataFrame:
    if df_existing is None or df_existing.empty:
        out = df_new.copy()
    else:
        out = pd.concat([df_existing, df_new], ignore_index=True)

    if out.empty or DATE_COL not in out.columns:
        return out

    dt = pd.to_datetime(out[DATE_COL], errors="coerce")
    out = out.loc[~dt.duplicated(keep="first")].copy()
    out["_dt"] = dt
    out = out.sort_values("_dt").drop(columns=["_dt"])
    return out


def _update_workbook(filepath: Path, dataset_id: str, sheets: dict[str, str]) -> list[str]:
    """
    Updates one workbook. Returns a list of human-readable messages of what changed.
    """
    messages: list[str] = []

    book = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")

    updated_any = False
    for sheet_name, market_name in sheets.items():
        df_existing = book.get(sheet_name, pd.DataFrame())
        last_dt = _latest_date_in_df(df_existing)

        if last_dt is None:
            start_date = "2009-01-01"
        else:
            start_date = (last_dt + timedelta(days=1)).strftime("%Y-%m-%d")

        df_new = _fetch_new_rows(market_name, start_date, dataset_id)
        if df_new.empty:
            continue

        book[sheet_name] = _append_and_sort(df_existing, df_new)
        updated_any = True

        new_min = pd.to_datetime(df_new[DATE_COL], errors="coerce").min()
        new_max = pd.to_datetime(df_new[DATE_COL], errors="coerce").max()
        messages.append(
            f"{filepath.name} | {sheet_name}: +{len(df_new)} rows ({new_min.date()} -> {new_max.date()})"
        )

    if not updated_any:
        return messages

    with pd.ExcelWriter(filepath, engine="xlsxwriter") as writer:
        for sname, df in book.items():
            df.to_excel(writer, sheet_name=sname, index=False)

    return messages


def run_update_check(data_dir: str | Path = ".", force: bool = False) -> dict:
    """
    1) Sentinel check (fast)
    2) If new data detected (or force=True), update all files/sheets.

    Returns:
      {
        "did_update": bool,
        "sentinel_local": date|None,
        "sentinel_cftc": date|None,
        "messages": [str, ...]
      }
    """
    data_dir = Path(data_dir)
    messages: list[str] = []

    # --- Sentinel ---
    cfg = FILES[SENTINEL_FILE]
    sentinel_path = data_dir / SENTINEL_FILE
    dataset_id = cfg["dataset_id"]
    market_name = cfg["sheets"][SENTINEL_SHEET]

    df_local = pd.read_excel(
        sentinel_path,
        sheet_name=SENTINEL_SHEET,
        engine="openpyxl",
        usecols=[DATE_COL],
    )
    local_last = _latest_date_in_df(df_local)
    cftc_last = _fetch_latest_cftc_date(market_name, dataset_id)

    sentinel_local = local_last.date() if local_last is not None else None
    sentinel_cftc = cftc_last.date() if cftc_last is not None else None

    if not force and local_last is not None and cftc_last is not None and local_last >= cftc_last:
        return {
            "did_update": False,
            "sentinel_local": sentinel_local,
            "sentinel_cftc": sentinel_cftc,
            "messages": messages,
        }

    # --- Full update ---
    did_update = False
    for fname, fcfg in FILES.items():
        path = data_dir / fname
        msgs = _update_workbook(path, fcfg["dataset_id"], fcfg["sheets"])
        if msgs:
            did_update = True
            messages.extend(msgs)

    return {
        "did_update": did_update,
        "sentinel_local": sentinel_local,
        "sentinel_cftc": sentinel_cftc,
        "messages": messages,
    }