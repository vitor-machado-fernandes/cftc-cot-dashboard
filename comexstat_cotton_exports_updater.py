from __future__ import annotations

import json
import math
import time
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from urllib3.exceptions import InsecureRequestWarning


API_URL = "https://api-comexstat.mdic.gov.br/general"
UPDATED_DATE_URL = "https://api-comexstat.mdic.gov.br/general/dates/updated"
WEEKLY_PUBLICATION_URL = "https://balanca.mdic.gov.br/balanca/publicacoes_dados_consolidados/nota.html"
DATA_FILENAME = "brazil_cotton_exports.parquet"
METADATA_FILENAME = "brazil_cotton_exports_metadata.json"
WEEKLY_SNAPSHOT_FILENAME = "brazil_cotton_exports_weekly_snapshots.parquet"
WEEKLY_METADATA_FILENAME = "brazil_cotton_exports_weekly_metadata.json"
DEFAULT_START_PERIOD = "2020-01"
DEFAULT_NCM_CODES = {
    "52010020": "Cotton, not carded or combed, simply ginned",
    "52010090": "Other cotton, not carded or combed",
}


def _current_year_end_period() -> str:
    today = date.today()
    return f"{today.year:04d}-12"


def _request_comexstat(payload: dict, language: str = "en") -> dict:
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    url = f"{API_URL}?language={language}"
    last_error: Exception | None = None
    for attempt in range(4):
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=90)
            if response.status_code == 429 and attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            response.raise_for_status()
            return response.json()
        except requests.exceptions.SSLError as exc:
            last_error = exc
            urllib3.disable_warnings(InsecureRequestWarning)
            response = requests.post(url, headers=headers, json=payload, timeout=90, verify=False)
            if response.status_code == 429 and attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as exc:
            last_error = exc
            if attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            raise
    raise RuntimeError(f"ComexStat request failed after retries: {last_error}")


def _request_text(url: str) -> str:
    try:
        response = requests.get(url, timeout=90)
        response.raise_for_status()
        return response.text
    except requests.exceptions.SSLError:
        urllib3.disable_warnings(InsecureRequestWarning)
        response = requests.get(url, timeout=90, verify=False)
        response.raise_for_status()
        return response.text


def fetch_comexstat_latest_update() -> dict:
    for attempt in range(4):
        try:
            response = requests.get(UPDATED_DATE_URL, timeout=60)
            if response.status_code == 429 and attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            response.raise_for_status()
            payload = response.json()
            return payload.get("data", {}) if isinstance(payload, dict) else {}
        except requests.exceptions.SSLError:
            urllib3.disable_warnings(InsecureRequestWarning)
            response = requests.get(UPDATED_DATE_URL, timeout=60, verify=False)
            if response.status_code == 429 and attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            response.raise_for_status()
            payload = response.json()
            return payload.get("data", {}) if isinstance(payload, dict) else {}
        except requests.exceptions.RequestException:
            if attempt < 3:
                time.sleep(15 * (attempt + 1))
                continue
            raise
    return {}


def _normalize_text(value: object) -> str:
    text = str(value or "")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().split())


def _parse_brazil_number(value: object) -> float | None:
    text = str(value or "").strip()
    if not text:
        return None
    text = text.replace("%", "").replace("\xa0", " ").strip()
    text = text.replace(".", "").replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return None
    return number if math.isfinite(number) else None


def _week_of_month(value: date) -> int:
    return int(math.ceil(value.day / 7))


def fetch_brazil_cotton_exports(
    from_period: str = DEFAULT_START_PERIOD,
    to_period: str | None = None,
    ncm_codes: list[str] | None = None,
) -> pd.DataFrame:
    to_period = to_period or _current_year_end_period()
    ncm_codes = ncm_codes or list(DEFAULT_NCM_CODES)
    payload = {
        "flow": "export",
        "monthDetail": True,
        "period": {"from": from_period, "to": to_period},
        "filters": [{"filter": "ncm", "values": ncm_codes}],
        "details": ["ncm"],
        "metrics": ["metricKG"],
    }

    result = _request_comexstat(payload)
    rows = result.get("data", {}).get("list", [])
    if not rows:
        return pd.DataFrame(
            columns=["date", "ncm_code", "cotton_type", "weight_kg", "weight_tons"]
        )

    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(
        df["year"].astype(str) + "-" + df["monthNumber"].astype(str),
        format="%Y-%m",
        errors="coerce",
    )
    df["weight_kg"] = pd.to_numeric(df["metricKG"], errors="coerce")
    df["weight_tons"] = df["weight_kg"] / 1000
    df = df.rename(columns={"coNcm": "ncm_code", "ncm": "cotton_type"})
    df["ncm_code"] = df["ncm_code"].astype(str)
    df["cotton_type"] = df["cotton_type"].fillna(df["ncm_code"].map(DEFAULT_NCM_CODES))

    keep_cols = ["date", "ncm_code", "cotton_type", "weight_kg", "weight_tons"]
    return df[keep_cols].dropna(subset=["date", "weight_kg"]).sort_values(
        ["date", "ncm_code"]
    )


def _normalized_for_compare(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    compare = df.copy()
    compare["date"] = pd.to_datetime(compare["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    compare["weight_kg"] = pd.to_numeric(compare["weight_kg"], errors="coerce").round(0)
    compare["weight_tons"] = pd.to_numeric(compare["weight_tons"], errors="coerce").round(3)
    return compare.sort_values(["date", "ncm_code"]).reset_index(drop=True)


def _load_local(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df


def load_brazil_cotton_exports(data_dir: str | Path = ".") -> pd.DataFrame:
    return _load_local(Path(data_dir) / DATA_FILENAME)


def _load_weekly_snapshots(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_parquet(path)
    if "snapshot_date" in df.columns:
        df["snapshot_date"] = pd.to_datetime(df["snapshot_date"], errors="coerce")
    if "period_start" in df.columns:
        df["period_start"] = pd.to_datetime(df["period_start"], errors="coerce")
    return df


def load_brazil_cotton_weekly_snapshots(data_dir: str | Path = ".") -> pd.DataFrame:
    return _load_weekly_snapshots(Path(data_dir) / WEEKLY_SNAPSHOT_FILENAME)


def fetch_weekly_raw_cotton_snapshot(publication_url: str = WEEKLY_PUBLICATION_URL) -> dict:
    html = _request_text(publication_url)
    soup = BeautifulSoup(html, "html.parser")
    today = date.today()
    normalized_target = "algodao em bruto"

    candidates: list[dict] = []
    for table_index, table in enumerate(soup.find_all("table")):
        headings = []
        previous = table
        for _ in range(8):
            previous = previous.find_previous(["h2", "h3", "h4", "p"])
            if previous is None:
                break
            text = previous.get_text(" ", strip=True)
            if text:
                headings.append(text)
        heading_text = " | ".join(headings)
        normalized_heading = _normalize_text(heading_text)
        is_weekly_product_table = "semana" in normalized_heading

        rows = table.find_all("tr")
        if not rows:
            continue
        for row in rows:
            cells = [
                cell.get_text(" ", strip=True)
                for cell in row.find_all(["th", "td"])
            ]
            if not cells:
                continue
            normalized_cells = [_normalize_text(cell) for cell in cells]
            if not any(normalized_target in cell for cell in normalized_cells):
                continue
            if not is_weekly_product_table:
                continue
            numeric_values = [
                parsed
                for parsed in (_parse_brazil_number(cell) for cell in cells)
                if parsed is not None
            ]
            if numeric_values:
                candidates.append(
                    {
                        "table_index": table_index,
                        "value_usd_million": numeric_values[0],
                        "heading_text": heading_text,
                        "raw_row": cells,
                    }
                )

    if not candidates:
        raise RuntimeError(
            "MDIC's public weekly table exposes numbered weeks for total trade, but this page does not expose a product-level weekly row for `Algodao em bruto`. "
            "Historical weekly cotton bars cannot be backfilled from ComexStat because its stable API is monthly."
        )

    selected = candidates[0]
    period_start = date(today.year, today.month, 1)
    return {
        "snapshot_date": today,
        "period_start": period_start,
        "period": f"{today.year:04d}-{today.month:02d}",
        "week_of_month": _week_of_month(today),
        "product": "Algodao em bruto",
        "value_usd_million_cumulative": selected["value_usd_million"],
        "source_url": publication_url,
        "source_table_index": selected["table_index"],
        "source_heading": selected["heading_text"],
        "raw_row": json.dumps(selected["raw_row"], ensure_ascii=False),
    }


def refresh_brazil_cotton_weekly_snapshots(
    data_dir: str | Path = ".",
    force: bool = False,
) -> dict:
    data_dir = Path(data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    path = data_dir / WEEKLY_SNAPSHOT_FILENAME
    metadata_path = data_dir / WEEKLY_METADATA_FILENAME
    existing = _load_weekly_snapshots(path)
    snapshot = fetch_weekly_raw_cotton_snapshot()
    snapshot_key = (snapshot["period"], snapshot["week_of_month"])

    updated = existing.copy()
    if not updated.empty:
        same_week = (
            (updated["period"].astype(str) == snapshot_key[0])
            & (pd.to_numeric(updated["week_of_month"], errors="coerce") == snapshot_key[1])
        )
        existing_value = (
            pd.to_numeric(updated.loc[same_week, "value_usd_million_cumulative"], errors="coerce").iloc[-1]
            if same_week.any()
            else None
        )
        should_write = force or not same_week.any() or existing_value != snapshot["value_usd_million_cumulative"]
        if same_week.any():
            updated = updated.loc[~same_week].copy()
    else:
        should_write = True

    if should_write:
        updated = pd.concat([updated, pd.DataFrame([snapshot])], ignore_index=True)
        updated["snapshot_date"] = pd.to_datetime(updated["snapshot_date"], errors="coerce")
        updated["period_start"] = pd.to_datetime(updated["period_start"], errors="coerce")
        updated = updated.sort_values(["period_start", "week_of_month", "snapshot_date"])
        updated.to_parquet(path, index=False)
        metadata = {
            "source": "MDIC preliminary weekly trade publication",
            "source_url": WEEKLY_PUBLICATION_URL,
            "product": "Algodao em bruto",
            "unit": "US$ million, cumulative month-to-date",
            "updated_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "latest_snapshot_date": snapshot["snapshot_date"].isoformat(),
            "rows": int(len(updated)),
            "note": "Historical weekly preliminary files may be overwritten by MDIC; this dataset is built from snapshots captured by the app.",
        }
        metadata_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    return {
        "did_update": bool(should_write),
        "path": str(path),
        "period": snapshot["period"],
        "week_of_month": int(snapshot["week_of_month"]),
        "value_usd_million_cumulative": float(snapshot["value_usd_million_cumulative"]),
        "rows": int(len(updated)) if not updated.empty else 0,
    }


def refresh_brazil_cotton_exports(
    data_dir: str | Path = ".",
    force: bool = False,
    from_period: str = DEFAULT_START_PERIOD,
    to_period: str | None = None,
    ncm_codes: list[str] | None = None,
) -> dict:
    data_dir = Path(data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    path = data_dir / DATA_FILENAME
    metadata_path = data_dir / METADATA_FILENAME
    to_period = to_period or _current_year_end_period()
    ncm_codes = ncm_codes or list(DEFAULT_NCM_CODES)

    local = _load_local(path)
    remote = fetch_brazil_cotton_exports(
        from_period=from_period,
        to_period=to_period,
        ncm_codes=ncm_codes,
    )

    local_compare = _normalized_for_compare(local)
    remote_compare = _normalized_for_compare(remote)
    did_update = force or not local_compare.equals(remote_compare)

    if did_update:
        remote.to_parquet(path, index=False)
        latest_date = remote["date"].max() if not remote.empty else pd.NaT
        official_update = fetch_comexstat_latest_update()
        metadata = {
            "source": "ComexStat / MDIC",
            "api_url": API_URL,
            "updated_date_url": UPDATED_DATE_URL,
            "flow": "export",
            "from_period": from_period,
            "to_period": to_period,
            "ncm_codes": ncm_codes,
            "comexstat_updated_date": official_update.get("updated"),
            "comexstat_updated_year": official_update.get("year"),
            "comexstat_updated_month_number": official_update.get("monthNumber"),
            "updated_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "latest_data_month": (
                pd.Timestamp(latest_date).strftime("%Y-%m") if pd.notna(latest_date) else None
            ),
            "rows": int(len(remote)),
        }
        metadata_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")

    latest_local = local["date"].max() if not local.empty and "date" in local.columns else pd.NaT
    latest_remote = remote["date"].max() if not remote.empty else pd.NaT
    return {
        "did_update": did_update,
        "path": str(path),
        "rows": int(len(remote)),
        "local_latest": pd.Timestamp(latest_local).date() if pd.notna(latest_local) else None,
        "remote_latest": pd.Timestamp(latest_remote).date() if pd.notna(latest_remote) else None,
        "from_period": from_period,
        "to_period": to_period,
    }


if __name__ == "__main__":
    result = refresh_brazil_cotton_exports()
    status = "updated" if result["did_update"] else "already up to date"
    print(
        "Brazilian cotton exports "
        f"{status}: {result['rows']} rows through {result['remote_latest'] or 'N/A'}."
    )
