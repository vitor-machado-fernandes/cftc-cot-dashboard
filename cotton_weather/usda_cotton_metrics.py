from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path

import pandas as pd
import requests

from cotton_weather.config import (
    USDA_COTTON_METRICS_FILE,
    USDA_COTTON_METRICS_METADATA_FILE,
    USDA_QUICKSTATS_URL,
)


STATE_NAME_TO_ABBR = {
    "ALABAMA": "AL",
    "ARIZONA": "AZ",
    "ARKANSAS": "AR",
    "CALIFORNIA": "CA",
    "FLORIDA": "FL",
    "GEORGIA": "GA",
    "KANSAS": "KS",
    "LOUISIANA": "LA",
    "MISSOURI": "MO",
    "MISSISSIPPI": "MS",
    "NORTH CAROLINA": "NC",
    "NEW MEXICO": "NM",
    "OKLAHOMA": "OK",
    "SOUTH CAROLINA": "SC",
    "TENNESSEE": "TN",
    "TEXAS": "TX",
    "VIRGINIA": "VA",
}

STATE_ABBR_TO_NAME = {abbr: name.title() for name, abbr in STATE_NAME_TO_ABBR.items()}

METRIC_CONFIGS = {
    "yield_lb_ac": {
        "statisticcat_desc": "YIELD",
        "unit_desc": "LB / ACRE",
        "label": "Yield",
    },
}


def _parse_numeric_value(value: object) -> float | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text in {"(D)", "(NA)", "(Z)"}:
        return None
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def _normalize_state_fields(row: pd.Series) -> tuple[str | None, str | None, bool]:
    agg_level = str(row.get("agg_level_desc", "")).strip().upper()
    state_alpha = str(row.get("state_alpha", "")).strip().upper()
    state_name = str(row.get("state_name", "")).strip().upper()
    location_desc = str(row.get("location_desc", "")).strip().upper()

    if agg_level == "NATIONAL" or state_name in {"US", "US TOTAL", "UNITED STATES"} or location_desc == "US TOTAL":
        return "US", "National", True

    if state_alpha in STATE_ABBR_TO_NAME:
        return state_alpha, STATE_ABBR_TO_NAME[state_alpha], False

    if state_name in STATE_NAME_TO_ABBR:
        state_abbr = STATE_NAME_TO_ABBR[state_name]
        return state_abbr, STATE_ABBR_TO_NAME[state_abbr], False

    return None, None, False


def _fetch_quickstats_rows(
    api_key: str,
    metric_key: str,
    year_start: int,
    year_end: int,
    agg_level: str,
    timeout: int = 180,
) -> list[dict]:
    metric_config = METRIC_CONFIGS[metric_key]
    params = {
        "key": api_key,
        "format": "JSON",
        "source_desc": "SURVEY",
        "sector_desc": "CROPS",
        "group_desc": "FIELD CROPS",
        "commodity_desc": "COTTON",
        "class_desc": "ALL CLASSES",
        "util_practice_desc": "ALL UTILIZATION PRACTICES",
        "prodn_practice_desc": "ALL PRODUCTION PRACTICES",
        "domain_desc": "TOTAL",
        "freq_desc": "ANNUAL",
        "reference_period_desc": "YEAR",
        "year__GE": year_start,
        "year__LE": year_end,
        "agg_level_desc": agg_level,
        "statisticcat_desc": metric_config["statisticcat_desc"],
        "unit_desc": metric_config["unit_desc"],
    }
    response = requests.get(USDA_QUICKSTATS_URL, params=params, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    return payload.get("data", [])


def build_usda_cotton_metrics(api_key: str, years: int = 30) -> pd.DataFrame:
    current_year = datetime.now().year
    year_start = current_year - years + 1
    all_rows: list[dict] = []

    for metric_key in METRIC_CONFIGS:
        for agg_level in ["STATE", "NATIONAL"]:
            raw_rows = _fetch_quickstats_rows(
                api_key=api_key,
                metric_key=metric_key,
                year_start=year_start,
                year_end=current_year,
                agg_level=agg_level,
            )
            for raw_row in raw_rows:
                state, state_display, is_national = _normalize_state_fields(pd.Series(raw_row))
                if state is None:
                    continue
                value = _parse_numeric_value(raw_row.get("Value"))
                year = pd.to_numeric(raw_row.get("year"), errors="coerce")
                if value is None or pd.isna(year):
                    continue
                all_rows.append(
                    {
                        "metric": metric_key,
                        "metric_label": METRIC_CONFIGS[metric_key]["label"],
                        "year": int(year),
                        "state": state,
                        "state_display": state_display,
                        "is_national": is_national,
                        "value": float(value),
                        "short_desc": raw_row.get("short_desc"),
                        "unit_desc": raw_row.get("unit_desc"),
                        "source_desc": raw_row.get("source_desc"),
                    }
                )

    acreage_rows: list[dict] = []
    for agg_level in ["STATE", "NATIONAL"]:
        for statisticcat_desc in ["AREA PLANTED", "AREA HARVESTED"]:
            response_rows = _fetch_quickstats_rows_from_params(
                api_key=api_key,
                year_start=year_start,
                year_end=current_year,
                agg_level=agg_level,
                statisticcat_desc=statisticcat_desc,
                unit_desc="ACRES",
            )
            for raw_row in response_rows:
                state, state_display, is_national = _normalize_state_fields(pd.Series(raw_row))
                if state is None:
                    continue
                value = _parse_numeric_value(raw_row.get("Value"))
                year = pd.to_numeric(raw_row.get("year"), errors="coerce")
                if value is None or pd.isna(year):
                    continue
                acreage_rows.append(
                    {
                        "year": int(year),
                        "state": state,
                        "state_display": state_display,
                        "is_national": is_national,
                        "statisticcat_desc": statisticcat_desc,
                        "value": float(value),
                    }
                )

    acreage_df = pd.DataFrame(acreage_rows)
    if not acreage_df.empty:
        acreage_pivot = (
            acreage_df.pivot_table(
                index=["year", "state", "state_display", "is_national"],
                columns="statisticcat_desc",
                values="value",
                aggfunc="last",
            )
            .reset_index()
        )
        planted = pd.to_numeric(acreage_pivot.get("AREA PLANTED"), errors="coerce")
        harvested = pd.to_numeric(acreage_pivot.get("AREA HARVESTED"), errors="coerce")
        valid = planted.notna() & harvested.notna() & (planted > 0)
        abandonment_df = acreage_pivot.loc[valid, ["year", "state", "state_display", "is_national"]].copy()
        abandonment_df["metric"] = "abandonment_pct"
        abandonment_df["metric_label"] = "Abandonment"
        abandonment_df["value"] = ((planted.loc[valid] - harvested.loc[valid]) / planted.loc[valid]) * 100.0
        abandonment_df["short_desc"] = "Computed from AREA PLANTED and AREA HARVESTED"
        abandonment_df["unit_desc"] = "PCT OF PLANTED"
        abandonment_df["source_desc"] = "SURVEY"
        all_rows.extend(abandonment_df.to_dict(orient="records"))

    output = pd.DataFrame(all_rows)
    if output.empty:
        return output

    output = (
        output.sort_values(["metric", "is_national", "state", "year"])
        .drop_duplicates(subset=["metric", "state", "year"], keep="last")
        .reset_index(drop=True)
    )
    USDA_COTTON_METRICS_FILE.parent.mkdir(parents=True, exist_ok=True)
    output.to_parquet(USDA_COTTON_METRICS_FILE, index=False)

    years_by_metric = {
        metric_key: sorted(output.loc[output["metric"] == metric_key, "year"].unique().tolist())
        for metric_key in sorted(output["metric"].unique())
    }
    metadata = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source": "USDA NASS Quick Stats",
        "source_url": USDA_QUICKSTATS_URL,
        "years_requested": years,
        "year_start": year_start,
        "year_end": current_year,
        "row_count": int(len(output)),
        "metrics": sorted(output["metric"].unique().tolist()),
        "states": sorted(output.loc[~output["is_national"], "state"].unique().tolist()),
        "years_by_metric": years_by_metric,
    }
    USDA_COTTON_METRICS_METADATA_FILE.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    return output


def load_usda_cotton_metrics() -> pd.DataFrame:
    if not USDA_COTTON_METRICS_FILE.exists():
        return pd.DataFrame()
    output = pd.read_parquet(USDA_COTTON_METRICS_FILE)
    if output.empty:
        return output
    output["year"] = pd.to_numeric(output["year"], errors="coerce").astype("Int64")
    output["value"] = pd.to_numeric(output["value"], errors="coerce")
    output["is_national"] = output["is_national"].fillna(False).astype(bool)
    output = output.dropna(subset=["year", "value", "state"]).copy()
    output["year"] = output["year"].astype(int)
    return output.sort_values(["metric", "is_national", "state", "year"]).reset_index(drop=True)


def _fetch_quickstats_rows_from_params(
    api_key: str,
    year_start: int,
    year_end: int,
    agg_level: str,
    statisticcat_desc: str,
    unit_desc: str,
    timeout: int = 180,
) -> list[dict]:
    params = {
        "key": api_key,
        "format": "JSON",
        "source_desc": "SURVEY",
        "sector_desc": "CROPS",
        "group_desc": "FIELD CROPS",
        "commodity_desc": "COTTON",
        "class_desc": "ALL CLASSES",
        "util_practice_desc": "ALL UTILIZATION PRACTICES",
        "prodn_practice_desc": "ALL PRODUCTION PRACTICES",
        "domain_desc": "TOTAL",
        "freq_desc": "ANNUAL",
        "reference_period_desc": "YEAR",
        "year__GE": year_start,
        "year__LE": year_end,
        "agg_level_desc": agg_level,
        "statisticcat_desc": statisticcat_desc,
        "unit_desc": unit_desc,
    }
    response = requests.get(USDA_QUICKSTATS_URL, params=params, timeout=timeout)
    response.raise_for_status()
    payload = response.json()
    return payload.get("data", [])


def load_usda_cotton_metrics_metadata() -> dict:
    if not USDA_COTTON_METRICS_METADATA_FILE.exists():
        return {}
    return json.loads(USDA_COTTON_METRICS_METADATA_FILE.read_text(encoding="utf-8"))


def file_signature(path: Path) -> tuple:
    if not path.exists():
        return (str(path), 0, 0)
    stat = path.stat()
    return (str(path), int(stat.st_size), int(stat.st_mtime))
