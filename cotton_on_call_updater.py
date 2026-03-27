from __future__ import annotations

import re
from datetime import timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup


BASE_URL = "https://www.cftc.gov"
CURRENT_REPORT_URL = "https://www.cftc.gov/MarketReports/CottonOnCall/index.htm"
HISTORICAL_INDEX_URL = (
    "https://www.cftc.gov/MarketReports/CottonOnCall/HistoricalCottonOn-Call/index.htm"
)
REPORT_PATH_PATTERNS = [
    re.compile(
        r"^/MarketReports/CottonOnCall/HistoricalCottonOn-Call/deaoncall\d{6,8}\.html$",
        re.IGNORECASE,
    ),
    re.compile(r"^/dea/cotton/deaoncall\d{6,8}\.htm$", re.IGNORECASE),
]
REPORT_WEEK_RE = re.compile(r"Weekly Report\s+(\d+)", re.IGNORECASE)
AS_OF_DATE_RE = re.compile(r"as of (\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RELEASE_DATE_RE = re.compile(
    r"Release after .*?,\s*([A-Za-z]+ \d{1,2}, \d{4})",
    re.IGNORECASE,
)
CONTRACT_MONTH_RE = re.compile(r"^(?:[A-Za-z]+)\s+\d{4}$")
NUMERIC_TOKEN_RE = re.compile(r"^-?[\d,]+$")
OUTPUT_FILE = "Cotton_OnCall.parquet"

NUMERIC_COLS = [
    "unfixed_call_sales",
    "sales_change_from_previous_week",
    "unfixed_call_purchases",
    "purchases_change_from_previous_week",
    "open_futures_contracts",
    "open_interest_change_from_previous_week",
    "report_week",
    "futures_month_num",
    "futures_year",
]


def _clean_int(value: str | int | float | None) -> int | None:
    if value is None:
        return None
    text = str(value).replace(",", "").strip()
    if not text or text.lower() == "nan":
        return None
    return int(text)


def _normalize_text_lines(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True).replace("\xa0", " ")
    lines = []
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if line:
            lines.append(line)
    return lines


def _normalize_tokens(html: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    return [
        re.sub(r"\s+", " ", token.replace("\xa0", " ")).strip()
        for token in soup.stripped_strings
        if re.sub(r"\s+", " ", token.replace("\xa0", " ")).strip()
    ]


def _extract_report_links(index_html: str) -> list[str]:
    soup = BeautifulSoup(index_html, "html.parser")
    links: list[str] = []
    seen: set[str] = set()

    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        parsed = urlparse(href)
        href_path = parsed.path if parsed.scheme else href
        if not any(pattern.match(href_path) for pattern in REPORT_PATH_PATTERNS):
            continue
        absolute_url = urljoin(BASE_URL, href)
        if absolute_url in seen:
            continue
        seen.add(absolute_url)
        links.append(absolute_url)

    return links


def _parse_report_page(report_url: str, html: str) -> pd.DataFrame:
    lines = _normalize_text_lines(html)
    tokens = _normalize_tokens(html)
    text = "\n".join(lines)

    week_match = REPORT_WEEK_RE.search(text)
    as_of_match = AS_OF_DATE_RE.search(text)
    release_match = RELEASE_DATE_RE.search(text)

    if as_of_match is None:
        raise ValueError(f"Could not find report date in {report_url}")

    report_week = int(week_match.group(1)) if week_match else None
    report_date = pd.to_datetime(as_of_match.group(1), format="%m/%d/%Y", errors="coerce")
    release_date = (
        pd.to_datetime(release_match.group(1), format="%B %d, %Y", errors="coerce")
        if release_match
        else pd.NaT
    )

    row_start = None
    for idx, token in enumerate(tokens):
        if CONTRACT_MONTH_RE.match(token):
            row_start = idx
            break

    if row_start is None:
        raise ValueError(f"Could not locate table rows in {report_url}")

    rows: list[dict] = []
    i = row_start
    while i < len(tokens):
        contract_month = tokens[i]
        if contract_month != "Totals" and not CONTRACT_MONTH_RE.match(contract_month):
            i += 1
            continue

        if i + 6 >= len(tokens):
            break

        numeric_values = tokens[i + 1 : i + 7]
        if not all(NUMERIC_TOKEN_RE.match(value) for value in numeric_values):
            i += 1
            continue

        is_total = contract_month == "Totals"
        futures_contract_date = (
            pd.to_datetime(contract_month, format="%B %Y", errors="coerce")
            if not is_total
            else pd.NaT
        )

        rows.append(
            {
                "report_date": report_date,
                "release_date": release_date,
                "report_week": report_week,
                "commodity": "Cotton",
                "market_name": "COTTON NO. 2 - ICE FUTURES U.S.",
                "exchange_name": "ICE Futures U.S.",
                "contract_month": contract_month,
                "futures_contract_date": futures_contract_date,
                "futures_month_name": None if is_total else futures_contract_date.strftime("%B"),
                "futures_month_num": None if is_total else futures_contract_date.month,
                "futures_year": None if is_total else futures_contract_date.year,
                "is_total": is_total,
                "unfixed_call_sales": _clean_int(numeric_values[0]),
                "sales_change_from_previous_week": _clean_int(numeric_values[1]),
                "unfixed_call_purchases": _clean_int(numeric_values[2]),
                "purchases_change_from_previous_week": _clean_int(numeric_values[3]),
                "open_futures_contracts": _clean_int(numeric_values[4]),
                "open_interest_change_from_previous_week": _clean_int(numeric_values[5]),
                "report_url": report_url,
                "scraped_at_utc": pd.Timestamp.now(tz=timezone.utc),
            }
        )
        i += 7

        if is_total:
            break

    if not rows:
        raise ValueError(f"Could not parse any table rows from {report_url}")

    return pd.DataFrame(rows)


def _coerce_output_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    for date_col in ["report_date", "release_date", "futures_contract_date", "scraped_at_utc"]:
        if date_col in out.columns:
            out[date_col] = pd.to_datetime(out[date_col], errors="coerce")

    for col in NUMERIC_COLS:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").astype("Int64")

    if "is_total" in out.columns:
        out["is_total"] = out["is_total"].fillna(False).astype(bool)

    return out


def _append_and_dedupe(existing: pd.DataFrame, new_rows: pd.DataFrame) -> pd.DataFrame:
    if existing.empty:
        combined = new_rows.copy()
    elif new_rows.empty:
        combined = existing.copy()
    else:
        combined = pd.concat([existing, new_rows], ignore_index=True)

    if combined.empty:
        return combined

    combined = _coerce_output_dtypes(combined)
    combined = combined.drop_duplicates(
        subset=["report_date", "contract_month"],
        keep="last",
    )
    combined = combined.sort_values(
        ["report_date", "is_total", "futures_contract_date", "contract_month"],
        na_position="last",
    ).reset_index(drop=True)
    return combined


def _read_existing_output(output_path: Path) -> pd.DataFrame:
    if not output_path.exists():
        return pd.DataFrame()
    return pd.read_parquet(output_path)


def _extract_url_release_date(report_url: str) -> pd.Timestamp | pd.NaT:
    match = re.search(r"deaoncall(\d{6,8})\.(?:htm|html)$", report_url, re.IGNORECASE)
    if not match:
        return pd.NaT

    stamp = match.group(1)
    if len(stamp) == 8:
        return pd.to_datetime(stamp, format="%m%d%Y", errors="coerce")

    return pd.to_datetime(stamp, format="%m%d%y", errors="coerce")


def build_cotton_on_call_parquet(
    data_dir: str | Path = ".",
    output_name: str = OUTPUT_FILE,
    force: bool = False,
    timeout_seconds: int = 60,
    verify: bool | str = True,
    trust_env: bool = True,
    recent_lookback_days: int = 21,
) -> dict:
    data_dir = Path(data_dir)
    output_path = data_dir / output_name

    if verify is False:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    session = requests.Session()
    session.trust_env = trust_env
    session.headers.update({"User-Agent": "cftc-cot-dashboard/1.0"})

    index_response = session.get(HISTORICAL_INDEX_URL, timeout=timeout_seconds, verify=verify)
    index_response.raise_for_status()
    report_urls = _extract_report_links(index_response.text)

    existing = pd.DataFrame() if force else _read_existing_output(output_path)
    existing_canonical = _append_and_dedupe(pd.DataFrame(), existing)
    existing_urls = set(existing["report_url"].dropna().astype(str)) if "report_url" in existing.columns else set()

    if force or existing.empty or "report_date" not in existing.columns:
        target_urls = report_urls if force else [url for url in report_urls if url not in existing_urls]
    else:
        latest_local_report_date = pd.to_datetime(existing["report_date"], errors="coerce").max()
        if pd.isna(latest_local_report_date):
            target_urls = [url for url in report_urls if url not in existing_urls]
        else:
            cutoff = latest_local_report_date.normalize() - pd.Timedelta(days=recent_lookback_days)
            recent_urls = []
            for url in report_urls:
                release_dt = _extract_url_release_date(url)
                if pd.notna(release_dt) and release_dt >= cutoff:
                    recent_urls.append(url)
            target_urls = [url for url in recent_urls if force or url not in existing_urls]

    # The "current report" page is a rolling URL whose contents change weekly,
    # so we must refresh it even if that URL already exists in the parquet.
    target_urls = [CURRENT_REPORT_URL] + target_urls
    target_urls = list(dict.fromkeys(target_urls))

    parsed_frames: list[pd.DataFrame] = []
    errors: list[str] = []

    for report_url in target_urls:
        try:
            response = session.get(report_url, timeout=timeout_seconds, verify=verify)
            response.raise_for_status()
            parsed_frames.append(_parse_report_page(report_url, response.text))
        except Exception as exc:
            errors.append(f"{report_url}: {exc}")

    new_rows = pd.concat(parsed_frames, ignore_index=True) if parsed_frames else pd.DataFrame()
    final_df = _append_and_dedupe(pd.DataFrame() if force else existing, new_rows)

    if force and not new_rows.empty:
        final_df = _append_and_dedupe(pd.DataFrame(), new_rows)

    temp_output_path = output_path.with_suffix(output_path.suffix + ".tmp")
    final_df.to_parquet(temp_output_path, index=False)
    temp_output_path.replace(output_path)

    return {
        "output_path": str(output_path),
        "report_count_found": len(report_urls),
        "report_count_fetched": len(target_urls),
        "row_count_written": len(final_df),
        "did_update": force or not final_df.equals(existing_canonical),
        "latest_report_date": (
            final_df["report_date"].max().date().isoformat() if not final_df.empty else None
        ),
        "errors": errors,
    }


if __name__ == "__main__":
    result = build_cotton_on_call_parquet()
    print(f"Wrote {result['row_count_written']:,} rows to {result['output_path']}")
    print(f"Reports found: {result['report_count_found']:,}")
    print(f"Reports fetched: {result['report_count_fetched']:,}")
    print(f"Latest report date: {result['latest_report_date']}")
    if result["errors"]:
        print("Errors:")
        for error in result["errors"]:
            print(f" - {error}")
