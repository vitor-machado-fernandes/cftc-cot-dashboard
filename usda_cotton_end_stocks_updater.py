from __future__ import annotations

import re
import warnings
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib3.exceptions import InsecureRequestWarning


ARCHIVE_URL = "https://esmis.nal.usda.gov/publication/world-agricultural-supply-and-demand-estimates?page=1"
BASE_URL = "https://esmis.nal.usda.gov"
DATA_FILE = Path("usda_cotton_end_stocks.xlsx")
PRICE_FILE = Path("Cotton Price Hist.xlsx")
UNIT = "Million 480 Pound Bales"
ENDING_STOCKS_SHEET = "EndingStocks"
USE_SHEET = "Use"


@dataclass(frozen=True)
class WasdeRelease:
    release_date: pd.Timestamp
    xml_url: str


def _build_session() -> requests.Session:
    session = requests.Session()
    session.trust_env = False
    return session


def _get(session: requests.Session, url: str) -> requests.Response:
    try:
        response = session.get(url, timeout=60)
    except requests.exceptions.SSLError:
        warnings.filterwarnings("ignore", category=InsecureRequestWarning)
        response = session.get(url, timeout=60, verify=False)

    response.raise_for_status()
    return response


def load_cotton_end_stocks(path: str | Path = DATA_FILE) -> pd.DataFrame:
    workbook_path = Path(path)
    if not workbook_path.exists():
        return pd.DataFrame()

    df = pd.read_excel(workbook_path, sheet_name=ENDING_STOCKS_SHEET)
    for col in ["ReportDate", "ReleaseDate"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def load_cotton_use(path: str | Path = DATA_FILE) -> pd.DataFrame:
    workbook_path = Path(path)
    if not workbook_path.exists():
        return pd.DataFrame()

    try:
        df = pd.read_excel(workbook_path, sheet_name=USE_SHEET)
    except ValueError:
        return pd.DataFrame()

    for col in ["ReportDate", "ReleaseDate"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _write_workbook(
    ending_stocks: pd.DataFrame,
    use: pd.DataFrame,
    path: str | Path = DATA_FILE,
) -> None:
    ending_stocks = add_stocks_use_ratio(ending_stocks, use)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        ending_stocks.to_excel(writer, sheet_name=ENDING_STOCKS_SHEET, index=False)
        use.to_excel(writer, sheet_name=USE_SHEET, index=False)


def add_stocks_use_ratio(
    ending_stocks: pd.DataFrame,
    use: pd.DataFrame,
) -> pd.DataFrame:
    out = ending_stocks.copy()
    if out.empty or use.empty:
        out["Stocks/Use"] = pd.NA
        return out

    out["ReleaseDate"] = pd.to_datetime(out["ReleaseDate"], errors="coerce").dt.normalize()
    use_lookup = use.copy()
    use_lookup["ReleaseDate"] = pd.to_datetime(
        use_lookup["ReleaseDate"],
        errors="coerce",
    ).dt.normalize()
    use_lookup["UseValue"] = pd.to_numeric(use_lookup["Value"], errors="coerce")
    use_lookup = use_lookup.dropna(subset=["ReleaseDate"]).drop_duplicates(
        subset=["ReleaseDate"],
        keep="last",
    )

    out = out.merge(
        use_lookup[["ReleaseDate", "UseValue"]],
        on="ReleaseDate",
        how="left",
    )
    out["Stocks/Use"] = pd.to_numeric(out["Value"], errors="coerce") / out["UseValue"]
    return out.drop(columns=["UseValue"])


def contract_from_market_year(market_year: str) -> str:
    match = re.search(r"(\d{4})/(\d{2,3})", str(market_year))
    if not match:
        raise ValueError(f"Could not parse market year {market_year!r}.")

    start_year = int(match.group(1))
    end_suffix = int(match.group(2)[-2:])
    end_year = (start_year // 100) * 100 + end_suffix
    if end_year <= start_year:
        end_year += 100

    return f"CTN{end_year % 100:02d}"


def _load_price_history(path: str | Path = PRICE_FILE) -> pd.DataFrame:
    price_path = Path(path)
    if not price_path.exists():
        return pd.DataFrame()

    prices = pd.read_excel(price_path)
    if "Date" not in prices.columns:
        return pd.DataFrame()

    prices["Date"] = pd.to_datetime(prices["Date"], errors="coerce").dt.normalize()
    return prices.dropna(subset=["Date"]).set_index("Date")


def _save_price_history(price_history: pd.DataFrame, path: str | Path = PRICE_FILE) -> None:
    out = price_history.copy()
    out.index = pd.to_datetime(out.index, errors="coerce").normalize()
    out = out[~out.index.isna()].copy()
    out = out.sort_index(ascending=False)
    out.insert(0, "Date", out.index)
    out.to_excel(path, index=False)


def _download_yahoo_contract_closes(
    contract: str,
    dates: list[pd.Timestamp],
) -> dict[pd.Timestamp, float]:
    if not dates:
        return {}

    try:
        import yfinance as yf
        from curl_cffi import requests as curl_requests
    except ImportError as exc:
        raise RuntimeError("Install `yfinance` before updating cotton futures prices.") from exc

    cache_dir = Path(".yfinance_cache")
    cache_dir.mkdir(exist_ok=True)
    try:
        yf.set_tz_cache_location(str(cache_dir))
    except Exception:
        pass

    clean_dates = sorted({pd.Timestamp(day).normalize() for day in dates})
    start = (clean_dates[0] - pd.Timedelta(days=3)).date().isoformat()
    end = (clean_dates[-1] + pd.Timedelta(days=2)).date().isoformat()

    session = curl_requests.Session(impersonate="chrome")
    session.verify = False

    ticker = f"{contract}.NYB"
    history = yf.download(
        ticker,
        start=start,
        end=end,
        progress=False,
        auto_adjust=False,
        threads=False,
        session=session,
    )
    if history.empty or "Close" not in history:
        return {}

    close = history["Close"]
    if isinstance(close, pd.DataFrame):
        close = close.iloc[:, 0]
    close.index = pd.to_datetime(close.index, errors="coerce").normalize()

    return {
        day: float(close.loc[day])
        for day in clean_dates
        if day in close.index and pd.notna(close.loc[day])
    }


def update_price_history_for_rows(
    rows: pd.DataFrame,
    price_path: str | Path = PRICE_FILE,
) -> dict:
    if rows.empty or not {"ReleaseDate", "Contract"}.issubset(rows.columns):
        return {"prices_updated": 0, "errors": []}

    price_history = _load_price_history(price_path)
    if price_history.empty:
        price_history = pd.DataFrame()
        price_history.index = pd.DatetimeIndex([], name="Date")

    needed: dict[str, list[pd.Timestamp]] = {}
    for row in rows.itertuples(index=False):
        contract = getattr(row, "Contract", None)
        release_date = getattr(row, "ReleaseDate", None)
        if not contract or pd.isna(release_date):
            continue

        release_day = pd.Timestamp(release_date).normalize()
        has_price = (
            contract in price_history.columns
            and release_day in price_history.index
            and pd.notna(price_history.at[release_day, contract])
        )
        if not has_price:
            needed.setdefault(contract, []).append(release_day)

    prices_updated = 0
    errors = []
    for contract, dates in needed.items():
        try:
            closes = _download_yahoo_contract_closes(contract, dates)
            if contract not in price_history.columns:
                price_history[contract] = pd.NA

            for release_day, close in closes.items():
                if release_day not in price_history.index:
                    price_history.loc[release_day, :] = pd.NA
                if pd.isna(price_history.at[release_day, contract]):
                    prices_updated += 1
                price_history.at[release_day, contract] = close

            missing = sorted(set(dates) - set(closes))
            if missing:
                missing_labels = ", ".join(day.strftime("%Y-%m-%d") for day in missing)
                errors.append(f"{contract}: no Yahoo close found for {missing_labels}")
        except Exception as exc:
            errors.append(f"{contract}: {exc}")

    if prices_updated:
        _save_price_history(price_history, price_path)

    return {"prices_updated": prices_updated, "errors": errors}


def fill_prices_from_history(
    rows: pd.DataFrame,
    price_path: str | Path = PRICE_FILE,
) -> pd.DataFrame:
    out = rows.copy()
    price_history = _load_price_history(price_path)
    if "Contract" not in out.columns:
        out["Contract"] = out["MarketYear"].map(contract_from_market_year)

    out["Price"] = [
        price_for_release(row.ReleaseDate, row.Contract, price_history)
        for row in out.itertuples(index=False)
    ]
    return out


def price_for_release(
    release_date: pd.Timestamp,
    contract: str,
    price_history: pd.DataFrame,
) -> float | None:
    if price_history.empty or contract not in price_history.columns:
        return None

    release_day = pd.Timestamp(release_date).normalize()
    if release_day not in price_history.index:
        return None

    value = price_history.at[release_day, contract]
    return float(value) if pd.notna(value) else None


def latest_local_release_date(path: str | Path = DATA_FILE) -> pd.Timestamp | None:
    df = load_cotton_end_stocks(path)
    if df.empty or "ReleaseDate" not in df.columns:
        return None

    latest = df["ReleaseDate"].max()
    return latest if pd.notna(latest) else None


def _parse_row_date(text: str) -> pd.Timestamp | None:
    match = re.search(
        r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(\d{4})\b",
        text,
    )
    if not match:
        return None
    return pd.Timestamp(datetime.strptime(match.group(0), "%b %d %Y").date())


def fetch_available_releases() -> list[WasdeRelease]:
    with _build_session() as session:
        html = _get(session, ARCHIVE_URL).text

    soup = BeautifulSoup(html, "html.parser")
    releases: dict[date, WasdeRelease] = {}

    for row in soup.find_all("tr"):
        release_date = _parse_row_date(row.get_text(" ", strip=True))
        if release_date is None:
            continue

        xml_links = [
            urljoin(BASE_URL, link["href"])
            for link in row.find_all("a", href=True)
            if link["href"].lower().endswith(".xml")
        ]
        if not xml_links:
            continue

        releases[release_date.date()] = WasdeRelease(
            release_date=release_date,
            xml_url=xml_links[0],
        )

    return sorted(releases.values(), key=lambda item: item.release_date)


def _clean_number(value: str | None) -> float:
    if value is None:
        raise ValueError("Missing numeric cell value.")
    match = re.search(r"-?\d+(?:\.\d+)?", value.replace(",", ""))
    if not match:
        raise ValueError(f"Could not parse numeric value from {value!r}.")
    return float(match.group(0))


def _wasde_number(root: ET.Element) -> int:
    for report in root.iter("Report"):
        page_title = report.attrib.get("page_title", "")
        match = re.search(r"WASDE\s*-\s*(\d+)", page_title)
        if match:
            return int(match.group(1))
    raise ValueError("Could not find WASDE number in XML.")


def _report_month(root: ET.Element) -> str:
    for report in root.iter("Report"):
        if report.attrib.get("sub_report_title") == "U.S. Cotton Supply and Use  1/":
            report_month = report.attrib.get("Report_Month")
            if report_month:
                return report_month.split()[0][:3]
    raise ValueError("Could not find U.S. Cotton report month in XML.")


def _market_year_label(value: str) -> str:
    match = re.search(r"\d{4}/\d{2,3}", value)
    if not match:
        raise ValueError(f"Could not parse market year from {value!r}.")
    return match.group(0)


def _parse_us_cotton_attribute_xml(
    xml_bytes: bytes,
    release_date: pd.Timestamp,
    attribute_name: str,
) -> dict:
    root = ET.fromstring(xml_bytes)
    wasde_number = _wasde_number(root)
    report_month = _report_month(root).lower()

    cotton_report = None
    for report in root.iter("Report"):
        if report.attrib.get("sub_report_title") == "U.S. Cotton Supply and Use  1/":
            cotton_report = report
            break
    if cotton_report is None:
        raise ValueError("Could not find U.S. Cotton Supply and Use table in XML.")

    attribute_row = None
    for elem in cotton_report.iter():
        if any(
            key.startswith("attribute") and value.strip() == attribute_name
            for key, value in elem.attrib.items()
        ):
            attribute_row = elem
            break
    if attribute_row is None:
        raise ValueError(f"Could not find {attribute_name} row in U.S. Cotton XML.")

    projected_rows = []
    for year_group in attribute_row.iter():
        market_year = next(
            (
                value
                for key, value in year_group.attrib.items()
                if key.startswith("market_year")
            ),
            None,
        )
        if not market_year or "proj" not in market_year.lower():
            continue

        month_group = next(
            (
                child
                for child in year_group.iter()
                if any(key.startswith("forecast_month") for key in child.attrib)
            ),
            None,
        )
        forecast_month = ""
        if month_group is not None:
            forecast_month = next(
                (
                    value
                    for key, value in month_group.attrib.items()
                    if key.startswith("forecast_month")
                ),
                "",
            )

        cell = next((cell for cell in year_group.iter("Cell")), None)
        value = _clean_number(next(iter(cell.attrib.values())) if cell is not None else None)
        projected_rows.append(
            {
                "market_year": _market_year_label(market_year),
                "forecast_month": forecast_month,
                "value": value,
            }
        )

    if not projected_rows:
        raise ValueError(f"No projected {attribute_name} values found in U.S. Cotton XML.")

    selected = projected_rows[-1]
    for row in projected_rows:
        if row["forecast_month"].lower()[:3] == report_month:
            selected = row

    return {
        "WasdeNumber": wasde_number,
        "ReportDate": pd.Timestamp(date(release_date.year, release_date.month, 1)),
        "ReportTitle": "U.S. Cotton Supply and Use",
        "Attribute": attribute_name,
        "ReliabilityProjection": None,
        "Commodity": "Cotton",
        "Region": "United States",
        "MarketYear": selected["market_year"],
        "ProjEstFlag": "Proj.",
        "AnnualQuarterFlag": "Annual",
        "Value": selected["value"],
        "Unit": UNIT,
        "ReleaseDate": release_date.normalize(),
        "ReleaseTime": "12:00:00",
        "ForecastYear": release_date.year,
        "ForecastMonth": release_date.month,
    }


def parse_us_cotton_ending_stocks_xml(
    xml_bytes: bytes,
    release_date: pd.Timestamp,
) -> dict:
    return _parse_us_cotton_attribute_xml(xml_bytes, release_date, "Ending Stocks")


def parse_us_cotton_use_xml(
    xml_bytes: bytes,
    release_date: pd.Timestamp,
) -> dict:
    return _parse_us_cotton_attribute_xml(xml_bytes, release_date, "Use, Total")


def refresh_usda_cotton_end_stocks(path: str | Path = DATA_FILE) -> dict:
    workbook_path = Path(path)
    existing = load_cotton_end_stocks(workbook_path)
    existing_use = load_cotton_use(workbook_path)
    latest_local = latest_local_release_date(workbook_path)
    releases = fetch_available_releases()

    pending = [
        release
        for release in releases
        if latest_local is None or release.release_date > latest_local
    ]

    if not pending:
        price_result = update_price_history_for_rows(existing)
        updated_existing = add_stocks_use_ratio(
            fill_prices_from_history(existing),
            existing_use,
        )
        if not updated_existing.equals(existing):
            _write_workbook(updated_existing, existing_use, workbook_path)

        return {
            "did_update": False,
            "rows_added": 0,
            "use_rows_added": 0,
            "latest_local": latest_local,
            "latest_remote": max((item.release_date for item in releases), default=None),
            "prices_updated": price_result["prices_updated"],
            "errors": price_result["errors"],
        }

    new_rows = []
    new_use_rows = []
    errors = []
    with _build_session() as session:
        for release in pending:
            try:
                xml_bytes = _get(session, release.xml_url).content
                row = parse_us_cotton_ending_stocks_xml(xml_bytes, release.release_date)
                row["Contract"] = contract_from_market_year(row["MarketYear"])
                new_rows.append(row)
                new_use_rows.append(parse_us_cotton_use_xml(xml_bytes, release.release_date))
            except Exception as exc:
                errors.append(f"{release.release_date:%Y-%m-%d}: {exc}")

    if not new_rows:
        return {
            "did_update": False,
            "rows_added": 0,
            "use_rows_added": 0,
            "latest_local": latest_local,
            "latest_remote": max((item.release_date for item in releases), default=None),
            "prices_updated": 0,
            "errors": errors,
        }

    new_rows_df = pd.DataFrame(new_rows)
    price_result = update_price_history_for_rows(new_rows_df)
    errors.extend(price_result["errors"])

    updated = pd.concat([existing, new_rows_df], ignore_index=True)
    updated["ReleaseDate"] = pd.to_datetime(updated["ReleaseDate"], errors="coerce")
    updated = (
        updated.sort_values(["ReleaseDate", "WasdeNumber"])
        .drop_duplicates(subset=["ReleaseDate"], keep="last")
        .reset_index(drop=True)
    )
    updated = fill_prices_from_history(updated)

    if new_use_rows:
        updated_use = pd.concat([existing_use, pd.DataFrame(new_use_rows)], ignore_index=True)
        updated_use["ReleaseDate"] = pd.to_datetime(updated_use["ReleaseDate"], errors="coerce")
        updated_use = (
            updated_use.sort_values(["ReleaseDate", "WasdeNumber"])
            .drop_duplicates(subset=["ReleaseDate"], keep="last")
            .reset_index(drop=True)
        )
    else:
        updated_use = existing_use.copy()

    _write_workbook(updated, updated_use, workbook_path)

    return {
        "did_update": True,
        "rows_added": len(updated) - len(existing),
        "use_rows_added": len(updated_use) - len(existing_use),
        "latest_local": latest_local,
        "latest_remote": updated["ReleaseDate"].max(),
        "prices_updated": price_result["prices_updated"],
        "errors": errors,
    }
