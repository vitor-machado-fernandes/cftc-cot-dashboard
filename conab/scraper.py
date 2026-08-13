"""Crawl and download CONAB crop progress workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import hashlib
import logging
from pathlib import Path
import re
import time
from typing import Iterable, Any
import unicodedata
from urllib.parse import parse_qs, urljoin, urlparse

import requests

LOGGER = logging.getLogger(__name__)

INDEX_URL = (
    "https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/"
    "safras/progresso-de-safra"
)
USER_AGENT = "CoT Streamlit CONAB progress updater/1.0"


@dataclass(frozen=True)
class Bulletin:
    """A CONAB weekly bulletin with an optional workbook attachment."""

    page_url: str
    title: str
    date_range: str
    bulletin_date: date | None
    xlsx_url: str


def _ascii_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text.lower())
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def _url_suffix(url: str) -> str:
    return Path(urlparse(url).path).suffix.lower()


def _is_workbook_url(url: str) -> bool:
    path = _ascii_text(urlparse(url).path)
    return _url_suffix(url) in {".xlsx", ".xls"} or "plantio-e-colheita" in path


def _is_attachment_url(url: str) -> bool:
    return _url_suffix(url) in {".xlsx", ".xls", ".pdf"} or _is_workbook_url(url)


def _soup(html: str) -> Any:
    from bs4 import BeautifulSoup

    return BeautifulSoup(html, "html.parser")


def parse_brazilian_date(text: str) -> date | None:
    """Parse common Brazilian date forms found in CONAB page text."""

    normalized = _ascii_text(text)
    month_map = {
        "janeiro": 1,
        "fevereiro": 2,
        "marco": 3,
        "abril": 4,
        "maio": 5,
        "junho": 6,
        "julho": 7,
        "agosto": 8,
        "setembro": 9,
        "outubro": 10,
        "novembro": 11,
        "dezembro": 12,
    }

    numeric_dates = re.findall(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", normalized)
    if numeric_dates:
        day, month, year = numeric_dates[-1]
        year_int = int(year)
        if year_int < 100:
            year_int += 2000
        try:
            return date(year_int, int(month), int(day))
        except ValueError:
            return None

    month_pattern = "|".join(month_map)
    verbose_dates = re.findall(
        rf"(\d{{1,2}})\s+de\s+({month_pattern})\s+de\s+(\d{{4}})",
        normalized,
    )
    if verbose_dates:
        day, month_name, year = verbose_dates[-1]
        try:
            return date(int(year), month_map[month_name], int(day))
        except ValueError:
            return None

    return None


def _session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def _get_with_retries(
    session: requests.Session,
    url: str,
    *,
    retries: int = 3,
    backoff: float = 1.5,
    timeout: int = 30,
) -> requests.Response:
    last_exc: Exception | None = None
    for attempt in range(retries):
        try:
            response = session.get(url, timeout=timeout)
            response.raise_for_status()
            return response
        except requests.RequestException as exc:
            last_exc = exc
            if attempt < retries - 1:
                time.sleep(backoff * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url}") from last_exc


def _page_links(soup: Any, base_url: str) -> Iterable[str]:
    for anchor in soup.find_all("a", href=True):
        href = str(anchor["href"]).strip()
        if href:
            yield urljoin(base_url, href)


def _find_bulletin_pages(soup: Any, base_url: str) -> list[str]:
    pages: list[str] = []
    seen: set[str] = set()
    for anchor in soup.find_all("a", href=True):
        text = anchor.get_text(" ", strip=True)
        href = urljoin(base_url, str(anchor["href"]).strip())
        haystack = _ascii_text(f"{text} {href}")
        if "plantio" in haystack or "colheita" in haystack or "progresso" in haystack:
            if not _is_attachment_url(href) and href not in seen:
                pages.append(href)
                seen.add(href)
    return pages


def _extract_index_workbooks(soup: Any, base_url: str) -> list[Bulletin]:
    """Extract direct Plantio e Colheita workbook links from the paginated index."""

    bulletins: list[Bulletin] = []
    seen: set[str] = set()
    for anchor in soup.find_all("a", href=True):
        text = anchor.get_text(" ", strip=True)
        lower_text = _ascii_text(text)
        href = urljoin(base_url, str(anchor["href"]).strip())
        if not _is_workbook_url(href):
            continue
        if "plantio" not in lower_text and "colheita" not in lower_text:
            continue
        if href in seen:
            continue

        heading = anchor.find_previous(["h1", "h2", "h3"])
        title = heading.get_text(" ", strip=True) if heading else text
        parent = anchor.find_parent()
        nearby_text = parent.get_text(" ", strip=True) if parent else text
        bulletin_date = parse_brazilian_date(f"{title} {nearby_text} {text} {href}")
        bulletins.append(
            Bulletin(
                page_url=base_url,
                title=title,
                date_range=text,
                bulletin_date=bulletin_date,
                xlsx_url=href,
            )
        )
        seen.add(href)
    return bulletins


def _find_next_page(soup: Any, current_url: str) -> str | None:
    for anchor in soup.find_all("a", href=True):
        text = _ascii_text(anchor.get_text(" ", strip=True))
        rel = " ".join(anchor.get("rel", [])).lower()
        href = urljoin(current_url, str(anchor["href"]).strip())
        if "proxima" in text or "next" in text or "next" in rel:
            return href

    parsed = urlparse(current_url)
    current_start = int(parse_qs(parsed.query).get("b_start:int", ["0"])[0])
    candidates: list[tuple[int, str]] = []
    for href in _page_links(soup, current_url):
        query = parse_qs(urlparse(href).query)
        if "b_start:int" not in query:
            continue
        try:
            start = int(query["b_start:int"][0])
        except (TypeError, ValueError):
            continue
        if start > current_start:
            candidates.append((start, href))
    return min(candidates)[1] if candidates else None


def _extract_workbook(page_html: str, page_url: str) -> tuple[str, str] | None:
    soup = _soup(page_html)
    fallback: tuple[str, str] | None = None
    for anchor in soup.find_all("a", href=True):
        text = anchor.get_text(" ", strip=True)
        href = urljoin(page_url, str(anchor["href"]).strip())
        lower_text = _ascii_text(text)
        if not _is_workbook_url(href):
            continue
        candidate = (text, href)
        if "plantio" in lower_text and "colheita" in lower_text:
            return candidate
        if fallback is None and ("plantio" in lower_text or "colheita" in lower_text):
            fallback = candidate
        elif fallback is None:
            fallback = candidate
    return fallback


def crawl_bulletins(
    *,
    index_url: str = INDEX_URL,
    max_pages: int = 250,
    delay_seconds: float = 0.4,
    stop_after_date: date | None = None,
) -> list[Bulletin]:
    """Crawl paginated CONAB bulletin pages and return workbook links."""

    session = _session()
    bulletins: list[Bulletin] = []
    seen_index_pages: set[str] = set()
    url: str | None = index_url

    while url and len(seen_index_pages) < max_pages:
        if url in seen_index_pages:
            break
        seen_index_pages.add(url)
        LOGGER.info("Crawling CONAB index page: %s", url)
        response = _get_with_retries(session, url)
        soup = _soup(response.text)
        saw_stored_bulletin = False
        added_new_bulletin = False

        direct_bulletins = _extract_index_workbooks(soup, url)
        for bulletin in direct_bulletins:
            if stop_after_date and bulletin.bulletin_date and bulletin.bulletin_date <= stop_after_date:
                saw_stored_bulletin = True
                continue
            bulletins.append(bulletin)
            added_new_bulletin = True

        detail_pages = [] if direct_bulletins else _find_bulletin_pages(soup, url)
        for page_url in detail_pages:
            try:
                time.sleep(delay_seconds)
                page_response = _get_with_retries(session, page_url)
            except RuntimeError as exc:
                LOGGER.warning("Skipping bulletin page %s: %s", page_url, exc)
                continue

            page_soup = _soup(page_response.text)
            title = page_soup.get_text(" ", strip=True)[:500]
            workbook = _extract_workbook(page_response.text, page_url)
            bulletin_date = parse_brazilian_date(title)
            if stop_after_date and bulletin_date and bulletin_date <= stop_after_date:
                saw_stored_bulletin = True
                continue
            if not workbook:
                LOGGER.info("No workbook attachment found for %s", page_url)
                continue
            link_text, xlsx_url = workbook
            bulletins.append(
                Bulletin(
                    page_url=page_url,
                    title=title,
                    date_range=link_text,
                    bulletin_date=bulletin_date,
                    xlsx_url=xlsx_url,
                )
            )
            added_new_bulletin = True

        time.sleep(delay_seconds)
        if stop_after_date and saw_stored_bulletin and not added_new_bulletin:
            break
        url = _find_next_page(soup, url)

    unique: dict[str, Bulletin] = {}
    for bulletin in bulletins:
        unique[bulletin.xlsx_url] = bulletin
    return sorted(
        unique.values(),
        key=lambda b: b.bulletin_date or date.min,
        reverse=True,
    )


def _safe_filename(url: str) -> str:
    parsed_name = Path(urlparse(url).path).name
    suffix = Path(parsed_name).suffix or ".xlsx"
    stem = Path(parsed_name).stem or "conab-progress"
    digest = hashlib.sha256(url.encode("utf-8")).hexdigest()[:10]
    safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "-", stem).strip("-")[:90]
    return f"{safe_stem}-{digest}{suffix}"


def download_workbook(
    url: str,
    *,
    cache_dir: str | Path = "data/raw",
    delay_seconds: float = 0.5,
    force: bool = False,
) -> Path:
    """Download an XLS/XLSX workbook once, returning its local cache path."""

    cache_path = Path(cache_dir)
    cache_path.mkdir(parents=True, exist_ok=True)
    output = cache_path / _safe_filename(url)
    if output.exists() and output.stat().st_size > 0 and not force:
        return output

    time.sleep(delay_seconds)
    session = _session()
    response = _get_with_retries(session, url)
    output.write_bytes(response.content)
    return output


def download_bulletins(
    bulletins: Iterable[Bulletin],
    *,
    cache_dir: str | Path = "data/raw",
    delay_seconds: float = 0.5,
) -> list[tuple[Bulletin, Path]]:
    """Download each bulletin workbook, skipping failed files."""

    downloaded: list[tuple[Bulletin, Path]] = []
    for bulletin in bulletins:
        try:
            path = download_workbook(
                bulletin.xlsx_url,
                cache_dir=cache_dir,
                delay_seconds=delay_seconds,
            )
            downloaded.append((bulletin, path))
        except RuntimeError as exc:
            LOGGER.warning("Could not download %s: %s", bulletin.xlsx_url, exc)
    return downloaded
