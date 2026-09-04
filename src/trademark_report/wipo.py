"""Read individual international trademark records from WIPO Madrid Monitor."""

from __future__ import annotations

import re
from urllib.parse import parse_qs, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from trademark_report.fips import NiceClass, TrademarkRecord
from trademark_report.network import (
    MAX_PAGE_BYTES,
    GetSession,
    checked_image_content,
    limited_response_content,
)

ALLOWED_WIPO_HOSTS = {"www3.wipo.int"}
DEFAULT_TIMEOUT_SECONDS = 25
USER_AGENT = "Mozilla/5.0"


class WipoParseError(ValueError):
    """The supplied page is not a supported Madrid Monitor record."""


def _clean(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = " ".join(value.replace("\xa0", " ").split())
    return cleaned or None


def _validate_wipo_url(url: str) -> str:
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower()
    if (
        parsed.scheme != "https"
        or host not in ALLOWED_WIPO_HOSTS
        or parsed.port not in (None, 443)
        or parsed.username
        or parsed.password
    ):
        raise WipoParseError(
            "Нужна HTTPS-ссылка на запись WIPO Madrid Monitor (www3.wipo.int)."
        )
    if not parsed.path.endswith("/madrid/monitor/en/showData.jsp"):
        raise WipoParseError("Ссылка не похожа на карточку WIPO Madrid Monitor.")
    record_ids = parse_qs(parsed.query).get("ID", [])
    if not record_ids or not re.fullmatch(r"ROM\.\d+", record_ids[0], re.IGNORECASE):
        raise WipoParseError("В ссылке WIPO не найден номер международной регистрации.")
    return parsed.geturl()


def _nice_classes(value: str | None) -> list[NiceClass]:
    numbers = {
        int(number)
        for number in re.findall(r"\b\d{1,2}\b", value or "")
        if 1 <= int(number) <= 45
    }
    return [NiceClass(str(number).zfill(2)) for number in sorted(numbers)]


def parse_wipo_html(html: bytes | str, source_url: str) -> TrademarkRecord:
    """Parse one Madrid Monitor record without making network requests."""

    source_url = _validate_wipo_url(source_url)
    if isinstance(html, bytes):
        html = html.decode("utf-8", errors="replace")

    soup = BeautifulSoup(html, "html.parser")
    page_text = _clean(soup.get_text(" ", strip=True)) or ""
    if "Too Many Requests" in page_text or "Access Denied" in page_text:
        raise WipoParseError(
            "WIPO временно ограничил загрузку. Подождите несколько минут и повторите."
        )

    row = soup.select_one("#markInformationHeader table.markInformation tbody tr")
    if row is None:
        raise WipoParseError("WIPO вернул страницу без данных международной регистрации.")

    parsed_url = urlparse(source_url)
    url_number = parse_qs(parsed_url.query)["ID"][0].split(".", 1)[1]
    header = soup.select_one("#headerStatusContainer td.markname h3")
    header_text = _clean(header.get_text(" ", strip=True) if header else None) or ""
    header_match = re.match(r"^(\d+)\s*-\s*(.*)$", header_text)
    registration_number = header_match.group(1) if header_match else url_number
    mark_name = _clean(header_match.group(2)) if header_match else None

    image = row.select_one("td.mark img")
    image_source = image.get("src") if image else None
    image_url = (
        urljoin(source_url, image_source)
        if isinstance(image_source, str) and image_source
        else None
    )
    if not mark_name and image:
        image_alt = image.get("alt")
        mark_name = _clean(image_alt) if isinstance(image_alt, str) else None

    date_nodes = row.select("td.date")
    registration_date = _clean(date_nodes[0].get_text(" ", strip=True)) if date_nodes else None
    owner_node = row.select_one("td.name")
    classes_node = row.select_one("td.nice")

    record = TrademarkRecord(
        source_url=source_url,
        database="ROM",
        mark_name=mark_name,
        registration_number=registration_number,
        registration_date=registration_date,
        owner=_clean(owner_node.get_text(" ", strip=True) if owner_node else None),
        nice_classes=_nice_classes(
            classes_node.get_text(" ", strip=True) if classes_node else None
        ),
        image_url=image_url,
    )
    if not record.registration_number:
        raise WipoParseError("В карточке WIPO не найден номер международной регистрации.")
    return record


def fetch_wipo_trademark(
    url: str,
    *,
    include_image: bool = True,
    session: GetSession | None = None,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> TrademarkRecord:
    """Download one user-selected Madrid Monitor record and its mark image."""

    url = _validate_wipo_url(url)
    owns_client = session is None
    client = session or requests.Session()
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    try:
        try:
            response = client.get(url, headers=headers, timeout=timeout)
            response.raise_for_status()
            page = limited_response_content(response, MAX_PAGE_BYTES, "Страница WIPO")
        except (requests.RequestException, ValueError) as exc:
            raise WipoParseError(f"Не удалось загрузить карточку WIPO: {exc}") from exc

        record = parse_wipo_html(page, response.url)
        if include_image and record.image_url:
            image_url = record.image_url
            parsed_image = urlparse(image_url)
            if (
                parsed_image.scheme != "https"
                or (parsed_image.hostname or "").lower() not in ALLOWED_WIPO_HOSTS
                or parsed_image.port not in (None, 443)
                or parsed_image.username
                or parsed_image.password
                or not parsed_image.path.endswith("/madrid/monitor/jsp/data.jsp")
            ):
                raise WipoParseError("Карточка WIPO содержит неподдерживаемую ссылку на изображение.")
            try:
                image_response = client.get(image_url, headers=headers, timeout=timeout)
                image_response.raise_for_status()
                image_bytes, content_type = checked_image_content(image_response, "WIPO")
            except (requests.RequestException, ValueError) as exc:
                raise WipoParseError(
                    f"Данные WIPO получены, но изображение не загрузилось: {exc}"
                ) from exc
            record.image_bytes = image_bytes
            record.image_content_type = content_type
        return record
    finally:
        if owns_client and isinstance(client, requests.Session):
            client.close()
