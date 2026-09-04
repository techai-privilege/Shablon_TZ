"""Read public trademark cards from the FIPS register."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup, Tag

from .network import MAX_PAGE_BYTES, GetSession, checked_image_content, limited_response_content

ALLOWED_FIPS_HOSTS = {"fips.ru", "www.fips.ru", "www1.fips.ru", "new.fips.ru"}
DEFAULT_TIMEOUT_SECONDS = 25
# A deliberately simple browser identifier is accepted more consistently by
# the DDoS protection in front of the public register than a fabricated full
# Chrome version string.
USER_AGENT = "Mozilla/5.0"


class FipsParseError(ValueError):
    """The supplied page is not a supported public FIPS trademark card."""


@dataclass(slots=True)
class NiceClass:
    number: str
    goods_and_services: str = ""


@dataclass(slots=True)
class TrademarkRecord:
    source_url: str
    database: str | None = None
    mark_name: str | None = None
    registration_number: str | None = None
    application_number: str | None = None
    status: str | None = None
    priority_date: str | None = None
    filing_date: str | None = None
    registration_date: str | None = None
    expiry_date: str | None = None
    owner: str | None = None
    applicant: str | None = None
    unprotected_elements: str | None = None
    nice_classes: list[NiceClass] = field(default_factory=list)
    image_url: str | None = None
    image_bytes: bytes | None = None
    image_content_type: str | None = None

    @property
    def nice_class_numbers(self) -> list[str]:
        return [item.number for item in self.nice_classes]


def _clean(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = " ".join(value.replace("\xa0", " ").split())
    return cleaned or None


def _validate_fips_url(url: str) -> str:
    parsed = urlparse(url.strip())
    if (
        parsed.scheme != "https"
        or (parsed.hostname or "").lower() not in ALLOWED_FIPS_HOSTS
        or parsed.port not in (None, 443)
        or parsed.username
        or parsed.password
    ):
        raise FipsParseError("Нужна HTTPS-ссылка на публичную карточку сайта fips.ru.")
    supported_path = (
        "registers-doc-view" in parsed.path or "fips_servl" in parsed.path
    ) and "fips_servlet" in parsed.path
    if not supported_path:
        raise FipsParseError("Ссылка не похожа на карточку реестра ФИПС.")
    return parsed.geturl()


def _paragraphs_by_code(soup: BeautifulSoup, code: str) -> list[Tag]:
    pattern = re.compile(rf"^\s*\({re.escape(code)}\)")
    return [
        p for p in soup.find_all("p") if pattern.search(p.get_text(" ", strip=True))
    ]


def _paragraph_by_code(soup: BeautifulSoup, code: str) -> Tag | None:
    paragraphs = _paragraphs_by_code(soup, code)
    return paragraphs[0] if paragraphs else None


def _bold_value_by_code(
    soup: BeautifulSoup, code: str, *, latest: bool = False
) -> str | None:
    paragraphs = _paragraphs_by_code(soup, code)
    paragraph = paragraphs[-1] if latest and paragraphs else (paragraphs[0] if paragraphs else None)
    if paragraph is None:
        return None
    bold = paragraph.find("b")
    return _clean(bold.get_text(" ", strip=True) if bold else None)


def _bold_value_by_label(soup: BeautifulSoup, label: str) -> str | None:
    paragraph = next(
        (p for p in soup.find_all("p") if label in p.get_text(" ", strip=True)),
        None,
    )
    if paragraph is None:
        return None
    bold = paragraph.find("b")
    return _clean(bold.get_text(" ", strip=True) if bold else None)


def _parse_nice_classes(soup: BeautifulSoup) -> list[NiceClass]:
    paragraph = _paragraph_by_code(soup, "511")
    if paragraph is None:
        return []

    classes: list[NiceClass] = []
    for bold in paragraph.find_all("b"):
        text = _clean(bold.get_text(" ", strip=True)) or ""
        match = re.match(r"^(\d{1,2})\s*(?:[-–—]\s*)?(.*)$", text)
        if match:
            classes.append(NiceClass(match.group(1).zfill(2), match.group(2).strip()))

    if classes:
        return classes

    # Some older cards expose the class list as one text block.
    text = paragraph.get_text(" ", strip=True)
    for match in re.finditer(r"(?:^|\s)(\d{1,2})\s*[-–—]\s*(.*?)(?=\s\d{1,2}\s*[-–—]|$)", text):
        classes.append(NiceClass(match.group(1).zfill(2), _clean(match.group(2)) or ""))
    return classes


def _find_mark_image_url(soup: BeautifulSoup, source_url: str) -> str | None:
    paragraph = _paragraph_by_code(soup, "540")
    container = paragraph or soup

    # Application cards normally link the full-size image from a thumbnail.
    # Prefer that link so the report does not contain the low-resolution "-s" file.
    sources = [link.get("href") for link in container.find_all("a")]
    sources.extend(image.get("src") for image in container.find_all("img"))
    for source in sources:
        if not isinstance(source, str) or not source:
            continue
        candidate = urljoin(source_url, source)
        parsed = urlparse(candidate)
        path = parsed.path.lower()
        if (
            parsed.scheme == "https"
            and (parsed.hostname or "").lower() in ALLOWED_FIPS_HOSTS
            and (
                "/ofpstorage/" in path
                or "/image/" in path
                or path.endswith((".jpg", ".jpeg", ".png", ".gif"))
            )
        ):
            return candidate
    return None


def parse_trademark_html(html: bytes | str, source_url: str) -> TrademarkRecord:
    """Parse a public RUTM/RUTMAP card without making network requests."""

    source_url = _validate_fips_url(source_url)
    if isinstance(html, bytes):
        # Public register pages currently declare and use Windows-1251.
        html = html.decode("windows-1251", errors="replace")

    soup = BeautifulSoup(html, "html.parser")
    page_text = _clean(soup.get_text(" ", strip=True)) or ""
    if "DDoS-Guard" in page_text or "Too Many Requests" in page_text:
        raise FipsParseError(
            "ФИПС временно ограничил частоту запросов. Подождите несколько минут и повторите."
        )
    if "(111)" not in page_text and "(210)" not in page_text:
        raise FipsParseError("ФИПС вернул страницу без данных товарного знака.")

    status_node = soup.select_one("tr.Status td")
    status_text = _clean(status_node.get_text(" ", strip=True) if status_node else None)
    if status_text:
        status_text = re.sub(
            r"^(?:Статус|Состояние делопроизводства):\s*",
            "",
            status_text,
            flags=re.IGNORECASE,
        )
        status_text = re.sub(
            r"\s*\(последнее изменение(?: статуса)?:.*\)\s*$",
            "",
            status_text,
            flags=re.IGNORECASE,
        )

    parsed_url = urlparse(source_url)
    db_match = re.search(r"(?:^|&)DB=([^&]+)", parsed_url.query, re.IGNORECASE)
    record = TrademarkRecord(
        source_url=source_url,
        database=db_match.group(1).upper() if db_match else None,
        registration_number=_bold_value_by_code(soup, "111"),
        application_number=_bold_value_by_code(soup, "210"),
        status=_clean(status_text),
        priority_date=_bold_value_by_label(soup, "Приоритет:"),
        filing_date=_bold_value_by_code(soup, "220") or _bold_value_by_code(soup, "200"),
        registration_date=_bold_value_by_code(soup, "151"),
        expiry_date=_bold_value_by_code(soup, "181"),
        # The base card contains the original holder. Later register events
        # (assignment, succession, or a name/address change) append another
        # (732) block at the bottom, so the final occurrence is authoritative.
        owner=_bold_value_by_code(soup, "732", latest=True),
        applicant=_bold_value_by_code(soup, "731"),
        unprotected_elements=_bold_value_by_code(soup, "526"),
        nice_classes=_parse_nice_classes(soup),
        image_url=_find_mark_image_url(soup, source_url),
    )
    if not (record.registration_number or record.application_number):
        raise FipsParseError("В карточке не найден номер регистрации или заявки.")
    return record


def fetch_trademark(
    url: str,
    *,
    include_image: bool = True,
    session: GetSession | None = None,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> TrademarkRecord:
    """Download and parse a public FIPS card, optionally including its image."""

    url = _validate_fips_url(url)
    owns_client = session is None
    client = session or requests.Session()
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "ru-RU,ru;q=0.9"}
    try:
        try:
            response = client.get(url, headers=headers, timeout=timeout)
            response.raise_for_status()
            page = limited_response_content(response, MAX_PAGE_BYTES, "Страница ФИПС")
        except (requests.RequestException, ValueError) as exc:
            raise FipsParseError(f"Не удалось загрузить карточку ФИПС: {exc}") from exc

        record = parse_trademark_html(page, response.url)
        if include_image and record.image_url:
            parsed_image = urlparse(record.image_url)
            if (
                parsed_image.scheme != "https"
                or (parsed_image.hostname or "").lower() not in ALLOWED_FIPS_HOSTS
                or parsed_image.port not in (None, 443)
                or parsed_image.username
                or parsed_image.password
            ):
                raise FipsParseError("Карточка содержит изображение с неподдерживаемого сайта.")
            try:
                image_response = client.get(record.image_url, headers=headers, timeout=timeout)
                image_response.raise_for_status()
                image_bytes, content_type = checked_image_content(image_response, "ФИПС")
            except (requests.RequestException, ValueError) as exc:
                raise FipsParseError(f"Данные карточки получены, но изображение не загрузилось: {exc}") from exc
            record.image_bytes = image_bytes
            record.image_content_type = content_type
        return record
    finally:
        if owns_client and isinstance(client, requests.Session):
            client.close()
