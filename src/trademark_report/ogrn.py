"""Lookup registration details by INN through FNS."""

from __future__ import annotations

from dataclasses import dataclass
from html import unescape
from io import BytesIO
import re
import time

from pypdf import PdfReader
from pypdf.errors import PdfReadError
import requests


FNS_SEARCH_URL = "https://egrul.nalog.ru/"


class OgrnLookupError(RuntimeError):
    """Raised when FNS cannot return a unique registration record for an INN."""


@dataclass(frozen=True)
class FnsRegistrationData:
    ogrn: str
    address: str = ""
    address_error: str = ""
    name: str = ""


_ADDRESS_CACHE: dict[str, str] = {}


def normalize_inn(value: str) -> str:
    return re.sub(r"\D", "", value)


def _registration_name(value: object) -> str:
    """Normalize the organization/IP name returned in an FNS search row."""

    text = unescape(str(value or ""))
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip(" ,;")


def prefer_full_registration_name(current: str, fetched: str) -> str:
    """Prefer the more complete name when one source returns only its fragment."""

    current = re.sub(r"\s+", " ", current or "").strip()
    fetched = re.sub(r"\s+", " ", fetched or "").strip()
    if not current:
        return fetched
    if not fetched:
        return current

    def comparable(value: str) -> str:
        value = value.casefold().replace("ё", "е")
        value = value.translate(str.maketrans({"«": '"', "»": '"', "„": '"', "“": '"'}))
        return re.sub(r"[^0-9a-zа-я]+", " ", value).strip()

    current_key = comparable(current)
    fetched_key = comparable(fetched)
    if fetched_key and fetched_key in current_key and len(fetched_key) < len(current_key):
        return current
    if current_key and current_key in fetched_key and len(current_key) < len(fetched_key):
        return fetched
    return fetched


def _valid_checksum(digits: str) -> bool:
    numbers = [int(item) for item in digits]
    if len(numbers) == 10:
        weights = (2, 4, 10, 3, 5, 9, 4, 6, 8)
        return sum(a * b for a, b in zip(numbers[:9], weights)) % 11 % 10 == numbers[9]
    if len(numbers) == 12:
        weights_11 = (7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
        weights_12 = (3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8)
        check_11 = sum(a * b for a, b in zip(numbers[:10], weights_11)) % 11 % 10
        check_12 = sum(a * b for a, b in zip(numbers[:11], weights_12)) % 11 % 10
        return check_11 == numbers[10] and check_12 == numbers[11]
    return False


def validate_inn(value: str) -> str:
    inn = normalize_inn(value)
    if len(inn) not in (10, 12):
        raise OgrnLookupError("ИНН должен содержать 10 цифр для организации или 12 цифр для ИП.")
    if not _valid_checksum(inn):
        raise OgrnLookupError("Контрольная сумма ИНН не совпадает. Проверьте номер.")
    return inn


def extract_registered_address(text: str) -> str:
    """Extract the current legal-entity address from FNS statement text."""

    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]
    heading = re.compile(
        r"^(?:\d+\s+)?адрес(?:\s+\(место\s+нахождения\))?\s+"
        r"юридического\s+лица(?:\s+(.*))?$",
        re.IGNORECASE,
    )
    match = next(
        ((index, found) for index, line in enumerate(lines) if (found := heading.match(line))),
        None,
    )
    start = match[0] + 1 if match else None
    if start is None:
        return ""

    parts: list[str] = []
    inline_value = (match[1].group(1) or "").strip(" ,")
    if inline_value:
        parts.append(inline_value)
    for line in lines[start:]:
        if re.match(r"^(?:\d+\s+)?грн и дата внесения", line, re.IGNORECASE):
            break
        if line:
            parts.append(line.strip(" ,"))
    # The extract table puts the ordinal of the next row immediately after the address.
    if parts and re.fullmatch(r"\d{1,3}", parts[-1]):
        parts.pop()
    return re.sub(r"\s*,\s*", ", ", ", ".join(parts)).strip(" ,")


def extract_full_registration_name(text: str) -> str:
    """Extract the full legal name from an FNS statement."""

    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines()]

    # The statement header is the most stable location: the full name is placed
    # before the explanatory label and may wrap across several PDF text lines.
    for index, line in enumerate(lines):
        if line.casefold() == "полное наименование юридического лица" and index:
            marker_index = next(
                (
                    previous
                    for previous in range(index - 1, -1, -1)
                    if lines[previous].casefold().startswith(
                        "настоящая выписка содержит сведения"
                    )
                ),
                None,
            )
            if marker_index is not None:
                header = " ".join(
                    item.strip(" ,")
                    for item in lines[marker_index:index]
                    if item.strip(" ,")
                )
                value = re.sub(
                    r"^настоящая\s+выписка\s+содержит\s+сведения\s+"
                    r"о\s+юридическом\s+лице\s*",
                    "",
                    header,
                    flags=re.IGNORECASE,
                )
            else:
                value = lines[index - 1].strip(" ,")
            if value and not re.fullmatch(r"\d+", value):
                return re.sub(r"\s+", " ", value).strip(" ,")

    heading = re.compile(
        r"^(?:\d+\s+)?полное наименование(?:\s+на русском языке)?(?:\s+(.*))?$",
        re.IGNORECASE,
    )
    for index, line in enumerate(lines):
        match = heading.match(line)
        if not match:
            continue
        parts: list[str] = []
        inline_value = (match.group(1) or "").strip(" ,")
        if inline_value:
            parts.append(inline_value)
        for value in lines[index + 1 :]:
            if re.match(r"^(?:\d+\s+)?грн и дата внесения", value, re.IGNORECASE):
                break
            if value:
                parts.append(value.strip(" ,"))
        if parts:
            return re.sub(r"\s+", " ", " ".join(parts)).strip(" ,")
    return ""


def _pdf_text(content: bytes) -> str:
    reader = PdfReader(BytesIO(content))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _remaining_timeout(deadline: float) -> float:
    return max(0.5, min(8.0, deadline - time.monotonic()))


def _fetch_statement_details(
    row: dict,
    client,
    headers: dict[str, str],
    timeout: float,
    *,
    attempts: int = 3,
) -> tuple[str, str, str]:
    token = str(row.get("t", "")).strip()
    if not token:
        return "", "", "ФНС не вернула идентификатор выписки"

    attempts = max(1, attempts)
    overall_deadline = time.monotonic() + max(timeout, 24.0)
    last_error = "выписка не была подготовлена"
    for attempt in range(attempts):
        remaining_attempts = attempts - attempt
        remaining_total = overall_deadline - time.monotonic()
        if remaining_total <= 0:
            break
        attempt_deadline = min(
            overall_deadline,
            time.monotonic() + max(4.0, remaining_total / remaining_attempts),
        )
        try:
            requested = client.get(
                f"{FNS_SEARCH_URL}vyp-request/{token}",
                headers=headers,
                timeout=_remaining_timeout(attempt_deadline),
            )
            requested.raise_for_status()
            if requested.json().get("captchaRequired"):
                return "", "", "ФНС запросила CAPTCHA"

            while time.monotonic() < attempt_deadline:
                status = client.get(
                    f"{FNS_SEARCH_URL}vyp-status/{token}",
                    headers=headers,
                    timeout=_remaining_timeout(attempt_deadline),
                )
                status.raise_for_status()
                state = status.json().get("status")
                if state == "ready":
                    break
                if state == "error":
                    last_error = "ФНС вернула ошибку при формировании выписки"
                    break
                time.sleep(0.5)
            else:
                last_error = "ФНС не успела подготовить выписку"
                state = "timeout"

            if state == "ready":
                downloaded = client.get(
                    f"{FNS_SEARCH_URL}vyp-download/{token}",
                    headers=headers,
                    timeout=_remaining_timeout(attempt_deadline),
                )
                downloaded.raise_for_status()
                statement_text = _pdf_text(downloaded.content)
                address = extract_registered_address(statement_text)
                full_name = extract_full_registration_name(statement_text)
                if address:
                    return address, full_name, ""
                last_error = "в выписке не найден публичный адрес"
                if full_name:
                    return "", full_name, last_error
        except requests.Timeout:
            last_error = "ФНС не ответила вовремя"
        except requests.RequestException as exc:
            last_error = f"ошибка связи с ФНС: {exc}"
        except (ValueError, OSError, PdfReadError):
            last_error = "не удалось прочитать PDF-выписку ФНС"

        if attempt + 1 < attempts and time.monotonic() < overall_deadline:
            time.sleep(min(1.5 * (attempt + 1), max(0.0, overall_deadline - time.monotonic())))
    return "", "", last_error


def fetch_registration_data(
    inn_value: str, *, timeout: float = 20.0, session=None
) -> FnsRegistrationData:
    """Return name, OGRN/OGRNIP and, when public, the registered address from FNS."""

    inn = validate_inn(inn_value)
    client = session or requests.Session()
    headers = {"User-Agent": "Mozilla/5.0 TrademarkConsentDesktop/1.0"}
    try:
        response = client.post(
            FNS_SEARCH_URL,
            data={"query": inn, "region": "", "PreventChromeAutocomplete": ""},
            headers=headers,
            timeout=timeout,
        )
        response.raise_for_status()
        payload = response.json()
        if payload.get("captchaRequired"):
            raise OgrnLookupError("ФНС запросила проверочный код. Введите данные вручную.")
        token = payload.get("t")
        if not token:
            raise OgrnLookupError("ФНС не вернула идентификатор поиска.")
        result = client.get(
            f"{FNS_SEARCH_URL}search-result/{token}", headers=headers, timeout=timeout
        )
        result.raise_for_status()
        rows = result.json().get("rows") or []
    except OgrnLookupError:
        raise
    except (requests.RequestException, ValueError) as exc:
        raise OgrnLookupError(f"Не удалось получить данные из ФНС: {exc}") from exc

    exact = [row for row in rows if normalize_inn(str(row.get("i", ""))) == inn]
    if not exact:
        raise OgrnLookupError("ФНС не нашла организацию или ИП с указанным ИНН.")
    ogrn_values = {str(row.get("o", "")).strip() for row in exact if row.get("o")}
    if len(ogrn_values) != 1:
        raise OgrnLookupError("ФНС не вернула единственный ОГРН. Введите его вручную.")
    row = exact[0]
    cached_address = _ADDRESS_CACHE.get(inn, "")
    address, full_name, address_error = _fetch_statement_details(
        row, client, headers, timeout
    )
    if address:
        _ADDRESS_CACHE[inn] = address
    elif cached_address:
        address = cached_address
        address_error = ""
    return FnsRegistrationData(
        ogrn=ogrn_values.pop(),
        address=address,
        address_error=address_error,
        name=full_name or _registration_name(row.get("n")),
    )


def fetch_ogrn(inn_value: str, *, timeout: float = 20.0, session=None) -> str:
    """Backward-compatible helper returning only OGRN/OGRNIP."""

    return fetch_registration_data(inn_value, timeout=timeout, session=session).ogrn
