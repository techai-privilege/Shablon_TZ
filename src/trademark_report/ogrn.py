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


def _pdf_text(content: bytes) -> str:
    reader = PdfReader(BytesIO(content))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def _remaining_timeout(deadline: float) -> float:
    return max(0.5, min(8.0, deadline - time.monotonic()))


def _fetch_address(
    row: dict,
    client,
    headers: dict[str, str],
    timeout: float,
    *,
    attempts: int = 3,
) -> tuple[str, str]:
    token = str(row.get("t", "")).strip()
    if not token:
        return "", "ФНС не вернула идентификатор выписки"

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
                return "", "ФНС запросила CAPTCHA"

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
                address = extract_registered_address(_pdf_text(downloaded.content))
                if address:
                    return address, ""
                last_error = "в выписке не найден публичный адрес"
        except requests.Timeout:
            last_error = "ФНС не ответила вовремя"
        except requests.RequestException as exc:
            last_error = f"ошибка связи с ФНС: {exc}"
        except (ValueError, OSError, PdfReadError):
            last_error = "не удалось прочитать PDF-выписку ФНС"

        if attempt + 1 < attempts and time.monotonic() < overall_deadline:
            time.sleep(min(1.5 * (attempt + 1), max(0.0, overall_deadline - time.monotonic())))
    return "", last_error


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
    address, address_error = _fetch_address(row, client, headers, timeout)
    if address:
        _ADDRESS_CACHE[inn] = address
    elif cached_address:
        address = cached_address
        address_error = ""
    return FnsRegistrationData(
        ogrn=ogrn_values.pop(),
        address=address,
        address_error=address_error,
        name=_registration_name(row.get("n")),
    )


def fetch_ogrn(inn_value: str, *, timeout: float = 20.0, session=None) -> str:
    """Backward-compatible helper returning only OGRN/OGRNIP."""

    return fetch_registration_data(inn_value, timeout=timeout, session=session).ogrn
