import pytest
import requests

from trademark_report import ogrn as ogrn_module
from trademark_report.ogrn import (
    OgrnLookupError,
    extract_registered_address,
    fetch_ogrn,
    fetch_registration_data,
    validate_inn,
)


class _Response:
    def __init__(self, payload, content=b""):
        self.payload = payload
        self.content = content

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


class _Session:
    def post(self, *args, **kwargs):
        return _Response({"t": "token", "captchaRequired": False})

    def get(self, url, *args, **kwargs):
        if "search-result" in url:
            return _Response({
                "rows": [
                    {
                        "i": "7707083893",
                        "o": "1027700132195",
                        "t": "statement-token",
                        "n": "ООО &quot;ТЕСТ&quot;",
                    }
                ]
            })
        if "vyp-request" in url:
            return _Response({"captchaRequired": False})
        if "vyp-status" in url:
            return _Response({"status": "ready"})
        if "vyp-download" in url:
            return _Response({}, content=b"mock-pdf")
        raise AssertionError(url)


def test_fns_lookup_returns_only_ogrn(monkeypatch):
    ogrn_module._ADDRESS_CACHE.clear()
    monkeypatch.setattr(
        ogrn_module,
        "_pdf_text",
        lambda _content: "Адрес юридического лица\n117312\n6\nГРН и дата внесения",
    )
    assert fetch_ogrn("7707083893", session=_Session()) == "1027700132195"


def test_fns_lookup_returns_ogrn_and_registered_address(monkeypatch):
    monkeypatch.setattr(
        ogrn_module,
        "_pdf_text",
        lambda _content: """
Адрес юридического лица
117312,
Г.МОСКВА,
УЛ. ВАВИЛОВА,
Д.19
6
ГРН и дата внесения в ЕГРЮЛ записи
""",
    )

    result = fetch_registration_data("7707083893", session=_Session())

    assert result.ogrn == "1027700132195"
    assert result.address == "117312, Г.МОСКВА, УЛ. ВАВИЛОВА, Д.19"
    assert result.name == 'ООО "ТЕСТ"'


def test_address_parser_does_not_return_unrelated_statement_text():
    assert extract_registered_address("Сведения о регистрации") == ""


def test_address_statement_is_retried_after_transient_network_error(monkeypatch):
    class FlakySession(_Session):
        statement_requests = 0

        def get(self, url, *args, **kwargs):
            if "vyp-request" in url:
                self.statement_requests += 1
                if self.statement_requests == 1:
                    raise requests.ConnectionError("временный сбой")
            return super().get(url, *args, **kwargs)

    session = FlakySession()
    ogrn_module._ADDRESS_CACHE.clear()
    monkeypatch.setattr(ogrn_module.time, "sleep", lambda _seconds: None)
    monkeypatch.setattr(
        ogrn_module,
        "_pdf_text",
        lambda _content: """
Адрес юридического лица
344002,
Г. РОСТОВ-НА-ДОНУ
6
ГРН и дата внесения в ЕГРЮЛ записи
""",
    )

    result = fetch_registration_data("7707083893", session=session)

    assert session.statement_requests == 2
    assert result.address == "344002, Г. РОСТОВ-НА-ДОНУ"
    assert result.address_error == ""


def test_address_failure_exposes_specific_reason(monkeypatch):
    class CaptchaSession(_Session):
        def get(self, url, *args, **kwargs):
            if "vyp-request" in url:
                return _Response({"captchaRequired": True})
            return super().get(url, *args, **kwargs)

    ogrn_module._ADDRESS_CACHE.clear()
    result = fetch_registration_data("7707083893", session=CaptchaSession())

    assert result.address == ""
    assert "CAPTCHA" in result.address_error


def test_inn_checksum_is_validated():
    assert validate_inn("7707 083 893") == "7707083893"
    with pytest.raises(OgrnLookupError):
        validate_inn("1234567890")
