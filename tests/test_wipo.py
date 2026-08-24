from __future__ import annotations

from dataclasses import dataclass

import pytest

from trademark_report.wipo import WipoParseError, fetch_wipo_trademark, parse_wipo_html


URL = "https://www3.wipo.int/madrid/monitor/en/showData.jsp?ID=ROM.1753467&DES=1"
IMAGE_URL = (
    "https://www3.wipo.int/madrid/monitor/jsp/data.jsp?"
    "KEY=ROM_ACT.1753467&TYPE=jpg&qi=test"
)
HTML = f"""
<html><body>
<div id="headerStatusContainer"><table><tr><td class="markname">
  <h3>1753467- YokoSun</h3>
</td></tr></table></div>
<div id="markInformationHeader">
  <table class="markInformation"><tbody><tr>
    <td class="mark"><img src="{IMAGE_URL}" alt="YokoSun"></td>
    <td class="name"><div>Obschestvo s ogranichennoi otvetstvennostyu
      &quot;Aziya Layf&quot;</div></td>
    <td class="date"><div>02.06.2023</div></td>
    <td class="date"><div>02.06.2033</div></td>
    <td class="nice"><div>03, 05, 16</div></td>
  </tr></tbody></table>
</div>
</body></html>
"""


def test_parse_wipo_record() -> None:
    record = parse_wipo_html(HTML, URL)

    assert record.database == "ROM"
    assert record.registration_number == "1753467"
    assert record.mark_name == "YokoSun"
    assert record.registration_date == "02.06.2023"
    assert record.owner == 'Obschestvo s ogranichennoi otvetstvennostyu "Aziya Layf"'
    assert record.nice_class_numbers == ["03", "05", "16"]
    assert record.image_url == IMAGE_URL


@dataclass
class FakeResponse:
    content: bytes
    url: str
    content_type: str

    @property
    def headers(self) -> dict[str, str]:
        return {"Content-Type": self.content_type}

    def raise_for_status(self) -> None:
        return None


class FakeSession:
    def __init__(self) -> None:
        self.urls: list[str] = []

    def get(self, url: str, **_: object) -> FakeResponse:
        self.urls.append(url)
        if len(self.urls) == 1:
            return FakeResponse(HTML.encode(), URL, "text/html")
        return FakeResponse(b"jpeg-data", IMAGE_URL, "image/jpeg")


def test_fetch_wipo_record_and_image() -> None:
    session = FakeSession()
    record = fetch_wipo_trademark(URL, session=session)

    assert session.urls == [URL, IMAGE_URL]
    assert record.image_bytes == b"jpeg-data"
    assert record.image_content_type == "image/jpeg"


@pytest.mark.parametrize(
    "url",
    [
        "https://example.com/madrid/monitor/en/showData.jsp?ID=ROM.1753467",
        "https://www3.wipo.int/madrid/monitor/en/showData.jsp?ID=BAD.1753467",
        "http://www3.wipo.int/madrid/monitor/en/showData.jsp?ID=ROM.1753467",
    ],
)
def test_rejects_unsupported_url(url: str) -> None:
    with pytest.raises(WipoParseError):
        parse_wipo_html(HTML, url)
