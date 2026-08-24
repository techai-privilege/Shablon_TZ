from trademark_report.fips import FipsParseError, parse_trademark_html


SOURCE_URL = (
    "https://new.fips.ru/registers-doc-view/"
    "fips_servlet?DB=RUTM&DocNumber=1227836&TypeFile=html"
)
APPLICATION_URL = (
    "https://new.fips.ru/registers-doc-view/"
    "fips_servlet?DB=RUTMAP&DocNumber=2025802265&TypeFile=html"
)


SAMPLE_HTML = """\
<html><body><div>ФЕДЕРАЛЬНАЯ СЛУЖБА ПО ИНТЕЛЛЕКТУАЛЬНОЙ СОБСТВЕННОСТИ</div>
<tr class="Status"><td>Статус: действует (последнее изменение статуса: 04.06.2026)</td></tr>
<p class="bib">(111) <i>Номер государственной регистрации:</i> <b>1227836</b></p>
<p class="bib">(210) <i>Номер заявки:</i> <b>2026708897</b></p>
<p class="bib2"><i>Приоритет:</i> <b>03.02.2026</b></p>
<p class="bib">(220) <i>Дата подачи заявки:</i> <b>03.02.2026</b></p>
<p class="bib">(151) <i>Дата государственной регистрации:</i> <b>03.06.2026</b></p>
<p class="bib">(540) <i>Изображение</i><img src="https://fips.ru/ofpstorage/TM/mark.jpg"></p>
<p class="bib">(732) <i>Правообладатель:</i><br><b>ООО «АЙЛЭНД» (RU)</b></p>
<p class="bib">(511) <i>Классы МКТУ:</i><br>
<b>35 - административные услуги.</b><br><b>39 - транспортировка.</b></p>
</body></html>
"""


APPLICATION_HTML = """\
<html><body>
<tr class="Status"><td>Состояние делопроизводства: Принято решение о регистрации
(последнее изменение: 26.06.2026)</td></tr>
<p class="bib">(210) <i>Номер заявки:</i> <b>2025802265</b></p>
<p class="bib">(200) <i>Дата поступления заявки:</i> <b>22.09.2025</b></p>
<p class="bib">(540) <i>Изображение заявляемого обозначения</i><br>
<a href="https://fips.ru/Image/RUTMAP_Images/new2025/2025802265.jpg">
<img src="https://fips.ru/Image/RUTMAP_Images/new2025/2025802265-s.jpg"></a></p>
<p class="bib">(731) <i>Заявитель:</i><br>
<b>Молчанов Дмитрий Герасимович (RU)</b></p>
<p class="bib">(511) <i>Классы МКТУ:</i><br>
<b>06 - металлические конструкции.</b><br>
<b>07 - машины и оборудование.</b></p>
</body></html>
"""


def test_parse_public_trademark_card():
    record = parse_trademark_html(SAMPLE_HTML, SOURCE_URL)

    assert record.database == "RUTM"
    assert record.registration_number == "1227836"
    assert record.application_number == "2026708897"
    assert record.registration_date == "03.06.2026"
    assert record.priority_date == "03.02.2026"
    assert record.owner == "ООО «АЙЛЭНД» (RU)"
    assert record.status == "действует"
    assert record.nice_class_numbers == ["35", "39"]
    assert record.image_url == "https://fips.ru/ofpstorage/TM/mark.jpg"


def test_parse_public_application_card():
    record = parse_trademark_html(APPLICATION_HTML, APPLICATION_URL)

    assert record.database == "RUTMAP"
    assert record.registration_number is None
    assert record.application_number == "2025802265"
    assert record.filing_date == "22.09.2025"
    assert record.applicant == "Молчанов Дмитрий Герасимович (RU)"
    assert record.status == "Принято решение о регистрации"
    assert record.nice_class_numbers == ["06", "07"]
    assert record.image_url == (
        "https://fips.ru/Image/RUTMAP_Images/new2025/2025802265.jpg"
    )


def test_rejects_non_fips_url():
    try:
        parse_trademark_html(SAMPLE_HTML, "https://example.com/card")
    except FipsParseError as exc:
        assert "fips.ru" in str(exc)
    else:
        raise AssertionError("Expected FipsParseError")
