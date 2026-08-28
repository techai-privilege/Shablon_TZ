import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton, QScrollArea

from app import MainWindow, RecordCard
from trademark_report.fips import NiceClass, TrademarkRecord
from trademark_report.ogrn import FnsRegistrationData
from trademark_report.software_models import SoftwareAuthor, SoftwareConsentData


def test_native_window_has_expected_tabs_and_no_web_server():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()

    assert [window.tabs.tabText(index) for index in range(window.tabs.count())] == [
        "Основные данные",
        "Заключение",
        "Сходные обозначения",
    ]
    assert [window.performer.itemText(index) for index in range(window.performer.count())] == [
        "Алина",
        "Маша",
        "Лера",
    ]
    assert window.centralWidget().objectName() == "appBackground"
    assert [window.navigation.item(index).text() for index in range(window.navigation.count())] == [
        "Отчёт по товарному знаку",
        "Программа для ЭВМ",
    ]
    assert window.pages.count() == 2
    assert not window.centralWidget()._background.isNull()
    assert not hasattr(window, "conclusion_additions")
    assert not hasattr(window, "excess_items")
    assert window.conclusion_preview.toPlainText().strip()
    assert any(
        button.text() == "Сформировать и сохранить DOCX"
        for button in window.findChildren(QPushButton)
    )
    window.russian.add_card()
    assert len(window.russian.cards) == 1

    window.close()
    application.processEvents()


def test_software_section_has_editable_preview_and_author_cards():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    data = SoftwareConsentData(
        program_name="Тестовая программа",
        applicant_name="ООО «Тест»",
        applicant_address="Россия, г. Москва",
        inn="7707083893",
        ogrn="1027700132195",
        authors=[SoftwareAuthor(full_name="Иванов Иван Иванович")],
    )
    window.software_page.set_data(data)
    window.navigation.setCurrentRow(1)
    application.processEvents()

    assert window.pages.currentWidget() is window.software_page
    assert len(window.software_page.author_editors) == 1
    assert window.software_page.author_tabs.count() == 1
    assert window.software_page.author_tabs.tabText(0) == "Автор 1"
    assert "Иванов Иван Иванович" in window.software_page.preview.toPlainText()
    window.software_page.author_editors[0].full_name.setText("Петров Петр Петрович")
    application.processEvents()
    assert "Петров Петр Петрович" in window.software_page.preview.toPlainText()

    window.close()
    application.processEvents()


def test_loading_multiple_authors_replaces_initial_empty_tab():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    data = SoftwareConsentData(
        authors=[
            SoftwareAuthor(full_name="Автор Один"),
            SoftwareAuthor(full_name="Автор Два"),
            SoftwareAuthor(full_name="Автор Три"),
        ]
    )

    window.software_page.set_data(data)
    application.processEvents()

    assert len(window.software_page.author_editors) == 3
    assert window.software_page.author_tabs.count() == 3
    assert [
        window.software_page.author_tabs.tabText(index)
        for index in range(window.software_page.author_tabs.count())
    ] == ["Автор 1", "Автор 2", "Автор 3"]
    assert all(editor.full_name.text() for editor in window.software_page.author_editors)

    window.close()
    application.processEvents()


def test_validation_does_not_require_passport_data_for_anonymous_authors():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    page.set_data(
        SoftwareConsentData(
            program_name="Тестовая программа",
            applicant_name="ООО «Тест»",
            applicant_address="Россия, г. Москва",
            inn="7707083893",
            ogrn="1027700132195",
            authors_will_be_mentioned=False,
            authors=[
                SoftwareAuthor(
                    full_name="Иванов Иван Иванович",
                    creative_contribution="Разработка алгоритма",
                )
            ],
        )
    )

    assert page.data().authors_will_be_mentioned is False
    assert page.validation_errors() == []
    assert "☐ упоминать его под своим именем" in page.preview.toPlainText()
    assert "☒ не упоминать его (анонимно)" in page.preview.toPlainText()

    window.close()
    application.processEvents()


def test_fns_result_fills_ogrn_and_applicant_address():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    page.inn.setText("7707083893")
    page._ogrn_lookup_inn = "7707083893"
    page._show_warnings([
        "В разделе I не найдено наименование заявителя.",
        "В разделе I не найден адрес заявителя — его нужно заполнить вручную.",
    ])

    page._ogrn_received(
        FnsRegistrationData(
            ogrn="1027700132195",
            address="117312, г. Москва",
            name="ООО «Тест»",
        )
    )

    assert page.ogrn.text() == "1027700132195"
    assert page.applicant_name.text() == "ООО «Тест»"
    assert page.applicant_address.toPlainText() == "117312, г. Москва"
    assert not page.warning_label.isVisible()

    window.close()
    application.processEvents()


def test_fns_result_does_not_overwrite_existing_applicant_name():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    page.inn.setText("7707083893")
    page.applicant_name.setText("Название из анкеты")
    page._ogrn_lookup_inn = "7707083893"

    page._ogrn_received(
        FnsRegistrationData(ogrn="1027700132195", name="Название из ФНС")
    )

    assert page.applicant_name.text() == "Название из анкеты"

    window.close()
    application.processEvents()


def test_edited_conclusion_is_passed_to_report_data():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.conclusion_preview.setPlainText("Исправленный вручную текст")

    report = window._report()

    assert report.conclusion_content is not None
    assert report.conclusion_content[0].runs[0].text == "Исправленный вручную текст"
    window.close()
    application.processEvents()


def test_window_adapts_to_compact_laptop_size():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.resize(800, 600)
    window.show()
    application.processEvents()

    margins = window.root_layout.contentsMargins()
    assert window.minimumWidth() <= 800
    assert window.minimumHeight() <= 600
    assert margins.left() == 10
    assert isinstance(window.tabs.widget(1), QScrollArea)
    assert window.conclusion_preview.minimumHeight() == 170

    window.close()
    application.processEvents()


def test_international_card_uses_wipo_and_applies_record():
    application = QApplication.instance() or QApplication([])
    card = RecordCard("international", 1)
    record = TrademarkRecord(
        source_url=(
            "https://www3.wipo.int/madrid/monitor/en/showData.jsp?"
            "ID=ROM.1753467&DES=1"
        ),
        database="ROM",
        mark_name="YokoSun",
        registration_number="1753467",
        registration_date="02.06.2023",
        owner='Obschestvo s ogranichennoi otvetstvennostyu "Aziya Layf"',
        nice_classes=[NiceClass("03"), NiceClass("05"), NiceClass("16")],
    )

    assert card.source_name == "WIPO"
    assert card.fetch_button.text() == "Заполнить из WIPO"
    card._apply_source(record)
    assert card.display_name.text() == "YokoSun"
    assert card.number.text() == "1753467"
    assert card.relevant_date.text() == "02.06.2023"
    assert card.classes.text() == "03, 05, 16"

    card.close()
    application.processEvents()
