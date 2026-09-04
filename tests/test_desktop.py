import os
import time
from typing import Any, cast

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QStandardPaths
from PySide6.QtWidgets import QApplication, QMessageBox, QPushButton, QScrollArea

from app import BackgroundWidget, MainWindow, RecordCard
from trademark_report.fips import NiceClass, TrademarkRecord
from trademark_report.ogrn import FnsRegistrationData
from trademark_report.software_models import (
    APPLICANT_INDIVIDUAL,
    SoftwareAuthor,
    SoftwareConsentData,
)


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
    central = window.centralWidget()
    assert isinstance(central, BackgroundWidget)
    assert central.objectName() == "appBackground"
    assert [window.navigation.item(index).text() for index in range(window.navigation.count())] == [
        "Отчёт по товарному знаку",
        "Программа для ЭВМ",
    ]
    assert window.pages.count() == 2
    assert not central._background.isNull()
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


def test_individual_applicant_does_not_require_or_preview_ogrn():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    page.set_data(
        SoftwareConsentData(
            applicant_type=APPLICANT_INDIVIDUAL,
            program_name="Заноза",
            applicant_name="Анфиногенов Семен Васильевич",
            applicant_address="Россия, г. Екатеринбург",
            inn="667803703833",
            authors_will_be_mentioned=False,
            authors=[
                SoftwareAuthor(
                    full_name="Старков Алексей Николаевич",
                    creative_contribution="Разработка всей программы",
                )
            ],
        )
    )
    application.processEvents()

    assert page.data().applicant_type == APPLICANT_INDIVIDUAL
    assert page.data().missing_common_fields() == []
    assert page.validation_errors() == []
    assert page.ogrn_container.isHidden()
    preview = page.preview.toPlainText()
    assert "ИНН: 667803703833" in preview
    assert "ОГРН:" not in preview
    assert "125993" in preview

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


def test_fns_result_replaces_short_applicant_name_with_full_name():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    page.inn.setText("7707083893")
    page.applicant_name.setText("Название из анкеты")
    page._ogrn_lookup_inn = "7707083893"

    page._ogrn_received(
        FnsRegistrationData(ogrn="1027700132195", name="Название из ФНС")
    )

    assert page.applicant_name.text() == "Название из ФНС"

    window.close()
    application.processEvents()


def test_fns_result_does_not_replace_full_applicant_name_with_its_fragment():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    page = window.software_page
    full_name = (
        "АКЦИОНЕРНОЕ ОБЩЕСТВО «ИНСТИТУТ ОБОРУДОВАНИЯ "
        "НЕФТЕПЕРЕРАБАТЫВАЮЩЕЙ ПРОМЫШЛЕННОСТИ»"
    )
    page.inn.setText("7707083893")
    page.applicant_name.setText(full_name)
    page._ogrn_lookup_inn = "7707083893"

    page._ogrn_received(
        FnsRegistrationData(
            ogrn="1027700132195",
            name='НЕФТЕПЕРЕРАБАТЫВАЮЩЕЙ ПРОМЫШЛЕННОСТИ"',
        )
    )

    assert page.applicant_name.text() == full_name
    assert page.applicant_name.cursorPosition() == 0
    assert page.applicant_name.toolTip() == full_name

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


def test_busy_record_card_cannot_be_removed_until_worker_finishes():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.russian.add_card()
    card = window.russian.cards[0]

    class BusyThread:
        def isRunning(self):
            return True

    cast(Any, card)._thread = BusyThread()
    window.russian.remove_card(card)

    assert window.russian.cards == [card]
    card._thread = None
    window.close()
    application.processEvents()


def test_incomplete_similar_record_is_rejected_before_report_generation():
    application = QApplication.instance() or QApplication([])
    window = MainWindow()
    window.russian.add_card()
    card = window.russian.cards[0]
    card.url.setText("https://new.fips.ru/registers-doc-view/fips_servlet?DB=RUTM")

    validation_error = window.russian.validation_error()
    assert validation_error is not None
    assert "заполните" in validation_error

    window.close()
    application.processEvents()


def test_consent_documents_are_saved_without_blocking_the_ui(monkeypatch, tmp_path):
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
            authors=[
                SoftwareAuthor(
                    full_name="Иванов Иван Иванович",
                    birth_date="01.01.1980",
                    address="Россия, г. Москва",
                    passport_series="4510",
                    passport_number="123456",
                    passport_issue_date="02.02.2020",
                    passport_issuer="ГУ МВД России по г. Москве",
                    creative_contribution="Разработка алгоритма",
                )
            ],
        )
    )
    monkeypatch.setattr(QStandardPaths, "writableLocation", lambda *_: str(tmp_path))
    monkeypatch.setattr(QMessageBox, "information", lambda *args, **kwargs: None)
    monkeypatch.setattr(QMessageBox, "critical", lambda *args, **kwargs: None)

    page.save_document()

    assert not page.save_button.isEnabled()
    deadline = time.monotonic() + 10
    while page._save_thread is not None and time.monotonic() < deadline:
        application.processEvents()
        time.sleep(0.01)
    application.processEvents()
    assert page._save_thread is None
    assert page.save_button.isEnabled()
    assert (tmp_path / "Согласие Иванов Иван Иванович.docx").is_file()

    window.close()
    application.processEvents()
