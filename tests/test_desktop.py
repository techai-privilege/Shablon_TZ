import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QPushButton, QScrollArea

from app import MainWindow, RecordCard
from trademark_report.fips import NiceClass, TrademarkRecord


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
