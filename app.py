"""Native desktop interface for trademark report generation."""

from __future__ import annotations

from datetime import date
import json
from pathlib import Path
import re
import sys

from PySide6.QtCore import QDate, QObject, QPoint, QRect, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import (
    QColor,
    QFont,
    QPainter,
    QPalette,
    QPixmap,
    QTextCharFormat,
    QTextCursor,
    QTextListFormat,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from trademark_report.document import save_report
from trademark_report.fips import FipsParseError, TrademarkRecord, fetch_trademark
from trademark_report.wipo import WipoParseError, fetch_wipo_trademark
from trademark_report.models import (
    CONCLUSION_VALUES,
    ConclusionParagraph,
    ConclusionRun,
    PERFORMERS,
    PROBABILITY_VALUES,
    RELATIVE_OPTIONS,
    ProbabilityEntry,
    ReportData,
    ReportNiceClass,
    SimilarRecord,
)
from trademark_report.templates import conclusion_paragraphs
from trademark_report.software_widget import SoftwareConsentWidget


ROOT = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def _load_classes() -> dict[str, str]:
    return json.loads((ROOT / "data" / "nice_classes.json").read_text(encoding="utf-8"))


def _scrollable(widget: QWidget) -> QScrollArea:
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.Shape.NoFrame)
    area.setWidget(widget)
    return area


class BackgroundWidget(QWidget):
    """Window surface that draws a subdued, aspect-filled background image."""

    def __init__(self, image_path: Path):
        super().__init__()
        self.setObjectName("appBackground")
        self._background = QPixmap(str(image_path))

    def paintEvent(self, event) -> None:  # noqa: N802 - Qt naming convention
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        if not self._background.isNull():
            scaled = self._background.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation,
            )
            left = max(0, (scaled.width() - self.width()) // 2)
            top = max(0, (scaled.height() - self.height()) // 2)
            source = QRect(left, top, self.width(), self.height())
            painter.drawPixmap(self.rect(), scaled, source)
        else:
            painter.fillRect(self.rect(), QColor("#f7eef3"))

        # The source remains recognizable, while text and form controls keep
        # visual priority. Change the alpha (fourth value) to tune intensity.
        painter.fillRect(self.rect(), QColor(250, 247, 249, 190))


class FetchWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, url: str, kind: str):
        super().__init__()
        self.url = url
        self.kind = kind

    @Slot()
    def run(self) -> None:
        try:
            if self.kind == "international":
                self.finished.emit(fetch_wipo_trademark(self.url))
            else:
                self.finished.emit(fetch_trademark(self.url))
        except (FipsParseError, WipoParseError) as exc:
            self.failed.emit(str(exc))
        except Exception as exc:  # pragma: no cover - defensive GUI boundary
            self.failed.emit(f"Непредвиденная ошибка: {exc}")


class RecordCard(QGroupBox):
    remove_requested = Signal(object)
    message = Signal(str)

    def __init__(self, kind: str, ordinal: int):
        super().__init__()
        self.kind = kind
        self.ordinal = ordinal
        self.image_bytes: bytes | None = None
        self._thread: QThread | None = None
        self._worker: FetchWorker | None = None
        self._build_ui()
        self._refresh_title()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        source_row = QHBoxLayout()
        self.url = QLineEdit()
        if self.kind == "international":
            self.source_name = "WIPO"
            self.url.setPlaceholderText(
                "https://www3.wipo.int/madrid/monitor/en/showData.jsp?ID=ROM.…&DES=1"
            )
            source_label = "Ссылка Madrid Monitor:"
        else:
            self.source_name = "ФИПС"
            self.url.setPlaceholderText("https://new.fips.ru/registers-doc-view/...")
            source_label = "Ссылка ФИПС:"
        self.fetch_button = QPushButton(f"Заполнить из {self.source_name}")
        self.fetch_button.clicked.connect(self._fetch)
        source_row.addWidget(QLabel(source_label))
        source_row.addWidget(self.url, 1)
        source_row.addWidget(self.fetch_button)
        layout.addLayout(source_row)

        form = QFormLayout()
        self.display_name = QLineEdit()
        self.display_name.setPlaceholderText("Например, FELICE")
        self.number = QLineEdit()
        self.status = QLineEdit()
        self.relevant_date = QLineEdit()
        self.relevant_date.setPlaceholderText("ДД.ММ.ГГГГ")
        self.owner = QTextEdit()
        self.owner.setMaximumHeight(72)
        self.classes = QLineEdit()
        self.classes.setPlaceholderText("Например, 03, 35")
        self.unprotected = QLineEdit()

        form.addRow("Обозначение сходного знака:", self.display_name)
        form.addRow("Номер заявки:" if self.kind == "application" else "Номер регистрации:", self.number)
        if self.kind == "application":
            form.addRow("Статус:", self.status)
            date_label = "Дата подачи:"
            owner_label = "Заявитель:"
        elif self.kind == "international":
            date_label = "Дата регистрации:"
            owner_label = "Правообладатель:"
        else:
            date_label = "Дата приоритета:"
            owner_label = "Правообладатель:"
        form.addRow(date_label, self.relevant_date)
        form.addRow(owner_label, self.owner)
        form.addRow("Однородные классы МКТУ:", self.classes)
        if self.kind == "russian":
            form.addRow("Неохраняемый элемент:", self.unprotected)
        layout.addLayout(form)

        image_row = QHBoxLayout()
        self.image = QLabel("Изображение не выбрано")
        self.image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image.setMinimumSize(220, 110)
        self.image.setStyleSheet(
            "border: 1px solid #cbd5e1; border-radius: 6px; "
            "background: #ffffff; color: #475569;"
        )
        image_actions = QVBoxLayout()
        choose = QPushButton("Выбрать изображение")
        choose.clicked.connect(self._choose_image)
        remove = QPushButton("Удалить карточку")
        remove.setProperty("secondary", True)
        remove.clicked.connect(lambda: self.remove_requested.emit(self))
        image_actions.addWidget(choose)
        image_actions.addWidget(remove)
        image_actions.addStretch()
        image_row.addWidget(self.image, 1)
        image_row.addLayout(image_actions)
        layout.addLayout(image_row)

        self.number.textChanged.connect(self._refresh_title)

    def _refresh_title(self) -> None:
        number = self.number.text().strip() or str(self.ordinal)
        self.setTitle(f"Карточка {number}")

    def set_ordinal(self, ordinal: int) -> None:
        self.ordinal = ordinal
        self._refresh_title()

    def _choose_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите изображение товарного знака",
            "",
            "Изображения (*.png *.jpg *.jpeg)",
        )
        if path:
            self.image_bytes = Path(path).read_bytes()
            self._show_image()

    def _show_image(self) -> None:
        pixmap = QPixmap()
        if self.image_bytes and pixmap.loadFromData(self.image_bytes):
            self.image.setPixmap(
                pixmap.scaled(
                    320,
                    150,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
            )
        else:
            self.image.setText("Изображение отсутствует")

    def _fetch(self) -> None:
        url = self.url.text().strip()
        if not url:
            QMessageBox.warning(
                self, "Нет ссылки", f"Вставьте ссылку на карточку {self.source_name}."
            )
            return
        self.fetch_button.setEnabled(False)
        self.fetch_button.setText("Загрузка…")
        self.message.emit(f"Загружаю карточку {self.source_name}…")

        self._thread = QThread(self)
        self._worker = FetchWorker(url, self.kind)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._apply_source)
        self._worker.failed.connect(self._fetch_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    @Slot(object)
    def _apply_source(self, source: TrademarkRecord) -> None:
        self.url.setText(source.source_url)
        if source.mark_name:
            self.display_name.setText(source.mark_name)
        self.number.setText(
            (source.application_number if self.kind == "application" else source.registration_number)
            or ""
        )
        if self.kind == "application":
            relevant_date = source.filing_date or source.priority_date
            owner = source.applicant or source.owner
            self.status.setText(source.status or "")
        elif self.kind == "international":
            relevant_date = source.registration_date or source.priority_date
            owner = source.owner or source.applicant
        else:
            relevant_date = source.priority_date or source.registration_date
            owner = source.owner or source.applicant
        self.relevant_date.setText(relevant_date or "")
        self.owner.setPlainText(owner or "")
        self.classes.setText(", ".join(source.nice_class_numbers))
        self.unprotected.setText(source.unprotected_elements or "")
        self.image_bytes = source.image_bytes
        self._show_image()
        self._finish_fetch(
            f"Данные {self.source_name} получены. Проверьте однородные классы МКТУ."
        )

    @Slot(str)
    def _fetch_failed(self, message: str) -> None:
        self._finish_fetch(f"Не удалось загрузить карточку {self.source_name}.")
        QMessageBox.critical(self, f"Ошибка {self.source_name}", message)

    def _finish_fetch(self, message: str) -> None:
        self.fetch_button.setEnabled(True)
        self.fetch_button.setText(f"Заполнить из {self.source_name}")
        self.message.emit(message)

    def value(self) -> SimilarRecord:
        return SimilarRecord(
            kind=self.kind,
            source_url=self.url.text().strip(),
            display_name=self.display_name.text().strip(),
            number=self.number.text().strip(),
            status=self.status.text().strip(),
            relevant_date=self.relevant_date.text().strip(),
            owner_or_applicant=self.owner.toPlainText().strip(),
            related_classes=self.classes.text().strip(),
            unprotected_element=self.unprotected.text().strip(),
            image_bytes=self.image_bytes,
        )


class RecordList(QWidget):
    message = Signal(str)

    def __init__(self, title: str, kind: str):
        super().__init__()
        self.kind = kind
        self.cards: list[RecordCard] = []
        layout = QVBoxLayout(self)
        heading = QLabel(title)
        heading.setProperty("heading", True)
        layout.addWidget(heading)

        add = QPushButton("Добавить карточку")
        add.clicked.connect(self.add_card)
        layout.addWidget(add, 0, Qt.AlignmentFlag.AlignLeft)

        self.cards_widget = QWidget()
        self.cards_layout = QVBoxLayout(self.cards_widget)
        self.cards_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.empty_hint = QLabel("Карточек пока нет. Нажмите «Добавить карточку».")
        self.empty_hint.setProperty("emptyHint", True)
        self.cards_layout.addWidget(self.empty_hint)
        layout.addWidget(_scrollable(self.cards_widget), 1)

    def add_card(self) -> None:
        self.empty_hint.hide()
        card = RecordCard(self.kind, len(self.cards) + 1)
        card.remove_requested.connect(self.remove_card)
        card.message.connect(self.message)
        self.cards.append(card)
        self.cards_layout.addWidget(card)

    @Slot(object)
    def remove_card(self, card: RecordCard) -> None:
        self.cards.remove(card)
        card.deleteLater()
        for index, item in enumerate(self.cards, 1):
            item.set_ordinal(index)
        self.empty_hint.setVisible(not self.cards)

    def values(self) -> list[SimilarRecord]:
        result = []
        for card in self.cards:
            value = card.value()
            if value.number or value.source_url or value.display_name:
                result.append(value)
        return result


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.classes_directory = _load_classes()
        self.setWindowTitle("Документы по интеллектуальной собственности")
        self.setMinimumSize(760, 540)
        screen = QApplication.screenAt(QPoint(0, 0)) or QApplication.primaryScreen()
        if screen is not None:
            available = screen.availableGeometry()
            self.resize(
                min(1240, max(760, int(available.width() * 0.92))),
                min(860, max(540, int(available.height() * 0.90))),
            )
        else:
            self.resize(1020, 700)
        self._build_ui()
        self._apply_style()
        self._refresh_conclusion_preview()

    def _build_ui(self) -> None:
        root = BackgroundWidget(ROOT / "assets" / "app_background.jpg")
        self.root_layout = QHBoxLayout(root)
        self.root_layout.setContentsMargins(20, 18, 20, 16)
        self.root_layout.setSpacing(10)

        self.navigation = QListWidget()
        self.navigation.setObjectName("mainNavigation")
        self.navigation.setFixedWidth(220)
        self.navigation.setWordWrap(True)
        self.navigation.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.navigation.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.navigation.addItems(["Отчёт по товарному знаку", "Программа для ЭВМ"])
        self.root_layout.addWidget(self.navigation)

        self.pages = QStackedWidget()
        trademark_page = QWidget()
        trademark_layout = QVBoxLayout(trademark_page)
        trademark_layout.setContentsMargins(0, 0, 0, 0)
        self.title_label = QLabel("Отчет о проверке товарного знака")
        self.title_label.setProperty("title", True)
        self.subtitle_label = QLabel(
            "Данные ФИПС и WIPO заполняются автоматически, экспертные выводы остаются под вашим контролем."
        )
        self.subtitle_label.setProperty("subtitle", True)
        trademark_layout.addWidget(self.title_label)
        trademark_layout.addWidget(self.subtitle_label)

        self.tabs = QTabWidget()
        self.tabs.tabBar().setUsesScrollButtons(True)
        self.tabs.tabBar().setElideMode(Qt.TextElideMode.ElideNone)
        self.tabs.addTab(self._main_tab(), "Основные данные")
        self.tabs.addTab(self._conclusion_tab(), "Заключение")
        self.tabs.addTab(self._records_tab(), "Сходные обозначения")
        self.tabs.currentChanged.connect(self._tab_changed)
        trademark_layout.addWidget(self.tabs, 1)

        generate = QPushButton("Сформировать и сохранить DOCX")
        generate.setMinimumHeight(46)
        generate.clicked.connect(self._generate)
        trademark_layout.addWidget(generate)
        self.pages.addWidget(trademark_page)

        self.software_page = SoftwareConsentWidget(ROOT / "assets" / "software_consent_template.docx")
        self.software_page.status_message.connect(self.statusBar().showMessage)
        self.pages.addWidget(self.software_page)
        self.root_layout.addWidget(self.pages, 1)
        self.navigation.currentRowChanged.connect(self.pages.setCurrentIndex)
        self.navigation.currentRowChanged.connect(self._main_section_changed)
        self.navigation.setCurrentRow(0)
        self.setCentralWidget(root)
        self.statusBar().showMessage("Готово")

    def _main_tab(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)

        report_box = QGroupBox("Данные отчета")
        form = QFormLayout(report_box)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        self.designation = QLineEdit()
        self.search_queries = QTextEdit()
        self.search_queries.setMaximumHeight(80)
        self.business_area = QTextEdit()
        self.business_area.setMaximumHeight(80)
        self.report_date = QDateEdit(QDate.currentDate())
        self.report_date.setCalendarPopup(True)
        self.report_date.setDisplayFormat("dd.MM.yyyy")
        form.addRow("Обозначение *:", self.designation)
        form.addRow("Дата отчета:", self.report_date)
        form.addRow("Поисковые запросы *:", self.search_queries)
        form.addRow("Сфера деятельности:", self.business_area)
        layout.addWidget(report_box)

        classes_box = QGroupBox("Классы МКТУ")
        classes_layout = QVBoxLayout(classes_box)
        add_row = QHBoxLayout()
        self.class_selector = QComboBox()
        for number in self.classes_directory:
            self.class_selector.addItem(f"Класс {number}", number)
        add_class = QPushButton("Добавить класс")
        add_class.clicked.connect(self._add_class)
        remove_class = QPushButton("Удалить выбранный")
        remove_class.setProperty("secondary", True)
        remove_class.clicked.connect(self._remove_classes)
        add_row.addWidget(self.class_selector)
        add_row.addWidget(add_class)
        add_row.addWidget(remove_class)
        add_row.addStretch()
        classes_layout.addLayout(add_row)
        self.class_table = QTableWidget(0, 2)
        self.class_table.setHorizontalHeaderLabels(["Класс", "Описание товаров и услуг"])
        self.class_table.verticalHeader().setVisible(False)
        self.class_table.setColumnWidth(0, 95)
        self.class_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.class_table.setAlternatingRowColors(True)
        self.class_table.horizontalHeader().setStretchLastSection(True)
        self.class_table.setMinimumHeight(180)
        classes_layout.addWidget(self.class_table)
        layout.addWidget(classes_box)

        dates_box = QGroupBox("Актуальность баз")
        dates_form = QFormLayout(dates_box)
        dates_form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        self.tm_database_date = QLineEdit()
        self.tm_database_date.setPlaceholderText("ДД.ММ.ГГГГ")
        self.app_database_date = QLineEdit()
        self.app_database_date.setPlaceholderText("ДД.ММ.ГГГГ")
        dates_form.addRow("Дата обновления базы товарных знаков *:", self.tm_database_date)
        dates_form.addRow("Дата обновления базы заявок *:", self.app_database_date)
        layout.addWidget(dates_box)

        absolute_box = QGroupBox("Абсолютные основания для отказа")
        absolute_layout = QVBoxLayout(absolute_box)
        self.absolute_checkbox = QCheckBox("Абсолютные основания имеются")
        self.absolute_text = QTextEdit()
        self.absolute_text.setPlaceholderText("Формулировка абсолютных оснований")
        self.absolute_text.setMaximumHeight(90)
        self.absolute_text.setEnabled(False)
        self.absolute_checkbox.toggled.connect(self.absolute_text.setEnabled)
        absolute_layout.addWidget(self.absolute_checkbox)
        absolute_layout.addWidget(self.absolute_text)
        layout.addWidget(absolute_box)

        relative_box = QGroupBox("Относительные препятствия для регистрации")
        relative_layout = QVBoxLayout(relative_box)
        self.relative_checks: list[QCheckBox] = []
        for option in RELATIVE_OPTIONS:
            checkbox = QCheckBox(option)
            checkbox.toggled.connect(lambda checked, current=checkbox: self._relative_changed(current, checked))
            self.relative_checks.append(checkbox)
            relative_layout.addWidget(checkbox)
        self.relative_checks[0].setChecked(True)
        layout.addWidget(relative_box)
        layout.addStretch()
        return _scrollable(content)

    def _conclusion_tab(self) -> QWidget:
        content = QWidget()
        layout = QVBoxLayout(content)

        box = QGroupBox("Экспертное заключение")
        form = QFormLayout(box)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        self.conclusion = QComboBox()
        self.conclusion.addItems(CONCLUSION_VALUES)
        self.conclusion.currentTextChanged.connect(self._conclusion_template_changed)
        self.performer = QComboBox()
        self.performer.addItems(PERFORMERS)
        form.addRow("Шаблон заключения *:", self.conclusion)
        form.addRow("Кто готовил отчет *:", self.performer)
        layout.addWidget(box)

        probability_box = QGroupBox("Вероятность регистрации")
        probability_layout = QVBoxLayout(probability_box)
        probability_actions = QHBoxLayout()
        add = QPushButton("Добавить строку")
        add.clicked.connect(self._add_probability)
        remove = QPushButton("Удалить выбранную")
        remove.setProperty("secondary", True)
        remove.clicked.connect(self._remove_probabilities)
        probability_actions.addWidget(add)
        probability_actions.addWidget(remove)
        probability_actions.addStretch()
        probability_layout.addLayout(probability_actions)
        self.probability_table = QTableWidget(0, 2)
        self.probability_table.setHorizontalHeaderLabels(["Уточнение", "Вероятность"])
        self.probability_table.verticalHeader().setVisible(False)
        self.probability_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.probability_table.setAlternatingRowColors(True)
        self.probability_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.probability_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Fixed
        )
        self.probability_table.setColumnWidth(1, 290)
        probability_layout.addWidget(self.probability_table)

        preview_box = QGroupBox("Редактируемый текст заключения")
        preview_layout = QVBoxLayout(preview_box)
        hint = QLabel(
            "Текст ниже попадёт в DOCX. При необходимости его можно исправить вручную."
        )
        hint.setProperty("subtitle", True)
        preview_layout.addWidget(hint)
        preview_actions = QHBoxLayout()
        refresh = QPushButton("Сформировать заново по шаблону")
        refresh.clicked.connect(self._confirm_refresh_conclusion)
        refresh.setProperty("secondary", True)
        bold = QPushButton("Жирный")
        bold.clicked.connect(self._toggle_conclusion_bold)
        italic = QPushButton("Курсив")
        italic.clicked.connect(self._toggle_conclusion_italic)
        preview_actions.addWidget(refresh)
        preview_actions.addWidget(bold)
        preview_actions.addWidget(italic)
        preview_actions.addStretch()
        preview_layout.addLayout(preview_actions)
        self.conclusion_preview = QTextEdit()
        self.conclusion_preview.setAcceptRichText(True)
        self.conclusion_preview.setMinimumHeight(210)
        self.conclusion_preview.setPlaceholderText("Текст заключения")
        self.conclusion_preview.document().modificationChanged.connect(
            self._conclusion_modification_changed
        )
        preview_layout.addWidget(self.conclusion_preview)
        layout.addWidget(preview_box, 2)
        layout.addWidget(probability_box, 1)
        self._add_probability()
        return _scrollable(content)

    def resizeEvent(self, event) -> None:  # noqa: N802 - Qt naming convention
        """Adapt spacing and nonessential chrome to the available window size."""
        super().resizeEvent(event)
        if not hasattr(self, "root_layout"):
            return
        compact = event.size().width() < 920 or event.size().height() < 690
        if compact:
            self.root_layout.setContentsMargins(10, 8, 10, 8)
            self.root_layout.setSpacing(6)
        else:
            self.root_layout.setContentsMargins(20, 18, 20, 16)
            self.root_layout.setSpacing(10)
        self.subtitle_label.setVisible(event.size().height() >= 620)
        self.navigation.setFixedWidth(180 if compact else 220)
        if hasattr(self, "conclusion_preview"):
            self.conclusion_preview.setMinimumHeight(170 if compact else 260)

    def _main_section_changed(self, index: int) -> None:
        if index == 1:
            self.setWindowTitle("Согласия авторов программы для ЭВМ")
        else:
            self.setWindowTitle("Отчет о проверке товарного знака")

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt naming convention
        if hasattr(self, "software_page"):
            self.software_page.stop_workers()
        super().closeEvent(event)

    def _conclusion_template_report(self) -> ReportData:
        """Build the subset of current data used by the conclusion templates."""
        selected_relative = [box.text() for box in self.relative_checks if box.isChecked()]
        return ReportData(
            designation=self.designation.text().strip(),
            search_queries=self.search_queries.toPlainText().strip(),
            relative_options=selected_relative or ["Отсутствуют"],
            conclusion=self.conclusion.currentText(),
            international_marks=self.international.values(),
            russian_marks=self.russian.values(),
            applications=self.applications.values(),
        )

    def _refresh_conclusion_preview(self) -> None:
        editor = self.conclusion_preview
        editor.blockSignals(True)
        editor.clear()
        cursor = editor.textCursor()
        for index, (role, text) in enumerate(
            conclusion_paragraphs(self._conclusion_template_report())
        ):
            if index:
                cursor.insertBlock()
                current_list = cursor.currentList()
                if role != "list" and current_list is not None:
                    current_list.remove(cursor.block())
            char_format = QTextCharFormat()
            char_format.setFontWeight(
                QFont.Weight.Bold if role in {"bold", "warning"} else QFont.Weight.Normal
            )
            char_format.setFontItalic(role == "italic")
            if role == "warning":
                char_format.setBackground(QColor("#fff200"))
            cursor.setCharFormat(char_format)
            cursor.insertText(text)
            if role == "list":
                list_format = QTextListFormat()
                list_format.setStyle(QTextListFormat.Style.ListLowerAlpha)
                list_format.setIndent(1)
                cursor.createList(list_format)
        editor.blockSignals(False)
        editor.document().setModified(False)
        self._conclusion_edited = False

    def _conclusion_modification_changed(self, modified: bool) -> None:
        if modified:
            self._conclusion_edited = True

    def _tab_changed(self, index: int) -> None:
        if self.tabs.tabText(index) == "Заключение" and not self._conclusion_edited:
            self._refresh_conclusion_preview()

    def _conclusion_template_changed(self, _value: str) -> None:
        if not getattr(self, "_conclusion_edited", False):
            self._refresh_conclusion_preview()

    def _confirm_refresh_conclusion(self) -> None:
        if getattr(self, "_conclusion_edited", False):
            answer = QMessageBox.question(
                self,
                "Заменить текст заключения?",
                "Ручные изменения будут удалены. Сформировать текст заново по шаблону?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        self._refresh_conclusion_preview()

    def _merge_conclusion_format(self, *, bold: bool | None = None, italic: bool | None = None) -> None:
        cursor = self.conclusion_preview.textCursor()
        char_format = QTextCharFormat()
        if bold is not None:
            char_format.setFontWeight(QFont.Weight.Bold if bold else QFont.Weight.Normal)
        if italic is not None:
            char_format.setFontItalic(italic)
        cursor.mergeCharFormat(char_format)
        self.conclusion_preview.mergeCurrentCharFormat(char_format)

    def _toggle_conclusion_bold(self) -> None:
        current = self.conclusion_preview.currentCharFormat().fontWeight()
        self._merge_conclusion_format(bold=current < QFont.Weight.Bold)

    def _toggle_conclusion_italic(self) -> None:
        current = self.conclusion_preview.currentCharFormat().fontItalic()
        self._merge_conclusion_format(italic=not current)

    def _conclusion_content(self) -> list[ConclusionParagraph]:
        result: list[ConclusionParagraph] = []
        block = self.conclusion_preview.document().begin()
        while block.isValid():
            runs: list[ConclusionRun] = []
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid() and fragment.text():
                    char_format = fragment.charFormat()
                    background = char_format.background()
                    runs.append(
                        ConclusionRun(
                            text=fragment.text(),
                            bold=char_format.fontWeight() >= QFont.Weight.Bold,
                            italic=char_format.fontItalic(),
                            highlighted=background.style() != Qt.BrushStyle.NoBrush,
                        )
                    )
                iterator += 1
            result.append(ConclusionParagraph(runs=runs, list_item=block.textList() is not None))
            block = block.next()
        return result

    def _records_tab(self) -> QWidget:
        tabs = QTabWidget()
        self.international = RecordList("Международные товарные знаки", "international")
        self.russian = RecordList("Товарные знаки РФ", "russian")
        self.applications = RecordList("Заявки на товарные знаки", "application")
        for widget, label in (
            (self.international, "Международные знаки"),
            (self.russian, "Товарные знаки РФ"),
            (self.applications, "Заявки"),
        ):
            widget.message.connect(self.statusBar().showMessage)
            tabs.addTab(widget, label)
        return tabs

    def _add_class(self) -> None:
        number = self.class_selector.currentData()
        for row in range(self.class_table.rowCount()):
            if self.class_table.item(row, 0).text() == number:
                self.class_table.selectRow(row)
                return
        row = self.class_table.rowCount()
        self.class_table.insertRow(row)
        number_item = QTableWidgetItem(number)
        number_item.setFlags(number_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.class_table.setItem(row, 0, number_item)
        self.class_table.setItem(row, 1, QTableWidgetItem(self.classes_directory.get(number, "")))
        self.class_table.resizeRowsToContents()

    def _remove_classes(self) -> None:
        rows = sorted({index.row() for index in self.class_table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.class_table.removeRow(row)

    def _relative_changed(self, current: QCheckBox, checked: bool) -> None:
        if not checked:
            return
        absent = self.relative_checks[0]
        if current is absent:
            for checkbox in self.relative_checks[1:]:
                checkbox.setChecked(False)
        else:
            absent.setChecked(False)

    def _add_probability(self) -> None:
        row = self.probability_table.rowCount()
        self.probability_table.insertRow(row)
        self.probability_table.setItem(row, 0, QTableWidgetItem(""))
        values = QComboBox()
        values.addItems(PROBABILITY_VALUES)
        self.probability_table.setCellWidget(row, 1, values)

    def _remove_probabilities(self) -> None:
        if self.probability_table.rowCount() <= 1:
            return
        rows = sorted({index.row() for index in self.probability_table.selectedIndexes()}, reverse=True)
        for row in rows:
            if self.probability_table.rowCount() > 1:
                self.probability_table.removeRow(row)

    def _nice_classes(self) -> list[ReportNiceClass]:
        return [
            ReportNiceClass(
                self.class_table.item(row, 0).text().strip(),
                self.class_table.item(row, 1).text().strip() if self.class_table.item(row, 1) else "",
            )
            for row in range(self.class_table.rowCount())
        ]

    def _probabilities(self) -> list[ProbabilityEntry]:
        result = []
        for row in range(self.probability_table.rowCount()):
            subject_item = self.probability_table.item(row, 0)
            value_widget = self.probability_table.cellWidget(row, 1)
            result.append(
                ProbabilityEntry(
                    value=value_widget.currentText(),
                    subject=subject_item.text().strip() if subject_item else "",
                )
            )
        return result

    def _validate(self) -> str | None:
        if not self.designation.text().strip():
            return "Заполните обозначение."
        if not self.search_queries.toPlainText().strip():
            return "Заполните поисковые запросы."
        classes = self._nice_classes()
        if not classes:
            return "Добавьте хотя бы один класс МКТУ."
        if any(not item.description for item in classes):
            return "Заполните описание каждого выбранного класса МКТУ."
        if not self.tm_database_date.text().strip() or not self.app_database_date.text().strip():
            return "Заполните даты обновления обеих баз."
        if self.absolute_checkbox.isChecked() and not self.absolute_text.toPlainText().strip():
            return "Добавьте формулировку абсолютных оснований."
        if not self.conclusion_preview.toPlainText().strip():
            return "Заключение не может быть пустым."
        return None

    def _report(self) -> ReportData:
        selected_relative = [box.text() for box in self.relative_checks if box.isChecked()]
        qdate = self.report_date.date()
        return ReportData(
            designation=self.designation.text().strip(),
            search_queries=self.search_queries.toPlainText().strip(),
            business_area=self.business_area.toPlainText().strip(),
            report_date=date(qdate.year(), qdate.month(), qdate.day()),
            nice_classes=self._nice_classes(),
            has_absolute_grounds=self.absolute_checkbox.isChecked(),
            absolute_grounds_text=self.absolute_text.toPlainText().strip(),
            relative_options=selected_relative or ["Отсутствуют"],
            trademarks_database_date=self.tm_database_date.text().strip(),
            applications_database_date=self.app_database_date.text().strip(),
            conclusion=self.conclusion.currentText(),
            conclusion_content=self._conclusion_content(),
            performer=self.performer.currentText(),
            probabilities=self._probabilities(),
            international_marks=self.international.values(),
            russian_marks=self.russian.values(),
            applications=self.applications.values(),
        )

    def _generate(self) -> None:
        error = self._validate()
        if error:
            QMessageBox.warning(self, "Не хватает данных", error)
            return
        safe_name = re.sub(
            r"[^0-9A-Za-zА-Яа-я_-]+", "_", self.designation.text().strip()
        ).strip("_") or "report"
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить отчет",
            str(Path.home() / "Documents" / f"Отчет_по_ТЗ_{safe_name}.docx"),
            "Документ Word (*.docx)",
        )
        if not path:
            return
        if not path.lower().endswith(".docx"):
            path += ".docx"
        try:
            save_report(self._report(), path)
        except Exception as exc:  # pragma: no cover - defensive GUI boundary
            QMessageBox.critical(self, "Не удалось сформировать отчет", str(exc))
            return
        self.statusBar().showMessage(f"Отчет сохранен: {path}")
        QMessageBox.information(self, "Готово", f"Отчет сохранен:\n{path}")

    def _apply_style(self) -> None:
        application = QApplication.instance()
        if application is not None:
            application.setStyle("Fusion")
            palette = QPalette()
            palette.setColor(QPalette.ColorRole.Window, QColor("#f4f6f8"))
            palette.setColor(QPalette.ColorRole.WindowText, QColor("#17212b"))
            palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#f8fafc"))
            palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#17212b"))
            palette.setColor(QPalette.ColorRole.Text, QColor("#17212b"))
            palette.setColor(QPalette.ColorRole.Button, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor("#17212b"))
            palette.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.Highlight, QColor("#e9426f"))
            palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
            palette.setColor(QPalette.ColorRole.PlaceholderText, QColor("#7b8794"))
            palette.setColor(QPalette.ColorRole.Link, QColor("#155eef"))
            application.setPalette(palette)
        self.setStyleSheet(
            """
            QMainWindow { background-color: #f7eef3; }
            QWidget { color: #17212b; font-size: 14px; }
            QWidget#appBackground { background: transparent; }
            QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget {
                background: transparent;
            }
            QLabel { background: transparent; color: #17212b; }
            QLabel[title="true"] { font-size: 26px; font-weight: 700; color: #243447; }
            QLabel[subtitle="true"] { color: #5f7183; margin-bottom: 8px; }
            QLabel[heading="true"] { font-size: 18px; font-weight: 700; margin: 6px 2px; }
            QLabel[emptyHint="true"] {
                color: #647587; font-size: 15px; padding: 18px; background: transparent;
            }
            QGroupBox {
                font-weight: 600; border: 1px solid #d6dde5; border-radius: 8px;
                margin-top: 14px; padding: 18px 12px 12px 12px;
                background: transparent; color: #17212b;
            }
            QGroupBox::title {
                subcontrol-origin: margin; subcontrol-position: top left;
                left: 12px; padding: 0 6px; color: #243447;
                background: rgba(255, 255, 255, 210);
            }
            QLineEdit, QTextEdit, QComboBox, QDateEdit, QTableWidget {
                background-color: rgba(255, 255, 255, 224); color: #17212b;
                selection-background-color: #e9426f; selection-color: #ffffff;
                border: 1px solid #c9d2dc; border-radius: 6px; padding: 6px;
            }
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus,
            QTableWidget:focus { border: 2px solid #e9426f; }
            QComboBox QAbstractItemView {
                background-color: #ffffff; color: #17212b; border: 1px solid #c9d2dc;
                selection-background-color: #e9426f; selection-color: #ffffff;
                outline: none; padding: 4px;
            }
            QComboBox::drop-down, QDateEdit::drop-down {
                border: none; width: 28px; background: #eef2f6;
            }
            QCheckBox { background: transparent; color: #17212b; spacing: 8px; padding: 3px; }
            QCheckBox::indicator { width: 17px; height: 17px; }
            QPushButton {
                background-color: #e9426f; color: #ffffff; border: none; border-radius: 6px;
                min-height: 20px; padding: 8px 14px; font-weight: 600;
            }
            QPushButton:hover { background-color: #d83763; }
            QPushButton:pressed { background-color: #bd2b54; }
            QPushButton:disabled { background-color: #aeb8c2; color: #f4f6f8; }
            QPushButton[secondary="true"] { background-color: #647587; color: #ffffff; }
            QPushButton[secondary="true"]:hover { background-color: #506173; }
            QTabWidget::pane {
                border: 1px solid #d6dde5; border-radius: 8px;
                background: transparent; top: -1px;
            }
            QTabBar::tab {
                background-color: #e8edf2; color: #334155; border: 1px solid #d6dde5;
                border-bottom: none; padding: 10px 18px; margin-right: 2px;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
            }
            QTabBar::tab:hover { background-color: #f7f9fb; }
            QTabBar::tab:selected { background-color: #ffffff; color: #d83763; font-weight: 700; }
            QHeaderView::section {
                background-color: #e8edf2; color: #243447; border: none;
                border-right: 1px solid #c9d2dc; border-bottom: 1px solid #c9d2dc;
                padding: 7px; font-weight: 600;
            }
            QTableWidget { gridline-color: #d6dde5; alternate-background-color: #f8fafc; }
            QStatusBar { background-color: #ffffff; color: #475569; border-top: 1px solid #d6dde5; }
            QListWidget#mainNavigation {
                background-color: rgba(255, 255, 255, 220); color: #243447;
                border: 1px solid #d6dde5; border-radius: 9px; padding: 6px;
                outline: none; font-weight: 600;
            }
            QListWidget#mainNavigation::item {
                min-height: 42px; padding: 8px 10px; border-radius: 6px; margin: 2px;
            }
            QListWidget#mainNavigation::item:selected {
                background-color: #e9426f; color: #ffffff;
            }
            QLabel[warning="true"] {
                background-color: rgba(255, 244, 229, 235); color: #8a4b08;
                border: 1px solid #f2c078; border-radius: 6px; padding: 8px;
            }
            QScrollBar:vertical { background: #eef2f6; width: 12px; margin: 0; }
            QScrollBar::handle:vertical { background: #a9b6c3; min-height: 28px; border-radius: 6px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
            """
        )


def _bring_window_to_front(app: QApplication, window: MainWindow) -> None:
    """Show and activate the native window, including when started from VS Code."""
    window.showNormal()
    screen = QApplication.screenAt(QPoint(0, 0)) or QApplication.primaryScreen()
    if screen is not None:
        frame = window.frameGeometry()
        frame.moveCenter(screen.availableGeometry().center())
        window.move(frame.topLeft())
    window.raise_()
    window.activateWindow()

    if sys.platform != "darwin":
        return
    try:
        from ctypes import c_bool, c_long, c_void_p, cdll

        objc = cdll.LoadLibrary("/usr/lib/libobjc.A.dylib")
        objc.objc_getClass.restype = c_void_p
        objc.sel_registerName.restype = c_void_p
        objc.objc_msgSend.restype = c_void_p
        objc.objc_msgSend.argtypes = [c_void_p, c_void_p]
        ns_application = objc.objc_getClass(b"NSApplication")
        shared_application = objc.objc_msgSend(
            ns_application, objc.sel_registerName(b"sharedApplication")
        )
        objc.objc_msgSend.restype = c_bool
        objc.objc_msgSend.argtypes = [c_void_p, c_void_p, c_long]
        objc.objc_msgSend(
            shared_application,
            objc.sel_registerName(b"setActivationPolicy:"),
            0,  # NSApplicationActivationPolicyRegular
        )
        objc.objc_msgSend.restype = None
        objc.objc_msgSend.argtypes = [c_void_p, c_void_p, c_bool]
        objc.objc_msgSend(
            shared_application,
            objc.sel_registerName(b"activateIgnoringOtherApps:"),
            True,
        )
    except (AttributeError, OSError):
        # Qt activation above remains the portable fallback.
        pass


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Документы по интеллектуальной собственности")
    window = MainWindow()
    window.show()
    QTimer.singleShot(0, lambda: _bring_window_to_front(app, window))
    return app.exec()


def _packaged_self_test() -> int:
    """Exercise bundled resources and DOCX generation without opening the GUI."""

    from tempfile import TemporaryDirectory

    from trademark_report.consent_document import save_consents
    from trademark_report.software_models import SoftwareAuthor, SoftwareConsentData

    report = ReportData(
        designation="TEST",
        search_queries="TEST",
        nice_classes=[ReportNiceClass("01", "контрольное описание")],
        trademarks_database_date="01.01.2026",
        applications_database_date="01.01.2026",
        probabilities=[ProbabilityEntry(value="ВЫСОКАЯ")],
    )
    try:
        with TemporaryDirectory() as directory:
            output = Path(directory) / "self-test.docx"
            save_report(report, output)
            consent_output = Path(directory) / "consent-self-test.docx"
            save_consents(
                SoftwareConsentData(
                    program_name="TEST",
                    applicant_name='ООО "TEST"',
                    applicant_address="TEST",
                    inn="7707083893",
                    ogrn="1027700132195",
                    authors=[
                        SoftwareAuthor(
                            full_name="Иванов Иван Иванович",
                            birth_date="01.01.1980",
                            address="TEST",
                            passport_series="4510",
                            passport_number="123456",
                            passport_issue_date="01.01.2020",
                            passport_issuer="TEST",
                            creative_contribution="TEST",
                        )
                    ],
                ),
                consent_output,
            )
            outputs = (output, consent_output)
            return 0 if all(item.is_file() and item.stat().st_size for item in outputs) else 1
    except Exception:
        return 1


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        raise SystemExit(_packaged_self_test())
    raise SystemExit(main())
