"""Native UI for questionnaire import, verification, preview and consent export."""

from __future__ import annotations

from datetime import date
from html import escape
from pathlib import Path
import re

from PySide6.QtCore import QDate, QObject, QStandardPaths, QThread, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTabWidget,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .consent_document import DEFAULT_CONSENT_TEMPLATE, save_author_consents
from .ogrn import (
    FnsRegistrationData,
    OgrnLookupError,
    fetch_registration_data,
    prefer_full_registration_name,
)
from .questionnaire import parse_questionnaire
from .software_models import SoftwareAuthor, SoftwareConsentData


def _scrollable(widget: QWidget) -> QScrollArea:
    area = QScrollArea()
    area.setWidgetResizable(True)
    area.setFrameShape(QScrollArea.Shape.NoFrame)
    area.setWidget(widget)
    return area


class OgrnWorker(QObject):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, inn: str):
        super().__init__()
        self.inn = inn

    @Slot()
    def run(self) -> None:
        try:
            self.finished.emit(fetch_registration_data(self.inn))
        except OgrnLookupError as exc:
            self.failed.emit(str(exc))
        except Exception as exc:  # pragma: no cover - defensive GUI boundary
            self.failed.emit(f"Непредвиденная ошибка при обращении к ФНС: {exc}")


class AuthorEditor(QWidget):
    changed = Signal()

    def __init__(self, author: SoftwareAuthor | None = None):
        super().__init__()
        form = QFormLayout(self)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        self.full_name = QLineEdit()
        self.birth_date = QLineEdit()
        self.birth_date.setPlaceholderText("ДД.ММ.ГГГГ")
        self.citizenship = QLineEdit("Российская Федерация")
        self.address = QTextEdit()
        self.address.setMaximumHeight(72)
        self.passport_series = QLineEdit()
        self.passport_number = QLineEdit()
        self.passport_issue_date = QLineEdit()
        self.passport_issue_date.setPlaceholderText("ДД.ММ.ГГГГ")
        self.passport_issuer = QTextEdit()
        self.passport_issuer.setMaximumHeight(72)
        self.creative_contribution = QTextEdit()
        self.creative_contribution.setMinimumHeight(100)
        self.rights_basis = QTextEdit()
        self.rights_basis.setMaximumHeight(72)

        form.addRow("ФИО *:", self.full_name)
        form.addRow("Дата рождения *:", self.birth_date)
        form.addRow("Гражданство *:", self.citizenship)
        form.addRow("Адрес *:", self.address)
        form.addRow("Серия паспорта *:", self.passport_series)
        form.addRow("Номер паспорта *:", self.passport_number)
        form.addRow("Дата выдачи *:", self.passport_issue_date)
        form.addRow("Кем выдан *:", self.passport_issuer)
        form.addRow("Творческий вклад *:", self.creative_contribution)
        form.addRow("Основание возникновения права:", self.rights_basis)

        for widget in (
            self.full_name,
            self.birth_date,
            self.citizenship,
            self.passport_series,
            self.passport_number,
            self.passport_issue_date,
        ):
            widget.textChanged.connect(self.changed)
        for widget in (self.address, self.passport_issuer, self.creative_contribution, self.rights_basis):
            widget.textChanged.connect(self.changed)
        self.set_value(author or SoftwareAuthor())

    def set_value(self, author: SoftwareAuthor) -> None:
        self.full_name.setText(author.full_name)
        self.birth_date.setText(author.birth_date)
        self.citizenship.setText(author.citizenship or "Российская Федерация")
        self.address.setPlainText(author.address)
        self.passport_series.setText(author.passport_series)
        self.passport_number.setText(author.passport_number)
        self.passport_issue_date.setText(author.passport_issue_date)
        self.passport_issuer.setPlainText(author.passport_issuer)
        self.creative_contribution.setPlainText(author.creative_contribution)
        self.rights_basis.setPlainText(author.rights_basis)

    def value(self) -> SoftwareAuthor:
        return SoftwareAuthor(
            full_name=self.full_name.text().strip(),
            birth_date=self.birth_date.text().strip(),
            citizenship=self.citizenship.text().strip(),
            address=self.address.toPlainText().strip(),
            passport_series=re.sub(r"\s", "", self.passport_series.text()),
            passport_number=re.sub(r"\s", "", self.passport_number.text()),
            passport_issue_date=self.passport_issue_date.text().strip(),
            passport_issuer=self.passport_issuer.toPlainText().strip(),
            creative_contribution=self.creative_contribution.toPlainText().strip(),
            rights_basis=self.rights_basis.toPlainText().strip(),
        )


class SoftwareConsentWidget(QWidget):
    status_message = Signal(str)

    def __init__(self, template_path: str | Path = DEFAULT_CONSENT_TEMPLATE):
        super().__init__()
        self.template_path = Path(template_path)
        self.author_editors: list[AuthorEditor] = []
        self._ogrn_thread: QThread | None = None
        self._ogrn_worker: OgrnWorker | None = None
        self._ogrn_lookup_inn = ""
        self._warnings: list[str] = []
        self._authors_will_be_mentioned: bool | None = None
        self._questionnaire_profile_id = ""
        self._questionnaire_profile_name = ""
        self._build_ui()
        self.add_author(SoftwareAuthor())
        self._refresh_preview()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        heading = QLabel("Согласия авторов программы для ЭВМ")
        heading.setProperty("title", True)
        subtitle = QLabel(
            "Загрузите заполненную анкету. Программа обработает только раздел I, "
            "создаст карточки авторов и подготовит отдельный DOCX для каждого автора."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("subtitle", True)
        layout.addWidget(heading)
        layout.addWidget(subtitle)

        source_row = QHBoxLayout()
        self.load_button = QPushButton("Загрузить анкету DOCX")
        self.load_button.clicked.connect(self.load_questionnaire)
        self.source_label = QLabel("Анкета не загружена")
        self.source_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        source_row.addWidget(self.load_button)
        source_row.addWidget(self.source_label, 1)
        layout.addLayout(source_row)

        self.warning_label = QLabel()
        self.warning_label.setWordWrap(True)
        self.warning_label.setProperty("warning", True)
        self.warning_label.hide()
        layout.addWidget(self.warning_label)

        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        form_content = QWidget()
        form_layout = QVBoxLayout(form_content)

        common_box = QGroupBox("Общие сведения")
        common_form = QFormLayout(common_box)
        common_form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        self.program_name = QLineEdit()
        self.applicant_name = QLineEdit()
        self.applicant_address = QTextEdit()
        self.applicant_address.setMaximumHeight(72)
        self.inn = QLineEdit()
        ogrn_row = QHBoxLayout()
        self.ogrn = QLineEdit()
        self.ogrn_button = QPushButton("Получить по ИНН")
        self.ogrn_button.clicked.connect(self.lookup_ogrn)
        ogrn_row.addWidget(self.ogrn, 1)
        ogrn_row.addWidget(self.ogrn_button)
        self.document_date = QDateEdit(QDate.currentDate())
        self.document_date.setCalendarPopup(True)
        self.document_date.setDisplayFormat("dd.MM.yyyy")
        common_form.addRow("Название программы *:", self.program_name)
        common_form.addRow("Заявитель *:", self.applicant_name)
        common_form.addRow("Адрес заявителя *:", self.applicant_address)
        common_form.addRow("ИНН *:", self.inn)
        common_form.addRow("ОГРН/ОГРНИП *:", ogrn_row)
        common_form.addRow("Дата:", self.document_date)
        form_layout.addWidget(common_box)

        authors_box = QGroupBox("Авторы")
        authors_layout = QVBoxLayout(authors_box)
        actions = QHBoxLayout()
        add = QPushButton("Добавить автора")
        add.clicked.connect(lambda: self.add_author(SoftwareAuthor()))
        remove = QPushButton("Удалить автора")
        remove.setProperty("secondary", True)
        remove.clicked.connect(self.remove_current_author)
        move_left = QPushButton("←")
        move_left.setToolTip("Переместить автора раньше")
        move_left.setProperty("secondary", True)
        move_left.clicked.connect(lambda: self.move_current_author(-1))
        move_right = QPushButton("→")
        move_right.setToolTip("Переместить автора позже")
        move_right.setProperty("secondary", True)
        move_right.clicked.connect(lambda: self.move_current_author(1))
        actions.addWidget(add)
        actions.addWidget(remove)
        actions.addWidget(move_left)
        actions.addWidget(move_right)
        actions.addStretch()
        self.author_count_label = QLabel()
        actions.addWidget(self.author_count_label)
        authors_layout.addLayout(actions)
        self.author_tabs = QTabWidget()
        self.author_tabs.currentChanged.connect(self._refresh_preview)
        authors_layout.addWidget(self.author_tabs)
        form_layout.addWidget(authors_box, 1)
        self.splitter.addWidget(_scrollable(form_content))

        preview_box = QGroupBox("Предпросмотр выбранного автора")
        preview_layout = QVBoxLayout(preview_box)
        preview_hint = QLabel(
            "Исправляйте значения в полях слева — предпросмотр обновляется автоматически."
        )
        preview_hint.setWordWrap(True)
        preview_hint.setProperty("subtitle", True)
        preview_layout.addWidget(preview_hint)
        self.preview = QTextBrowser()
        self.preview.setOpenExternalLinks(False)
        preview_layout.addWidget(self.preview, 1)
        self.splitter.addWidget(preview_box)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 1)
        layout.addWidget(self.splitter, 1)

        self.save_button = QPushButton("Сформировать согласия на рабочий стол")
        self.save_button.setMinimumHeight(46)
        self.save_button.clicked.connect(self.save_document)
        layout.addWidget(self.save_button)

        for widget in (self.program_name, self.inn, self.ogrn):
            widget.textChanged.connect(self._refresh_preview)
        self.applicant_name.textChanged.connect(self._applicant_name_changed)
        self.applicant_address.textChanged.connect(self._applicant_address_changed)
        self.document_date.dateChanged.connect(self._refresh_preview)

    def resizeEvent(self, event) -> None:  # noqa: N802 - Qt convention
        super().resizeEvent(event)
        vertical = event.size().width() < 900
        self.splitter.setOrientation(Qt.Orientation.Vertical if vertical else Qt.Orientation.Horizontal)
        if vertical:
            self.splitter.setSizes([max(320, int(event.size().height() * 0.58)), 250])

    @Slot()
    def load_questionnaire(self, path: str | None = None) -> None:
        if not path:
            selected, _ = QFileDialog.getOpenFileName(
                self, "Выберите заполненную анкету", "", "Документ Word (*.docx)"
            )
            path = selected
        if not path:
            return
        try:
            result = parse_questionnaire(path)
        except Exception as exc:
            QMessageBox.critical(self, "Не удалось прочитать анкету", str(exc))
            return
        self.source_label.setText(Path(path).name)
        self.source_label.setToolTip(
            f"Профиль анкеты: {result.data.questionnaire_profile_name}"
        )
        self.set_data(result.data)
        self._show_warnings(result.warnings)
        self.status_message.emit(
            f"Анкета загружена: {Path(path).name}; "
            f"профиль: {result.data.questionnaire_profile_name}"
        )
        if result.data.inn:
            self.lookup_ogrn(silent=True)

    def set_data(self, data: SoftwareConsentData) -> None:
        self._authors_will_be_mentioned = data.authors_will_be_mentioned
        self._questionnaire_profile_id = data.questionnaire_profile_id
        self._questionnaire_profile_name = data.questionnaire_profile_name
        self.program_name.setText(data.program_name)
        self._set_applicant_name(data.applicant_name)
        self.applicant_address.setPlainText(data.applicant_address)
        self.inn.setText(data.inn)
        self.ogrn.setText(data.ogrn)
        self.document_date.setDate(
            QDate(data.document_date.year, data.document_date.month, data.document_date.day)
        )
        self.author_editors.clear()
        while self.author_tabs.count():
            page = self.author_tabs.widget(0)
            self.author_tabs.removeTab(0)
            page.deleteLater()
        for author in data.authors or [SoftwareAuthor()]:
            self.add_author(author)
        self._refresh_preview()

    def add_author(self, author: SoftwareAuthor) -> None:
        editor = AuthorEditor(author)
        editor.changed.connect(self._author_changed)
        self.author_editors.append(editor)
        self.author_tabs.addTab(_scrollable(editor), f"Автор {len(self.author_editors)}")
        self.author_tabs.setCurrentIndex(len(self.author_editors) - 1)
        self._refresh_author_titles()

    @Slot()
    def remove_current_author(self) -> None:
        if len(self.author_editors) <= 1:
            QMessageBox.warning(self, "Нельзя удалить", "В документе должен быть хотя бы один автор.")
            return
        index = self.author_tabs.currentIndex()
        page = self.author_tabs.widget(index)
        editor = self.author_editors.pop(index)
        self.author_tabs.removeTab(index)
        page.deleteLater()
        editor.deleteLater()
        self._refresh_author_titles()
        self._refresh_preview()

    def move_current_author(self, offset: int) -> None:
        index = self.author_tabs.currentIndex()
        target = index + offset
        if index < 0 or target < 0 or target >= len(self.author_editors):
            return
        page = self.author_tabs.widget(index)
        editor = self.author_editors.pop(index)
        self.author_tabs.removeTab(index)
        self.author_editors.insert(target, editor)
        self.author_tabs.insertTab(target, page, "")
        self.author_tabs.setCurrentIndex(target)
        self._refresh_author_titles()
        self._refresh_preview()

    def _refresh_author_titles(self) -> None:
        for index, editor in enumerate(self.author_editors):
            self.author_tabs.setTabText(index, f"Автор {index + 1}")
        self.author_count_label.setText(f"Количество: {len(self.author_editors)}")

    @Slot()
    def _author_changed(self) -> None:
        self._refresh_author_titles()
        self._refresh_preview()

    def _show_warnings(self, warnings: list[str]) -> None:
        self._warnings = list(warnings)
        self.warning_label.setVisible(bool(self._warnings))
        self.warning_label.setText(
            "Проверьте распознавание:\n• " + "\n• ".join(self._warnings)
            if self._warnings
            else ""
        )

    @Slot()
    def _applicant_address_changed(self) -> None:
        if self.applicant_address.toPlainText().strip():
            remaining = [
                warning
                for warning in self._warnings
                if "не найден адрес заявителя" not in warning.casefold()
            ]
            if remaining != self._warnings:
                self._show_warnings(remaining)
        self._refresh_preview()

    @Slot()
    def _applicant_name_changed(self) -> None:
        self.applicant_name.setToolTip(self.applicant_name.text().strip())
        if self.applicant_name.text().strip():
            remaining = [
                warning
                for warning in self._warnings
                if "не найдено наименование заявителя" not in warning.casefold()
            ]
            if remaining != self._warnings:
                self._show_warnings(remaining)
        self._refresh_preview()

    def _set_applicant_name(self, value: str) -> None:
        """Set a long name without leaving the one-line editor scrolled to its end."""

        self.applicant_name.setText(value)
        self.applicant_name.setCursorPosition(0)

    @Slot()
    def lookup_ogrn(self, *, silent: bool = False) -> None:
        inn = self.inn.text().strip()
        if not inn:
            if not silent:
                QMessageBox.warning(self, "Нет ИНН", "Введите или загрузите ИНН заявителя.")
            return
        if self._ogrn_thread and self._ogrn_thread.isRunning():
            return
        self.ogrn_button.setEnabled(False)
        self.ogrn_button.setText("Поиск…")
        self.status_message.emit("Получаю ОГРН и адрес по ИНН из ФНС…")
        self._ogrn_lookup_inn = re.sub(r"\D", "", inn)
        self._ogrn_thread = QThread(self)
        self._ogrn_worker = OgrnWorker(inn)
        self._ogrn_worker.moveToThread(self._ogrn_thread)
        self._ogrn_thread.started.connect(self._ogrn_worker.run)
        self._ogrn_worker.finished.connect(self._ogrn_received)
        self._ogrn_worker.failed.connect(lambda message: self._ogrn_failed(message, silent))
        self._ogrn_worker.finished.connect(self._ogrn_thread.quit)
        self._ogrn_worker.failed.connect(self._ogrn_thread.quit)
        self._ogrn_thread.finished.connect(self._ogrn_worker.deleteLater)
        self._ogrn_thread.finished.connect(self._ogrn_thread.deleteLater)
        self._ogrn_thread.finished.connect(self._ogrn_finished)
        self._ogrn_thread.start()

    @Slot()
    def _ogrn_finished(self) -> None:
        self._ogrn_worker = None
        self._ogrn_thread = None

    def stop_workers(self) -> None:
        """Finish the network worker before the application destroys Qt objects."""
        if self._ogrn_thread and self._ogrn_thread.isRunning():
            self._ogrn_thread.quit()
            self._ogrn_thread.wait(22000)

    @Slot(object)
    def _ogrn_received(self, result: FnsRegistrationData) -> None:
        if re.sub(r"\D", "", self.inn.text()) == self._ogrn_lookup_inn:
            self.ogrn.setText(result.ogrn)
            received = ["ОГРН"]
            if result.name:
                name = prefer_full_registration_name(
                    self.applicant_name.text(), result.name
                )
                self._set_applicant_name(name)
                received.append("наименование заявителя")
            if result.address:
                self.applicant_address.setPlainText(result.address)
                received.append("адрес")
                self._finish_ogrn(
                    f"Получены из ФНС: {', '.join(received)} — {result.ogrn}"
                )
            else:
                self._finish_ogrn(
                    f"Получены из ФНС: {', '.join(received)} — {result.ogrn}; "
                    "адрес не получен: "
                    f"{result.address_error or 'неизвестная ошибка'}. "
                    "Нажмите «Получить по ИНН», чтобы повторить."
                )
        else:
            self._finish_ogrn("ИНН изменился во время поиска — результат ФНС не применён.")

    def _ogrn_failed(self, message: str, silent: bool) -> None:
        self._finish_ogrn("ОГРН не получен — его можно ввести вручную.")
        if not silent:
            QMessageBox.warning(self, "Не удалось получить ОГРН", message)

    def _finish_ogrn(self, message: str) -> None:
        self.ogrn_button.setEnabled(True)
        self.ogrn_button.setText("Получить по ИНН")
        self.status_message.emit(message)

    def data(self) -> SoftwareConsentData:
        qdate = self.document_date.date()
        return SoftwareConsentData(
            program_name=self.program_name.text().strip(),
            applicant_name=self.applicant_name.text().strip(),
            applicant_address=self.applicant_address.toPlainText().strip(),
            inn=re.sub(r"\D", "", self.inn.text()),
            ogrn=re.sub(r"\D", "", self.ogrn.text()),
            application_number="",
            document_date=date(qdate.year(), qdate.month(), qdate.day()),
            authors=[editor.value() for editor in self.author_editors],
            authors_will_be_mentioned=self._authors_will_be_mentioned,
            source_path=self.source_label.text() if self.source_label.text() != "Анкета не загружена" else "",
            questionnaire_profile_id=self._questionnaire_profile_id,
            questionnaire_profile_name=self._questionnaire_profile_name,
        )

    def validation_errors(self) -> list[str]:
        data = self.data()
        errors = [f"Не заполнено: {field}." for field in data.missing_common_fields()]
        if data.inn and len(data.inn) not in (10, 12):
            errors.append("ИНН должен содержать 10 или 12 цифр.")
        if data.ogrn and len(data.ogrn) not in (13, 15):
            errors.append("ОГРН должен содержать 13 цифр, ОГРНИП — 15 цифр.")
        for index, author in enumerate(data.authors, 1):
            missing = author.missing_fields(
                require_personal_data=data.authors_will_be_mentioned is not False
            )
            if missing:
                errors.append(f"Автор {index}: не заполнены {', '.join(missing)}.")
        return errors

    @Slot()
    def _refresh_preview(self) -> None:
        if not self.author_editors:
            self.preview.clear()
            return
        index = max(0, min(self.author_tabs.currentIndex(), len(self.author_editors) - 1))
        data = self.data()
        author = data.authors[index]

        def value(text: str) -> str:
            return escape(text).replace("\n", "<br>") or '<span class="missing">не заполнено</span>'

        def passport_value(text: str) -> str:
            return value(re.sub(r"\s+", " ", text.replace("\u200b", "")).strip())

        issuer = re.split(
            r"(?i)\b(?:код\s+подразделения|дата\s+выдачи)\s*[:—-]?",
            author.passport_issuer,
            maxsplit=1,
        )[0].strip(" ,;:-")
        passport = (
            f"Паспорт гражданина РФ, серия «{passport_value(author.passport_series)}» "
            f"номер «{passport_value(author.passport_number)}» выдан "
            f"{passport_value(author.passport_issue_date)} {passport_value(issuer)}"
        )
        current_date = data.document_date.strftime("%d.%m.%Y")
        author_mention_choice = (
            "☐ упоминать его под своим именем &nbsp;&nbsp; ☒ не упоминать его (анонимно)"
            if data.authors_will_be_mentioned is False
            else "☒ упоминать его под своим именем &nbsp;&nbsp; ☐ не упоминать его (анонимно)"
        )
        html = f"""
        <style>
          body {{ color:#17212b; font-family:'Times New Roman'; background:#eef2f6; }}
          .page {{ background:white; border:1px solid #cbd5e1; margin:8px; padding:24px;
                   box-shadow:0 2px 7px rgba(15,23,42,.12); line-height:1.3; }}
          h3 {{ text-align:center; margin:18px 0; }}
          .right {{ text-align:right; }} .label {{ font-weight:bold; }}
          .missing {{ color:#b42318; background:#fee4e2; padding:1px 3px; }}
          .legal {{ margin:14px 0; text-align:justify; }}
        </style>
        <div class="page">
          <div class="right">В Федеральную службу<br>по интеллектуальной собственности<br>
          Бережковская наб., д. 30, корп. 1,<br>г. Москва, Г-59, ГСП-1, 119991,<br>Российская Федерация</div>
          <p>Название программы для ЭВМ или базы данных<br>«{value(data.program_name)}»</p>
          <h3>Согласие на обработку персональных данных</h3>
          <p><span class="label">Ф. И. О. субъекта персональных данных</span> {value(author.full_name)}</p>
          <p><span class="label">Адрес места жительства</span> {value(author.address)}</p>
          <p class="legal"><span class="label">Документ, удостоверяющий личность</span><br>{passport}</p>
          <p class="legal">Подтверждаю согласие на обработку моих персональных данных в целях
          предоставления Федеральной службой по интеллектуальной собственности государственной услуги.</p>
          <p>Подпись _________________ / {value(author.full_name)} /</p><p>{current_date}</p>
        </div>
        <div class="page">
          <h3>Согласие автора на указание сведений об авторе, указанных в заявлении</h3>
          <p>Название: «{value(data.program_name)}»</p>
          <p><span class="label">Правообладатель (Заявитель):</span><br>{value(data.applicant_name)}<br>
          {value(data.applicant_address)}<br>ОГРН: {value(data.ogrn)} &nbsp;&nbsp; ИНН: {value(data.inn)}</p>
          <p><span class="label">Фамилия имя отчество:</span> {value(author.full_name)}<br>
          <span class="label">Дата рождения:</span> {value(author.birth_date)} &nbsp;
          <span class="label">Гражданство:</span> {value(author.citizenship)}</p>
          <p><span class="label">Место постоянного жительства:</span><br>{value(author.address)} (RU)</p>
          <p><span class="label">Краткое описание творческого вклада:</span><br>
          {value(author.creative_contribution)}</p>
          <p>{author_mention_choice}</p>
          <p>Подпись автора: _________________ / {value(author.full_name)} /</p><p>{current_date}</p>
        </div>
        """
        self.preview.setHtml(html)

    @Slot()
    def save_document(self) -> None:
        errors = self.validation_errors()
        if errors:
            QMessageBox.warning(
                self,
                "Не хватает данных",
                "Проверьте обязательные поля:\n\n" + "\n".join(f"• {item}" for item in errors),
            )
            return
        data = self.data()
        desktop_value = QStandardPaths.writableLocation(
            QStandardPaths.StandardLocation.DesktopLocation
        )
        desktop = Path(desktop_value) if desktop_value else Path.home() / "Desktop"
        try:
            paths = save_author_consents(data, desktop, self.template_path)
        except Exception as exc:  # pragma: no cover - defensive GUI boundary
            QMessageBox.critical(self, "Не удалось сформировать согласия", str(exc))
            return
        self.status_message.emit(
            f"Создано согласий: {len(paths)}. Папка: {desktop}"
        )
        QMessageBox.information(
            self,
            "Готово",
            f"На рабочем столе создано документов: {len(paths)}.\n\nПапка:\n{desktop}",
        )
