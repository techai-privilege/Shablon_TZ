"""Create a combined, two-page-per-author consent DOCX from the retained template."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from datetime import datetime
from pathlib import Path
import re
import sys
from tempfile import NamedTemporaryFile, TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree

from .software_models import SoftwareAuthor, SoftwareConsentData


def _resource_root() -> Path:
    bundle_root = getattr(sys, "_MEIPASS", None)
    if getattr(sys, "frozen", False) and bundle_root:
        return Path(bundle_root)
    return Path(__file__).resolve().parents[2]


DEFAULT_CONSENT_TEMPLATE = _resource_root() / "assets" / "software_consent_template.docx"


def _iter_paragraphs(document: Document):
    yield from document.paragraphs
    for table in document.tables:
        for row in table.rows:
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:
                    continue
                seen.add(id(cell._tc))
                yield from cell.paragraphs
                for nested in cell.tables:
                    for nested_row in nested.rows:
                        for nested_cell in nested_row.cells:
                            yield from nested_cell.paragraphs


def _replace_in_runs(paragraph, old: str, new: str, *, count: int = -1) -> int:
    replaced = 0
    while count < 0 or replaced < count:
        runs = paragraph.runs
        full = "".join(run.text for run in runs)
        start = full.find(old)
        if start < 0:
            break
        end = start + len(old)
        positions = []
        cursor = 0
        for index, run in enumerate(runs):
            positions.append((index, cursor, cursor + len(run.text)))
            cursor += len(run.text)
        first = next((item for item in positions if item[1] <= start < item[2] or item[1] == start == item[2]), None)
        last = next((item for item in reversed(positions) if item[1] < end <= item[2]), None)
        if first is None or last is None:
            break
        first_index, first_start, _ = first
        last_index, last_start, _ = last
        prefix = runs[first_index].text[: start - first_start]
        suffix = runs[last_index].text[end - last_start :]
        runs[first_index].text = prefix + new + suffix
        for index in range(first_index + 1, last_index + 1):
            runs[index].text = ""
        replaced += 1
    return replaced


def _set_text_preserve(paragraph, text: str) -> None:
    runs = paragraph.runs
    if not runs:
        paragraph.add_run(text)
        return
    runs[0].text = text
    for run in runs[1:]:
        run.text = ""


def _passport_text(author: SoftwareAuthor) -> str:
    when_and_where = " ".join(
        item for item in (author.passport_issue_date, author.passport_issuer) if item
    )
    return (
        f"Паспорт гражданина РФ, серия «{author.passport_series}» "
        f"номер «{author.passport_number}» выдан «{when_and_where}»"
    )


def _fill_author_document(document: Document, data: SoftwareConsentData, author: SoftwareAuthor) -> None:
    date_text = data.document_date.strftime("%d.%m.%Y")
    birth = None
    try:
        birth = datetime.strptime(author.birth_date, "%d.%m.%Y")
    except ValueError:
        pass
    page_one = True
    exact_ellipsis_seen = 0
    signature_seen = 0
    date_seen = 0

    for paragraph in _iter_paragraphs(document):
        text = paragraph.text.strip()
        if not text:
            continue
        if text.startswith("Согласие автора на указание сведений"):
            page_one = False
        if "Название программы для ЭВМ или базы данных" in text:
            continue
        if text == "«…»":
            exact_ellipsis_seen += 1
            value = data.program_name if page_one or exact_ellipsis_seen == 1 else author.creative_contribution
            _set_text_preserve(paragraph, f"«{value}»")
            if not page_one and exact_ellipsis_seen > 1:
                size = 7.5 if len(value) > 500 else 8.5 if len(value) > 300 else 9.5
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.size = Pt(size)
            continue
        if text.startswith("№ заявки"):
            line = f"№ заявки {data.application_number}" if data.application_number else "№ заявки " + "_" * 58
            _set_text_preserve(paragraph, line)
            continue
        if text.startswith("Заявка №"):
            line = f"Заявка № {data.application_number}" if data.application_number else "Заявка № " + "_" * 62
            _set_text_preserve(paragraph, line)
            continue
        if text.startswith("Название:"):
            _set_text_preserve(paragraph, f"Название: «{data.program_name}»")
            continue
        if text.startswith("Ф. И. О. субъекта персональных данных"):
            _set_text_preserve(paragraph, f"Ф. И. О. субъекта персональных данных  {author.full_name}")
            continue
        if text.startswith("Адрес места жительства"):
            _set_text_preserve(paragraph, f"Адрес места жительства  {author.address}")
            continue
        if text.startswith("Документ, удостоверяющий личность"):
            prefix = "Документ, удостоверяющий личность субъекта персональных данных, дата его выдачи\nи выдавший орган  "
            _set_text_preserve(paragraph, prefix + _passport_text(author))
            continue
        if text.startswith("Общество с ограниченной ответственностью"):
            _set_text_preserve(paragraph, data.applicant_name)
            continue
        if text.startswith("Россия, «") and text.endswith("»"):
            _set_text_preserve(paragraph, data.applicant_address)
            continue
        if text.startswith("ОГРН:"):
            _set_text_preserve(paragraph, f"ОГРН: {data.ogrn}     ИНН: {data.inn}")
            continue
        if text.startswith("Фамилия имя отчество:"):
            _set_text_preserve(paragraph, f"Фамилия имя отчество: {author.full_name}")
            continue
        if text.startswith("Дата рождения:"):
            if birth:
                value = (
                    f"Дата рождения: число: {birth:%d}   месяц: {birth:%m}   год: {birth:%Y}   "
                    f"Гражданство: {author.citizenship}"
                )
            else:
                value = f"Дата рождения: {author.birth_date}   Гражданство: {author.citizenship}"
            _set_text_preserve(paragraph, value)
            continue
        if (
            data.authors_will_be_mentioned is False
            and "упоминать его под своим именем" in text
            and "не упоминать его (анонимно)" in text
        ):
            for run in paragraph.runs:
                run.text = (
                    run.text.replace("☒", "__CHECKED_BOX__")
                    .replace("☐", "☒")
                    .replace("__CHECKED_BOX__", "☐")
                )
            continue
        if text.startswith("Россия, «") and text.endswith("» (RU)"):
            _set_text_preserve(paragraph, f"{author.address} (RU)")
            continue
        if text.startswith("Подпись") and "ФИО" in text:
            signature_seen += 1
            label = "Подпись автора:" if not page_one else "Подпись"
            _set_text_preserve(paragraph, f"{label} ___________________/ {author.full_name} /")
            continue
        if text == "Дата":
            date_seen += 1
            _set_text_preserve(paragraph, date_text)

    # The template reserves a very tall final row for the attorney/date block.
    # A smaller minimum keeps long, unabridged author contributions on page 2.
    if document.tables and len(document.tables[0].rows) >= 6:
        row = document.tables[0].rows[5]
        cell = row.cells[0]
        empty_after_text = False
        seen_text = False
        for paragraph in list(cell.paragraphs):
            if paragraph.text.strip():
                seen_text = True
                continue
            keep = seen_text and not empty_after_text
            if keep:
                empty_after_text = True
            else:
                paragraph._element.getparent().remove(paragraph._element)
        tr_pr = row._tr.get_or_add_trPr()
        height = tr_pr.find(qn("w:trHeight"))
        if height is None:
            height = OxmlElement("w:trHeight")
            tr_pr.append(height)
        height.set(qn("w:val"), "1100")


def _clear_body(document: Document) -> None:
    body = document._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def _set_page_break_before(paragraph_element) -> None:
    """Keep every author's two-page block separate from the previous author."""

    paragraph_properties = paragraph_element.find(qn("w:pPr"))
    if paragraph_properties is None:
        paragraph_properties = OxmlElement("w:pPr")
        paragraph_element.insert(0, paragraph_properties)
    if paragraph_properties.find(qn("w:pageBreakBefore")) is None:
        paragraph_properties.append(OxmlElement("w:pageBreakBefore"))


def _strip_comments(path: Path) -> None:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with ZipFile(path, "r") as source, NamedTemporaryFile(suffix=".docx", delete=False) as temporary:
        temp_path = Path(temporary.name)
        with ZipFile(temporary, "w", ZIP_DEFLATED) as target:
            for item in source.infolist():
                name = item.filename
                if name.startswith("word/comments") or name == "word/people.xml":
                    continue
                data = source.read(name)
                if name == "word/document.xml":
                    root = etree.fromstring(data)
                    for tag in ("commentRangeStart", "commentRangeEnd"):
                        for node in root.xpath(f".//w:{tag}", namespaces=ns):
                            node.getparent().remove(node)
                    for node in root.xpath(".//w:commentReference", namespaces=ns):
                        run = node.getparent()
                        run.getparent().remove(run)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif name == "word/_rels/document.xml.rels":
                    root = etree.fromstring(data)
                    for relationship in list(root):
                        if any(part in relationship.get("Type", "") for part in ("comments", "people")):
                            root.remove(relationship)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif name == "[Content_Types].xml":
                    root = etree.fromstring(data)
                    for override in list(root):
                        if any(part in override.get("PartName", "") for part in ("comments", "people")):
                            root.remove(override)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                target.writestr(item, data)
    temp_path.replace(path)


def save_consents(
    data: SoftwareConsentData,
    output_path: str | Path,
    template_path: str | Path = DEFAULT_CONSENT_TEMPLATE,
) -> Path:
    if not data.authors:
        raise ValueError("Добавьте хотя бы одного автора.")
    template = Path(template_path)
    if not template.is_file():
        raise FileNotFoundError(f"Не найден шаблон согласия: {template}")

    final = Document(template)
    _clear_body(final)
    body = final._element.body
    section_properties = body.sectPr
    for author_index, author in enumerate(data.authors):
        author_document = Document(template)
        _fill_author_document(author_document, data, author)
        children = [
            deepcopy(child)
            for child in author_document._element.body
            if child.tag != qn("w:sectPr")
        ]
        if author_index:
            first_paragraph = next(
                (child for child in children if child.tag == qn("w:p")),
                None,
            )
            if first_paragraph is not None:
                _set_page_break_before(first_paragraph)
        for child in children:
            body.insert(body.index(section_properties), child)

    destination = Path(output_path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    final.save(destination)
    _strip_comments(destination)
    return destination


def _safe_author_name(value: str) -> str:
    """Return a file-name-safe FIO on both macOS and Windows."""

    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value or "Автор"


def _available_output_path(directory: Path, stem: str, reserved: set[Path]) -> Path:
    candidate = directory / f"{stem}.docx"
    counter = 2
    while candidate.exists() or candidate in reserved:
        candidate = directory / f"{stem} ({counter}).docx"
        counter += 1
    reserved.add(candidate)
    return candidate


def save_author_consents(
    data: SoftwareConsentData,
    output_directory: str | Path,
    template_path: str | Path = DEFAULT_CONSENT_TEMPLATE,
) -> list[Path]:
    """Create one two-page consent DOCX per author without overwriting files."""

    if not data.authors:
        raise ValueError("Добавьте хотя бы одного автора.")
    directory = Path(output_directory)
    directory.mkdir(parents=True, exist_ok=True)
    reserved: set[Path] = set()
    destinations = [
        _available_output_path(
            directory,
            f"Согласие {_safe_author_name(author.full_name)}",
            reserved,
        )
        for author in data.authors
    ]

    # Generate the whole batch first, so a failure does not leave a partial set.
    with TemporaryDirectory(prefix=".consents-", dir=directory) as temporary:
        temporary_directory = Path(temporary)
        staged: list[Path] = []
        for index, author in enumerate(data.authors, 1):
            staged_path = temporary_directory / f"{index}.docx"
            save_consents(
                replace(data, authors=[author]),
                staged_path,
                template_path,
            )
            staged.append(staged_path)
        for source, destination in zip(staged, destinations):
            source.replace(destination)
    return destinations
