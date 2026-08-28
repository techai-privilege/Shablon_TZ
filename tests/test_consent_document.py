from datetime import date
from zipfile import ZipFile

from docx import Document

from trademark_report.consent_document import (
    DEFAULT_CONSENT_TEMPLATE,
    save_author_consents,
    save_consents,
)
from trademark_report.software_models import SoftwareAuthor, SoftwareConsentData


def _data() -> SoftwareConsentData:
    return SoftwareConsentData(
        program_name="Тестовая программа",
        applicant_name="ООО «Тест»",
        applicant_address="Россия, г. Москва, ул. Тестовая, д. 1",
        inn="7707083893",
        ogrn="1027700132195",
        application_number="2026888888",
        document_date=date(2026, 8, 26),
        authors=[
            SoftwareAuthor(
                full_name="Иванов Иван Иванович",
                birth_date="01.01.1980",
                address="Россия, г. Москва, ул. Первая, д. 1",
                passport_series="4510",
                passport_number="123456",
                passport_issue_date="02.02.2020",
                passport_issuer="ГУ МВД России по г. Москве",
                creative_contribution="Разработка алгоритма",
            ),
            SoftwareAuthor(
                full_name="Петров Петр Петрович",
                birth_date="03.03.1981",
                address="Россия, г. Москва, ул. Вторая, д. 2",
                passport_series="4511",
                passport_number="654321",
                passport_issue_date="04.04.2021",
                passport_issuer="ГУ МВД России по г. Москве",
                creative_contribution="Написание исходного текста программы",
            ),
        ],
    )


def test_combined_consents_preserve_template_and_remove_comments(tmp_path):
    output = tmp_path / "consents.docx"
    save_consents(_data(), output)

    document = Document(output)
    assert len(document.tables) == 2
    assert len(document.comments) == 0
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    text += "\n" + "\n".join(
        cell.text for table in document.tables for row in table.rows for cell in row.cells
    )
    assert "Иванов Иван Иванович" in text
    assert "Петров Петр Петрович" in text
    assert "Тестовая программа" in text
    assert "1027700132195" in text
    assert "26.08.2026" in text
    assert "ФИО" not in text
    recipient_paragraphs = [
        paragraph
        for paragraph in document.paragraphs
        if paragraph.text.strip() == "В Федеральную службу"
    ]
    assert len(recipient_paragraphs) == 2
    assert recipient_paragraphs[0].paragraph_format.page_break_before is not True
    assert recipient_paragraphs[1].paragraph_format.page_break_before is True

    with ZipFile(output) as archive:
        names = set(archive.namelist())
        assert "word/comments.xml" not in names
        assert b"commentReference" not in archive.read("word/document.xml")


def test_consent_template_is_bundled():
    assert DEFAULT_CONSENT_TEMPLATE.is_file()


def test_anonymous_author_choice_is_checked_in_generated_document(tmp_path):
    output = tmp_path / "anonymous-consent.docx"
    data = _data()
    data.authors_will_be_mentioned = False

    save_consents(data, output)

    document = Document(output)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    text += "\n" + "\n".join(
        cell.text for table in document.tables for row in table.rows for cell in row.cells
    )
    assert "☐ упоминать его под своим именем     ☒ не упоминать его (анонимно)" in text
    assert "☒ упоминать его под своим именем     ☐ не упоминать его (анонимно)" not in text


def test_separate_consent_is_created_for_each_author(tmp_path):
    paths = save_author_consents(_data(), tmp_path)

    assert [path.name for path in paths] == [
        "Согласие Иванов Иван Иванович.docx",
        "Согласие Петров Петр Петрович.docx",
    ]
    first = Document(paths[0])
    second = Document(paths[1])
    first_text = "\n".join(paragraph.text for paragraph in first.paragraphs)
    second_text = "\n".join(paragraph.text for paragraph in second.paragraphs)
    assert len(first.tables) == 1
    assert len(second.tables) == 1
    assert "Иванов Иван Иванович" in first_text
    assert "Петров Петр Петрович" not in first_text
    assert "Петров Петр Петрович" in second_text
    assert "Иванов Иван Иванович" not in second_text


def test_separate_consents_do_not_overwrite_existing_files(tmp_path):
    existing = tmp_path / "Согласие Иванов Иван Иванович.docx"
    existing.write_bytes(b"existing")

    paths = save_author_consents(_data(), tmp_path)

    assert existing.read_bytes() == b"existing"
    assert paths[0].name == "Согласие Иванов Иван Иванович (2).docx"
