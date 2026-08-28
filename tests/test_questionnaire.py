from pathlib import Path

from docx import Document

from trademark_report.questionnaire import parse_questionnaire


def _questionnaire(path: Path) -> None:
    document = Document()
    table = document.add_table(rows=0, cols=2)
    values = [
        ("РАЗДЕЛ I – СВЕДЕНИЯ О ПО ДЛЯ РОСПАТЕНТА", ""),
        ("Наименование программы", "Тестовая программа"),
        ("Наименование заявителя и ИНН", "ООО «Тест», ИНН 7707083893"),
        ("Количество авторов", "2"),
        (
            "Паспортные данные авторов",
            "Автор 1:\n1. Иванов Иван Иванович\n2. 01.01.1980\n"
            "3. Россия, г. Москва, ул. Тестовая, д. 1\n4. 4510 123456\n"
            "5. ГУ МВД России по г. Москве 02.02.2020\n\n"
            "Автор 2:\n1. Петров Петр Петрович\n2. 03.03.1981\n"
            "3. Россия, г. Москва, ул. Вторая, д. 2\n4. 4511 654321\n"
            "5. ГУ МВД России по г. Москве 04.04.2021",
        ),
        (
            "Творческий вклад авторов",
            "Автор 1:\n- Разработка алгоритма\nАвтор 2:\n- Написание исходного текста программы",
        ),
    ]
    for label, value in values:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value

    ignored = document.add_table(rows=2, cols=2)
    ignored.cell(0, 0).text = "РАЗДЕЛ II – СВЕДЕНИЯ О ПО ДЛЯ РЕЕСТРА"
    ignored.cell(1, 0).text = "Наименование программы"
    ignored.cell(1, 1).text = "Это значение нельзя импортировать"
    document.save(path)


def test_parser_uses_only_section_one_and_splits_authors(tmp_path):
    path = tmp_path / "questionnaire.docx"
    _questionnaire(path)

    result = parse_questionnaire(path)

    assert result.data.program_name == "Тестовая программа"
    assert result.data.applicant_name == "ООО «Тест»"
    assert result.data.inn == "7707083893"
    assert result.data.declared_author_count == 2
    assert len(result.data.authors) == 2
    assert result.data.authors[0].full_name == "Иванов Иван Иванович"
    assert result.data.authors[0].passport_series == "4510"
    assert result.data.authors[0].passport_number == "123456"
    assert result.data.authors[0].creative_contribution == "Разработка алгоритма"
    assert result.data.authors[1].full_name == "Петров Петр Петрович"
    assert result.data.authors[1].birth_date == "03.03.1981"
    assert "Это значение нельзя импортировать" not in result.data.program_name


def test_parser_reads_program_name_above_section_one_table(tmp_path):
    path = tmp_path / "extended-questionnaire.docx"
    document = Document()
    document.add_paragraph("ПО «Складские системы RMS»")
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("РАЗДЕЛ I – СВЕДЕНИЯ О ПО ДЛЯ РОСПАТЕНТА", ""),
        ("Область (сфера) применения ПО", "Складская логистика"),
        ("ИНН заявителя", "6321416298"),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    result = parse_questionnaire(path)

    assert result.data.program_name == "Складские системы RMS"


def test_parser_does_not_warn_about_passport_data_when_authors_are_not_mentioned(tmp_path):
    path = tmp_path / "authors-not-mentioned.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("РАЗДЕЛ I – СВЕДЕНИЯ О ПО ДЛЯ РОСПАТЕНТА", ""),
        ("Наименование программы", "Тестовая программа"),
        ("Наименование заявителя и ИНН", "ООО «Тест», ИНН 7707083893"),
        ("Количество авторов", "1"),
        ("Будут ли упоминаться авторы при регистрации программы для ЭВМ?", "Нет"),
        ("ФИО авторов", "Иванов Иван Иванович"),
        ("Творческий вклад авторов", "Автор 1:\n- Разработка алгоритма"),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    result = parse_questionnaire(path)

    assert result.data.authors_will_be_mentioned is False
    author_warnings = [warning for warning in result.warnings if warning.startswith("Автор 1:")]
    assert author_warnings == []


def test_parser_keeps_non_passport_warnings_when_authors_are_not_mentioned(tmp_path):
    path = tmp_path / "authors-not-mentioned-and-empty.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("РАЗДЕЛ I – СВЕДЕНИЯ О ПО ДЛЯ РОСПАТЕНТА", ""),
        ("Наименование программы", "Тестовая программа"),
        ("Наименование заявителя и ИНН", "ООО «Тест», ИНН 7707083893"),
        ("Количество авторов", "1"),
        ("Будут ли указываться авторы при регистрации программы для ЭВМ?", "нет"),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    result = parse_questionnaire(path)

    author_warnings = [warning for warning in result.warnings if warning.startswith("Автор 1:")]
    assert author_warnings == ["Автор 1: не найдены ФИО, творческий вклад."]
