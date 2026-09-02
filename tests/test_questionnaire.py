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


def test_parser_removes_service_fields_from_passport_issuer(tmp_path):
    path = tmp_path / "passport-service-fields.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("РАЗДЕЛ I – СВЕДЕНИЯ О ПО ДЛЯ РОСПАТЕНТА", ""),
        ("Наименование программы", "Тестовая программа"),
        ("Количество авторов", "1"),
        (
            "Паспортные данные авторов",
            "Автор 1:\n1. Иванов Иван Иванович\n2. 01.01.1980\n"
            "3. Россия, г. Москва\n4. 4510 123456\n"
            "5. ГУ МВД России по г. Москве 02.02.2020 "
            "Код подразделения: 770-068 Дата выдачи:",
        ),
        ("Творческий вклад авторов", "Автор 1:\n- Разработка алгоритма"),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    author = parse_questionnaire(path).data.authors[0]

    assert author.passport_issue_date == "02.02.2020"
    assert author.passport_issuer == "ГУ МВД России по г. Москве"


def test_parser_reads_section_one_split_across_two_tables(tmp_path):
    path = tmp_path / "split-section.docx"
    document = Document()
    first = document.add_table(rows=0, cols=2)
    for label, value in (
        ("Наименование программы", "1YES Analitik"),
        ("Область (сфера) применения", "Аналитика"),
    ):
        cells = first.add_row().cells
        cells[0].text = label
        cells[1].text = value
    second = document.add_table(rows=0, cols=2)
    for label, value in (
        ("Наименование заявителя и ИНН", "ООО «Старшип.про», ИНН 1683025264"),
        ("Количество авторов", "1"),
        (
            "Паспортные данные автора 1",
            "- УРБАН АЛЕКСАНДР СЕРГЕЕВИЧ\n- 25.03.1988\n"
            "- 141407, ГОРОД ХИМКИ, ПРОСПЕКТ ЮБИЛЕЙНЫЙ, ДОМ 49\n"
            "- 4615 931690, ТП №1 ОУФМС РОССИИ ПО МОСКОВСКОЙ ОБЛАСТИ, 25.06. 2015",
        ),
        ("Описание творческого вклада автора 1", "- разработка спецификаций программы"),
    ):
        cells = second.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    result = parse_questionnaire(path)
    author = result.data.authors[0]

    assert result.data.applicant_name == "ООО «Старшип.про»"
    assert result.data.inn == "1683025264"
    assert author.full_name == "Урбан Александр Сергеевич"
    assert author.passport_series == "4615"
    assert author.passport_number == "931690"
    assert author.passport_issue_date == "25.06.2015"
    assert author.passport_issuer.startswith("ТП №1 ОУФМС")


def test_parser_reads_labeled_passport_line_and_address(tmp_path):
    path = tmp_path / "labeled-passport.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("Наименование программы", "Лина"),
        ("Количество авторов", "1"),
        (
            "Паспортные данные Автора 1 / Описание творческого вклада автора 1",
            "Автор: Баскаков Александр Сергеевич\n"
            "- дата рождения: 15.05.2002\n"
            "- адрес места жительства: г. Москва, ул. Черкизовская Б., д. 32\n"
            "- серия и номер паспорта, дата выдачи и выдавший орган: "
            "4522 868797 02.06.2022 ГУ МВД России по г. Москве",
        ),
        (
            "Паспортные данные Автора 1 / Описание творческого вклада автора 1",
            "- разработка всей программы в целом",
        ),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    author = parse_questionnaire(path).data.authors[0]

    assert author.address == "г. Москва, ул. Черкизовская Б., д. 32"
    assert author.passport_series == "4522"
    assert author.passport_number == "868797"
    assert author.passport_issue_date == "02.06.2022"
    assert author.passport_issuer == "ГУ МВД России по г. Москве"


def test_parser_supports_foreign_passport_and_textual_issue_date(tmp_path):
    path = tmp_path / "passport-variants.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    rows = (
        ("Наименование программы", "КИСС"),
        ("Количество авторов", "2"),
        (
            "Паспортные данные Автора 1",
            "Скориков Максим Сергеевич\n15.07.1991\n"
            "420500, Республика Татарстан, г. Иннополис\n"
            "Паспорт РК № 10930256\nвыдан МИД РК, 29.03.2017",
        ),
        ("Описание творческого вклада автора 1", "- разработка программы в целом"),
        (
            "Паспортные данные Автора 2",
            "Корнеев Кирилл Андреевич\n03.09.1997 года рождения\n"
            "664009, Иркутская область, город Иркутск\nПаспорт: 2517 461759\n"
            "Отделом УФМС России по Иркутской области, выдан «14» ноября 2017 г.",
        ),
        ("Описание творческого вклада автора 2", "- разработка программы в целом"),
    )
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    authors = parse_questionnaire(path).data.authors

    assert (authors[0].passport_series, authors[0].passport_number) == ("РК", "10930256")
    assert authors[0].address.startswith("420500")
    assert authors[1].passport_issue_date == "14.11.2017"
    assert authors[1].passport_issuer.startswith("Отделом УФМС")


def test_parser_separates_contribution_from_passport_and_address(tmp_path):
    path = tmp_path / "combined-passport-contribution.docx"
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in (
        ("Наименование программы", "Пересвет"),
        ("Количество авторов", "1"),
        (
            "Паспортные данные Автора 1 / Описание творческого вклада автора 1",
            "Кобелев Денис Александрович, 11.12.1986 г.р., паспорт РФ серии "
            "5707 №013139, выдан 16.01.2007 г. Управлением внутренних дел города Кунгура, "
            "зарегистрированного по адресу: Пермский край, г. Кунгур, ул. Новая, д. 16.\n"
            "- Разработка системной архитектуры\n- Разработка интерфейсов",
        ),
    ):
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)

    author = parse_questionnaire(path).data.authors[0]

    assert author.address == "Пермский край, г. Кунгур, ул. Новая, д. 16."
    assert author.passport_issuer == "Управлением внутренних дел города Кунгура"
    assert author.creative_contribution == "Разработка системной архитектуры; Разработка интерфейсов"


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
