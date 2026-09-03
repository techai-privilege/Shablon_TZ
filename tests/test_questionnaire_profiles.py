import json

import pytest
from docx import Document

from trademark_report.questionnaire import parse_questionnaire
from trademark_report.questionnaire_profiles import (
    DEFAULT_PROFILE_PATH,
    default_profile_library,
    load_profile_library,
)


def _write_questionnaire(path, rows):
    document = Document()
    table = document.add_table(rows=0, cols=2)
    for label, value in rows:
        cells = table.add_row().cells
        cells[0].text = label
        cells[1].text = value
    document.save(path)


def test_default_library_selects_known_and_fallback_profiles():
    library = default_profile_library()

    standard = library.select_profile(
        [
            "Наименование программы",
            "Количество авторов",
            "Программа создана за счет средств:",
            "Наименование заявителя и ИНН",
        ]
    )
    extended = library.select_profile(
        ["Область (сфера) применения", "ИНН заявителя"]
    )
    fallback = library.select_profile(["Наименование программы", "Неизвестное поле"])

    assert standard.profile_id == "cp_software_2026"
    assert extended.profile_id == "extended_software_questionnaire"
    assert fallback.profile_id == "generic_section_one"


def test_external_library_can_teach_new_label_without_python_changes(tmp_path):
    questionnaire = tmp_path / "new-format.docx"
    _write_questionnaire(
        questionnaire,
        [
            ("Наименование программы", "Новая программа"),
            ("Реквизиты правообладателя", "ООО «Пример», ИНН 7707083893"),
            ("Количество авторов", "0"),
        ],
    )

    payload = json.loads(DEFAULT_PROFILE_PATH.read_text(encoding="utf-8"))
    payload["profiles"].insert(
        0,
        {
            "id": "rightsholder_variant",
            "name": "Анкета с реквизитами правообладателя",
            "priority": 500,
            "match": {"all": ["Реквизиты правообладателя"]},
            "field_overrides": {
                "applicant": {
                    "match": "exact",
                    "labels": ["Реквизиты правообладателя"],
                }
            },
        },
    )
    custom_library_path = tmp_path / "profiles.json"
    custom_library_path.write_text(
        json.dumps(payload, ensure_ascii=False), encoding="utf-8"
    )
    library = load_profile_library(custom_library_path)

    result = parse_questionnaire(questionnaire, profile_library=library)

    assert result.data.applicant_name == "ООО «Пример»"
    assert result.data.inn == "7707083893"
    assert result.data.questionnaire_profile_id == "rightsholder_variant"
def test_library_rejects_unknown_schema_version(tmp_path):
    path = tmp_path / "profiles.json"
    path.write_text('{"schema_version": 99}', encoding="utf-8")

    with pytest.raises(ValueError, match="schema_version 1"):
        load_profile_library(path)
