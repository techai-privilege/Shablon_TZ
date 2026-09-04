"""Data structures for software-registration questionnaires and consents."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime

APPLICANT_LEGAL_ENTITY = "legal_entity"
APPLICANT_INDIVIDUAL = "individual"
APPLICANT_SOLE_PROPRIETOR = "sole_proprietor"
APPLICANT_TYPE_LABELS = {
    APPLICANT_LEGAL_ENTITY: "Юридическое лицо",
    APPLICANT_INDIVIDUAL: "Физическое лицо",
    APPLICANT_SOLE_PROPRIETOR: "Индивидуальный предприниматель",
}


def detect_applicant_type(name: str, inn: str, ogrn: str = "") -> str:
    """Infer applicant kind from explicit registration identifiers and wording."""

    normalized_name = re.sub(r"\s+", " ", name).strip().casefold()
    inn_digits = re.sub(r"\D", "", inn)
    ogrn_digits = re.sub(r"\D", "", ogrn)
    if len(ogrn_digits) == 15:
        return APPLICANT_SOLE_PROPRIETOR
    if len(ogrn_digits) == 13:
        return APPLICANT_LEGAL_ENTITY
    if (
        normalized_name.startswith("ип ")
        or (
            "индивидуальн" in normalized_name
            and "предпринимател" in normalized_name
        )
    ):
        return APPLICANT_SOLE_PROPRIETOR
    if len(inn_digits) == 12:
        return APPLICANT_INDIVIDUAL
    return APPLICANT_LEGAL_ENTITY


def is_valid_date_text(value: str) -> bool:
    """Return whether a text value is an exact, real DD.MM.YYYY date."""

    try:
        parsed = datetime.strptime(value.strip(), "%d.%m.%Y")
    except ValueError:
        return False
    return parsed.strftime("%d.%m.%Y") == value.strip()


def is_valid_ogrn(value: str) -> bool:
    """Validate OGRN/OGRNIP length and check digit."""

    digits = re.sub(r"\D", "", value)
    if len(digits) == 13:
        return int(digits[:12]) % 11 % 10 == int(digits[-1])
    if len(digits) == 15:
        return int(digits[:14]) % 13 % 10 == int(digits[-1])
    return False


@dataclass(slots=True)
class SoftwareAuthor:
    full_name: str = ""
    birth_date: str = ""
    citizenship: str = "Российская Федерация"
    address: str = ""
    passport_series: str = ""
    passport_number: str = ""
    passport_issue_date: str = ""
    passport_issuer: str = ""
    creative_contribution: str = ""
    rights_basis: str = ""
    source_text: str = ""

    def missing_fields(self, *, require_personal_data: bool = True) -> list[str]:
        labels = {
            "full_name": "ФИО",
            "creative_contribution": "творческий вклад",
        }
        if require_personal_data:
            labels.update(
                {
                    "birth_date": "дата рождения",
                    "address": "адрес",
                    "passport_series": "серия паспорта",
                    "passport_number": "номер паспорта",
                    "passport_issue_date": "дата выдачи паспорта",
                    "passport_issuer": "кем выдан паспорт",
                }
            )
        return [label for field_name, label in labels.items() if not getattr(self, field_name).strip()]


@dataclass(slots=True)
class SoftwareConsentData:
    program_name: str = ""
    applicant_name: str = ""
    applicant_address: str = ""
    inn: str = ""
    ogrn: str = ""
    applicant_type: str = APPLICANT_LEGAL_ENTITY
    application_number: str = ""
    document_date: date = field(default_factory=date.today)
    authors: list[SoftwareAuthor] = field(default_factory=list)
    authors_will_be_mentioned: bool | None = None
    declared_author_count: int | None = None
    source_path: str = ""
    questionnaire_profile_id: str = ""
    questionnaire_profile_name: str = ""

    def missing_common_fields(self) -> list[str]:
        labels = {
            "program_name": "название программы",
            "applicant_name": "наименование заявителя",
            "applicant_address": "адрес заявителя",
            "inn": "ИНН",
        }
        missing = [
            label for field_name, label in labels.items() if not getattr(self, field_name).strip()
        ]
        if self.applicant_type == APPLICANT_LEGAL_ENTITY and not self.ogrn.strip():
            missing.append("ОГРН")
        elif self.applicant_type == APPLICANT_SOLE_PROPRIETOR and not self.ogrn.strip():
            missing.append("ОГРНИП")
        return missing


@dataclass(slots=True)
class QuestionnaireParseResult:
    data: SoftwareConsentData
    warnings: list[str] = field(default_factory=list)
