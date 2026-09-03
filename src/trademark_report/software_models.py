"""Data structures for software-registration questionnaires and consents."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


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
            "ogrn": "ОГРН/ОГРНИП",
        }
        return [label for field_name, label in labels.items() if not getattr(self, field_name).strip()]


@dataclass(slots=True)
class QuestionnaireParseResult:
    data: SoftwareConsentData
    warnings: list[str] = field(default_factory=list)
