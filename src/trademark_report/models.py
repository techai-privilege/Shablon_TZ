"""Application data structures used by the UI and DOCX generator."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


PROBABILITY_VALUES = (
    "ВЫСОКАЯ",
    "ВЫШЕ СРЕДНЕЙ",
    "СРЕДНЯЯ",
    "НИЖЕ СРЕДНЕЙ",
    "НИЗКАЯ",
)

CONCLUSION_NONE = "Нет препятствий"
CONCLUSION_OBSTACLES = "Есть препятствия"
CONCLUSION_CONSENT = "Письма-согласия"
CONCLUSION_VALUES = (CONCLUSION_NONE, CONCLUSION_OBSTACLES, CONCLUSION_CONSENT)

RELATIVE_OPTIONS = (
    "Отсутствуют",
    "Найден препятствующий товарный знак (см. Приложение 1)",
    "Найден товарный знак, который может препятствовать регистрации (см. Приложение 1)",
    "Найдена заявка, которая будет препятствовать регистрации (см. Приложение 1)",
    "Найдена заявка, которая может препятствовать регистрации (см. Приложение 1)",
)

PERFORMERS = {
    "Алина": "Дукки Алина",
    "Маша": "Шмыкова Мария",
    "Лера": "Чернова Валерия",
}


@dataclass(slots=True)
class ReportNiceClass:
    number: str
    description: str


@dataclass(slots=True)
class ProbabilityEntry:
    value: str
    subject: str = ""


@dataclass(slots=True)
class ConclusionRun:
    text: str
    bold: bool = False
    italic: bool = False
    highlighted: bool = False


@dataclass(slots=True)
class ConclusionParagraph:
    runs: list[ConclusionRun] = field(default_factory=list)
    list_item: bool = False


@dataclass(slots=True)
class SimilarRecord:
    kind: str
    source_url: str = ""
    display_name: str = ""
    number: str = ""
    status: str = ""
    relevant_date: str = ""
    owner_or_applicant: str = ""
    related_classes: str = ""
    unprotected_element: str = ""
    image_bytes: bytes | None = None


@dataclass(slots=True)
class ReportData:
    designation: str
    search_queries: str
    business_area: str = ""
    report_date: date = field(default_factory=date.today)
    nice_classes: list[ReportNiceClass] = field(default_factory=list)
    has_absolute_grounds: bool = False
    absolute_grounds_text: str = ""
    relative_options: list[str] = field(default_factory=lambda: ["Отсутствуют"])
    trademarks_database_date: str = ""
    applications_database_date: str = ""
    conclusion: str = CONCLUSION_NONE
    conclusion_content: list[ConclusionParagraph] | None = None
    performer: str = "Алина"
    probabilities: list[ProbabilityEntry] = field(default_factory=list)
    international_marks: list[SimilarRecord] = field(default_factory=list)
    russian_marks: list[SimilarRecord] = field(default_factory=list)
    applications: list[SimilarRecord] = field(default_factory=list)

    @property
    def has_appendix(self) -> bool:
        return bool(self.international_marks or self.russian_marks or self.applications)
