"""Text templates mirrored from the Doczilla questionnaire screenshots."""

from __future__ import annotations

from .models import (
    CONCLUSION_CONSENT,
    CONCLUSION_NONE,
    CONCLUSION_OBSTACLES,
    ReportData,
)


EXPERT_DISCLAIMER = (
    "Экспертиза заявленного обозначения проводится экспертом ФИПСа (Роспатента), "
    "в связи с чем процесс проверки обладает долей субъективизма. В отчете указана "
    "часть сходных товарных знаков. Эксперт может противопоставить больше."
)

RECOMMENDATION_DISCLAIMER = (
    "Информация, содержащаяся в данном отчете, носит рекомендательный характер. "
    "Окончательное решение о соответствии заявленного обозначения требованиям "
    "законодательства выносит государственный эксперт."
)

INFRINGEMENT_WARNING = (
    "Обращаем внимание, что использование Вашего обозначения может быть признано "
    "нарушением исключительных прав на сходные существующие товарные знаки."
)


def _quoted_records(records) -> str:
    labels = [
        item.display_name.strip() or item.number.strip()
        for item in records
        if item.display_name.strip() or item.number.strip()
    ]
    return ", ".join(f"«{label}»" for label in labels)


def _found_records(report: ReportData) -> str:
    parts = []
    marks = _quoted_records(report.international_marks + report.russian_marks)
    applications = _quoted_records(report.applications)
    if marks:
        parts.append(f"товарные знаки {marks}")
    if applications:
        parts.append(f"заявки на товарные знаки {applications}")
    if not parts:
        return "сходные товарные знаки и/или заявки"
    return " и ".join(parts)


def _obstacle_verb(report: ReportData) -> str:
    definite = any(
        "препятствующий товарный знак" in option.lower()
        or "будет препятствовать регистрации" in option.lower()
        for option in report.relative_options
    )
    return "будут" if definite else "могут"


def conclusion_paragraphs(report: ReportData) -> list[tuple[str, str]]:
    """Return `(role, text)` paragraphs for the selected conclusion."""

    designation = report.designation.strip()
    found = _found_records(report)
    obstacle_verb = _obstacle_verb(report)

    if report.conclusion == CONCLUSION_NONE:
        paragraphs = [
            (
                "normal",
                f"В результате проведения проверки обозначения «{designation}» не найдено "
                "препятствий для его регистрации в качестве товарного знака.",
            ),
            (
                "bold",
                f"Рекомендуем заявлять обозначение «{designation}» на государственную "
                "регистрацию в качестве товарного знака.",
            ),
        ]
    elif report.conclusion == CONCLUSION_CONSENT:
        paragraphs = [
            (
                "normal",
                f"В результате проведения проверки обозначения «{designation}» найдены {found}, "
                f"которые {obstacle_verb} препятствовать регистрации.",
            ),
            (
                "normal",
                f"Существует возможность подать обозначение «{designation}» на регистрацию в "
                "качестве товарного знака. При получении запроса с противопоставлением найденных "
                "товарных знаков/заявок обратиться к правообладателям противопоставленных "
                "товарных знаков/заявок с целью получения от них писем-согласий на регистрацию.",
            ),
            ("warning", INFRINGEMENT_WARNING),
        ]
    elif report.conclusion == CONCLUSION_OBSTACLES:
        paragraphs = [
            (
                "normal",
                f"В результате проведения проверки обозначения «{designation}» найдены {found}, "
                f"которые {obstacle_verb} препятствовать регистрации.",
            ),
            (
                "bold",
                "Рекомендуем разработать новое обозначение и направить его на предварительную проверку.",
            ),
            ("normal", "Возможные варианты действий:"),
            (
                "list",
                f"Изменить обозначение полностью или доработать текущее. Например, составить "
                f"словосочетание со словом «{designation}». Важно, чтобы слово «{designation}» "
                "не было акцентным.",
            ),
            (
                "list",
                f"Подать обозначение «{designation}» на регистрацию в качестве товарного знака. "
                "В случае получения запроса с противопоставлением найденных товарных знаков "
                "обратиться к правообладателям противопоставленных товарных знаков с целью "
                "получения от них писем-согласий на регистрацию.",
            ),
            (
                "list",
                "Разработать обозначение, в котором словесный элемент будет выполнен оригинально "
                "и сложночитаемо. В таком случае защита будет распространяться только на "
                "оригинальное обозначение, а не на само слово.",
            ),
            (
                "list",
                f"Подать заявку на регистрацию комбинированного обозначения (логотипа), в "
                f"котором слово «{designation}» будет неохраняемым элементом.",
            ),
            ("warning", INFRINGEMENT_WARNING),
        ]
    else:
        raise ValueError(f"Неизвестный шаблон заключения: {report.conclusion}")

    paragraphs.extend((("italic", EXPERT_DISCLAIMER), ("italic", RECOMMENDATION_DISCLAIMER)))
    return paragraphs
