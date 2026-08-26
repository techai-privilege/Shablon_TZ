from trademark_report.models import ReportData, SimilarRecord
from trademark_report.templates import conclusion_paragraphs


def _texts(report: ReportData) -> list[str]:
    return [text for _, text in conclusion_paragraphs(report)]


def test_obstacle_conclusion_separates_marks_and_applications():
    report = ReportData(
        designation="FELICE",
        search_queries="FELICE",
        conclusion="Есть препятствия",
        relative_options=["Найден препятствующий товарный знак (см. Приложение 1)"],
        russian_marks=[SimilarRecord(kind="russian", display_name="FELICHE")],
        applications=[SimilarRecord(kind="application", number="2026708897")],
    )

    texts = _texts(report)
    assert (
        "найдены товарные знаки «FELICHE» и заявки на товарные знаки «2026708897», "
        "которые будут препятствовать регистрации."
    ) in texts[0]
    assert "словосочетание со словом «FELICE»" in texts[3]
    assert "слово «FELICE» будет неохраняемым элементом" in texts[6]


def test_consent_conclusion_uses_potential_wording():
    report = ReportData(
        designation="SAULUK",
        search_queries="SAULUK",
        conclusion="Письма-согласия",
        relative_options=["Найдена заявка, которая может препятствовать регистрации (см. Приложение 1)"],
        applications=[SimilarRecord(kind="application", display_name="SAULUK WAY")],
    )

    texts = _texts(report)
    assert "которые могут препятствовать регистрации" in texts[0]
    assert "получения от них писем-согласий на регистрацию" in texts[1]
