"""Extract consent fields from the supported software questionnaire variants."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import re

from docx import Document

from .software_models import QuestionnaireParseResult, SoftwareAuthor, SoftwareConsentData


DATE_RE = re.compile(r"\b(\d{2}\.\d{2}\.\d{4})\s*(?:г\.?|года)?", re.IGNORECASE)
NAME_RE = re.compile(
    r"\b([А-ЯЁ][а-яё-]+(?:-[А-ЯЁ][а-яё-]+)?\s+"
    r"[А-ЯЁ][а-яё-]+(?:-[А-ЯЁ][а-яё-]+)?\s+"
    r"[А-ЯЁ][а-яё-]+(?:-[А-ЯЁ][а-яё-]+)?)\b"
)


def _clean(value: str) -> str:
    value = value.replace("\xa0", " ").replace("\u200b", "")
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    return value.strip(" \n;,")


def _norm(value: str) -> str:
    return re.sub(r"[^а-яёa-z0-9]+", " ", value.lower().replace("ё", "е")).strip()


def _cell_text(cell) -> str:
    return _clean("\n".join(paragraph.text for paragraph in cell.paragraphs if paragraph.text.strip()))


def _candidate_table(document: Document):
    candidates = []
    for index, table in enumerate(document.tables):
        whole = _norm(" ".join(cell.text for row in table.rows for cell in row.cells))
        if "раздел ii" in whole and "раздел i " not in whole:
            break
        if "наименование программы" not in whole and "область сфера применения" not in whole:
            continue
        score = 0
        for row in table.rows:
            if len(row.cells) >= 2 and _cell_text(row.cells[1]):
                score += 1
        candidates.append((score, -index, table))
    if not candidates:
        raise ValueError("В документе не найдена таблица раздела I со сведениями для Роспатента.")
    return max(candidates, key=lambda item: (item[0], item[1]))[2]


def _rows(table) -> list[tuple[str, str]]:
    result = []
    for row in table.rows:
        if len(row.cells) < 2:
            continue
        label = _cell_text(row.cells[0])
        value = _cell_text(row.cells[1])
        if re.match(r"^раздел\s+ii(?:\s|$)", _norm(label)):
            break
        if label and _norm(label) != _norm(value):
            result.append((label, value))
    return result


def _first_value(rows, *needles: str) -> str:
    normalized = [(_norm(label), value) for label, value in rows]
    for needle in needles:
        target = _norm(needle)
        for label, value in normalized:
            if target in label and value:
                return value
    return ""


def _first_value_exact(rows, *labels: str) -> str:
    wanted = {_norm(label) for label in labels}
    for label, value in rows:
        if _norm(label).rstrip(":") in wanted and value:
            return value
    return ""


def _authors_will_be_mentioned(rows) -> bool | None:
    """Read an explicit author-disclosure answer from section I.

    Some questionnaire versions use ``упоминаться``, while others use
    ``указываться``.  An ambiguous, unedited ``да / нет`` value must not hide
    validation errors, so it is represented as ``None``.
    """

    value = _first_value(
        rows,
        "Будут ли упоминаться авторы",
        "Будут ли указываться авторы",
    )
    normalized = _norm(value)
    if re.match(r"^нет(?:\s|$)", normalized):
        return False
    if re.match(r"^да(?:\s|$)", normalized) and "нет" not in normalized.split():
        return True
    return None


def _extract_inn(value: str) -> str:
    match = re.search(r"(?<!\d)(\d{10}|\d{12})(?!\d)", value.replace(" ", ""))
    return match.group(1) if match else ""


def _extract_applicant(value: str, inn: str) -> tuple[str, str]:
    if not value:
        return "", ""
    text = value
    if inn:
        text = text.replace(inn, "")
    text = re.sub(r"\bИНН\s*[:№-]?", "", text, flags=re.IGNORECASE)
    lines = [_clean(line) for line in text.splitlines() if _clean(line)]
    address = ""
    name_parts = []
    for line in lines:
        if re.search(r"\bадрес\b", line, re.IGNORECASE):
            address = re.sub(r"^.*?адрес\s*[:—-]?\s*", "", line, flags=re.IGNORECASE)
        else:
            name_parts.append(line)
    name = _clean(" ".join(name_parts)).strip(" ,;-()")
    return name, address


def _extract_program_name(document: Document) -> str:
    """Read the standalone `ПО «…»` title used by extended questionnaires."""

    pattern = re.compile(r"(?i)^ПО\s*[«\"]\s*(.+?)\s*[»\"]\s*$")
    for paragraph in document.paragraphs:
        match = pattern.match(_clean(paragraph.text))
        if match:
            return _clean(match.group(1))
    return ""


def _extract_names(value: str) -> list[str]:
    result = []
    for match in NAME_RE.finditer(value):
        name = match.group(1)
        if name not in result and not any(
            phrase in name for phrase in ("Российская Федерация", "Федерального закона")
        ):
            result.append(name)
    return result


def _split_author_chunks(value: str) -> list[str]:
    text = _clean(value)
    if not text:
        return []
    explicit = re.compile(r"(?im)(?=^[ \t]*Автор[ \t]+\d+[ \t]*:)")
    explicit_starts = [match.start() for match in explicit.finditer(text)]
    if explicit_starts:
        return [chunk for chunk in [
            _clean(text[start : explicit_starts[index + 1] if index + 1 < len(explicit_starts) else None])
            for index, start in enumerate(explicit_starts)
        ] if chunk]
    marker = re.compile(
        r"(?im)(?:^|\n)\s*(?:автор\s*)?\d+\s*[.:)]\s*"
        r"(?=[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё-]+)"
    )
    matches = list(marker.finditer(text))
    if not matches:
        return [text]
    chunks = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        chunks.append(_clean(text[match.end() : end]))
    return [chunk for chunk in chunks if chunk]


def _strip_author_prefix(value: str) -> str:
    return _clean(re.sub(r"(?i)^\s*(?:автор\s*)?\d+\s*[.:)]\s*", "", value))


def _extract_passport(author: SoftwareAuthor, value: str) -> None:
    text = _strip_author_prefix(value)
    author.source_text = text
    names = _extract_names(text)
    if names and not author.full_name:
        author.full_name = names[0]

    birth = re.search(r"(?i)(?:дата\s+рождения\s*[:—-]?\s*|)(\d{2}\.\d{2}\.\d{4})\s*(?:г\.?р\.?)", text)
    if not birth:
        birth = re.search(r"(?i)дата\s+рождения\s*[:—-]?\s*(\d{2}\.\d{2}\.\d{4})", text)
    dates = DATE_RE.findall(text)
    if birth:
        author.birth_date = birth.group(1)
    elif dates:
        author.birth_date = dates[0]

    passport_patterns = (
        r"(?i)паспорт(?:\s+гражданина\s+РФ)?\s*[:—-]?\s*(?:серия|серии)?\s*[«\"]?(\d{2})\s*(\d{2})[»\"]?\s*[,;]?\s*(?:номер|№)?\s*[«\"]?(\d{6})",
        r"(?i)паспорт(?:\s+РФ)?\s*(\d{4})\s*(\d{6})",
        r"(?i)(?:серия|серии)\s*[«\"]?(\d{2})\s*(\d{2})[»\"]?\s*[,;]?\s*(?:номер|№)\s*[«\"]?(\d{6})",
        r"(?<!\d)(\d{2})\s+(\d{2})\s+(\d{6})(?!\d)",
    )
    for pattern in passport_patterns:
        match = re.search(pattern, text)
        if not match:
            continue
        groups = match.groups()
        if len(groups) == 3:
            author.passport_series = groups[0] + groups[1]
            author.passport_number = groups[2]
        else:
            author.passport_series, author.passport_number = groups
        break

    indexed = {
        int(n): _clean(v)
        for n, v in re.findall(r"(?m)^[ \t]*([1-5])\.[ \t]*(.+)$", text)
    }
    if indexed:
        author.full_name = author.full_name or indexed.get(1, "")
        indexed_birth = DATE_RE.search(indexed.get(2, ""))
        author.birth_date = author.birth_date or (indexed_birth.group(1) if indexed_birth else "")
        author.address = author.address or indexed.get(3, "")
        digits = re.sub(r"\D", "", indexed.get(4, ""))
        if len(digits) >= 10:
            author.passport_series = author.passport_series or digits[:4]
            author.passport_number = author.passport_number or digits[4:10]
        author.passport_issuer = author.passport_issuer or indexed.get(5, "")
        issue_dates = DATE_RE.findall(indexed.get(5, ""))
        if issue_dates:
            author.passport_issue_date = issue_dates[-1]
            author.passport_issuer = _clean(indexed[5].replace(issue_dates[-1], ""))

    issue_label = re.search(r"(?i)дата\s+выдачи\s*[:—-]?\s*(\d{2}\.\d{2}\.\d{4})", text)
    issued_before = re.search(
        r"(?is)\bвыдан(?:ный|а)?\s+(\d{2}\.\d{2}\.\d{4})\s*(?:г\.\s*)?(.+?)(?=\b(?:зарегистр|прожива|место\s+жительства|дата\s+рождения|код\s+подразделения)\b|$)",
        text,
    )
    issued_after = re.search(
        r"(?is)\bвыдан(?:ный|а)?\s+(.+?)\s+(\d{2}\.\d{2}\.\d{4})(?=\s|$)", text
    )
    if issue_label:
        author.passport_issue_date = issue_label.group(1)
    elif issued_before:
        author.passport_issue_date = issued_before.group(1)
    elif issued_after:
        author.passport_issue_date = issued_after.group(2)
    elif len(dates) >= 2:
        author.passport_issue_date = next((item for item in dates if item != author.birth_date), "")

    if issued_before:
        author.passport_issuer = _clean(issued_before.group(2))
    elif issued_after:
        author.passport_issuer = _clean(issued_after.group(1))
    else:
        issuer = re.search(
            r"(?is)\bвыдан(?:ный|а)?\s+(.+?)(?=\b(?:код\s+подразделения|дата\s+выдачи|место\s+жительства|дата\s+рождения|зарегистр)\b|$)",
            text,
        )
        if issuer:
            author.passport_issuer = _clean(issuer.group(1))
    if not author.passport_issuer:
        issuer = re.search(
            r"(?is)орган,?\s+выдавший\s+документ\s*:\s*(?:орган\s+)?(.+?)(?=\b(?:зарегистр|прожива)\b|$)",
            text,
        )
        if issuer:
            author.passport_issuer = _clean(issuer.group(1))
    if author.passport_issue_date and author.passport_issuer:
        author.passport_issuer = _clean(author.passport_issuer.replace(author.passport_issue_date, ""))

    address_patterns = (
        r"(?is)(?:зарегистрирован\w*|проживающ\w*)\s+(?:по\s+)?адресу\s*[:—-]?\s*(.+)$",
        r"(?is)место\s+жительства\s*[:—-]?\s*(.+?)(?=\bдата\s+рождения\b|$)",
    )
    for pattern in address_patterns:
        match = re.search(pattern, text)
        if match:
            author.address = _clean(match.group(1)).lstrip("-–—•● ")
            break
    if not author.address:
        lines = [_clean(line) for line in text.splitlines() if _clean(line)]
        for line in lines:
            if re.search(r"(?i)\b(россия|российская федерация|обл\.?|край|г\.|город|ул\.|д\.)\b", line) and not re.search(r"(?i)паспорт|выдан", line):
                if author.full_name not in line:
                    author.address = line.lstrip("-–—•● ")
                    break


def _clean_contribution(value: str) -> str:
    value = re.sub(r"(?i)^\s*Автор\s+\d+\s*:\s*", "", value)
    lines = []
    for line in value.splitlines():
        line = re.sub(r"^\s*[-–—•●]+\s*", "", line).strip()
        if line:
            lines.append(line.rstrip(";."))
    if len(lines) > 1 and NAME_RE.fullmatch(lines[0].rstrip(".")):
        lines.pop(0)
    return "; ".join(lines)


def parse_questionnaire(path: str | Path) -> QuestionnaireParseResult:
    source = Path(path)
    document = Document(source)
    rows = _rows(_candidate_table(document))
    warnings: list[str] = []
    authors_will_be_mentioned = _authors_will_be_mentioned(rows)

    program_name = _first_value(rows, "Наименование программы") or _extract_program_name(document)
    applicant_raw = _first_value_exact(
        rows,
        "Наименование заявителя и ИНН",
        "Данные заявителя",
        "Заявитель",
        "ИНН заявителя",
    )
    inn = _extract_inn(applicant_raw)
    applicant_name, applicant_address = _extract_applicant(applicant_raw, inn)
    count_value = _first_value(rows, "Количество авторов")
    count_match = re.search(r"\d+", count_value)
    declared_count = int(count_match.group()) if count_match else None

    names_value = _first_value(rows, "ФИО авторов")
    known_names = _extract_names(names_value)
    authors: list[SoftwareAuthor] = [SoftwareAuthor(full_name=name) for name in known_names]
    by_number: defaultdict[int, dict[str, list[str]]] = defaultdict(lambda: defaultdict(list))
    sequential_passports: list[str] = []
    sequential_contributions: list[str] = []
    sequential_bases: list[str] = []

    for label, value in rows:
        normalized = _norm(label)
        if not value:
            continue
        number_match = re.search(r"автора?\s*(\d+)", normalized)
        number = int(number_match.group(1)) if number_match else None
        looks_passport = "паспортн" in normalized
        looks_contribution = "творческ" in normalized or "описание творческого вклада" in normalized
        looks_basis = "основан" in normalized and "прав" in normalized

        if looks_basis:
            chunks = _split_author_chunks(value)
            if number:
                by_number[number]["basis"].append(value)
            elif len(chunks) > 1:
                sequential_bases.extend(chunks)
            else:
                sequential_bases.append(value)
            continue
        if looks_passport and looks_contribution:
            if re.search(r"(?i)паспорт|серия|номер|выдан|дата рождения|г\.р\.", value):
                chunks = _split_author_chunks(value)
                if len(chunks) > 1:
                    sequential_passports.extend(chunks)
                else:
                    (by_number[number]["passport"] if number else sequential_passports).append(value)
                if re.search(r"(?im)^\s*[-–—•●]+\s*(?:разработ|написан|системн|формализац|программ)", value):
                    contribution = "\n".join(
                        line for line in value.splitlines()
                        if re.match(r"(?i)^\s*[-–—•●]+\s*(?:разработ|написан|системн|формализац|программ)", line)
                    )
                    if contribution:
                        (by_number[number]["contribution"] if number else sequential_contributions).append(contribution)
            else:
                chunks = _split_author_chunks(value)
                if len(chunks) > 1:
                    sequential_contributions.extend(chunks)
                else:
                    (by_number[number]["contribution"] if number else sequential_contributions).append(value)
            continue
        if looks_passport:
            chunks = _split_author_chunks(value)
            if len(chunks) > 1:
                sequential_passports.extend(chunks)
            elif number:
                by_number[number]["passport"].append(value)
            elif chunks:
                sequential_passports.extend(chunks)
            continue
        if looks_contribution:
            chunks = _split_author_chunks(value)
            if len(chunks) > 1:
                sequential_contributions.extend(chunks)
            elif number:
                by_number[number]["contribution"].append(value)
            elif chunks:
                sequential_contributions.extend(chunks)

    target_count = max(
        declared_count or 0,
        len(authors),
        max(by_number, default=0),
        len(sequential_passports),
        len(sequential_contributions),
    )
    while len(authors) < target_count:
        authors.append(SoftwareAuthor())

    for index, author in enumerate(authors, 1):
        passport_values = by_number[index]["passport"] or (
            [sequential_passports[index - 1]] if index <= len(sequential_passports) else []
        )
        if passport_values:
            _extract_passport(author, "\n".join(passport_values))
        contribution_values = by_number[index]["contribution"] or (
            [sequential_contributions[index - 1]] if index <= len(sequential_contributions) else []
        )
        if contribution_values:
            author.creative_contribution = _clean_contribution("\n".join(contribution_values))
        basis_values = by_number[index]["basis"] or (
            [sequential_bases[index - 1]] if index <= len(sequential_bases) else []
        )
        if basis_values:
            author.rights_basis = _clean_contribution("\n".join(basis_values))

    if declared_count is not None and len(authors) != declared_count:
        warnings.append(
            f"В анкете указано авторов: {declared_count}; распознано карточек: {len(authors)}."
        )
    if not authors:
        warnings.append("В разделе I не найдены сведения об авторах.")
    for index, author in enumerate(authors, 1):
        missing = author.missing_fields(
            require_personal_data=authors_will_be_mentioned is not False
        )
        if missing:
            warnings.append(f"Автор {index}: не найдены {', '.join(missing)}.")
    if not applicant_address:
        warnings.append("В разделе I не найден адрес заявителя — его нужно заполнить вручную.")

    data = SoftwareConsentData(
        program_name=program_name,
        applicant_name=applicant_name,
        applicant_address=applicant_address,
        inn=inn,
        authors=authors,
        authors_will_be_mentioned=authors_will_be_mentioned,
        declared_author_count=declared_count,
        source_path=str(source),
    )
    if not data.program_name:
        warnings.insert(0, "В разделе I не найдено название программы.")
    if not data.applicant_name:
        warnings.insert(0, "В разделе I не найдено наименование заявителя.")
    if not data.inn:
        warnings.insert(0, "В разделе I не найден ИНН заявителя.")
    return QuestionnaireParseResult(data=data, warnings=warnings)
