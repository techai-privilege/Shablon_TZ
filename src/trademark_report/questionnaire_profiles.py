"""Load and apply declarative profiles for software questionnaires.

Profiles describe stable questionnaire structure and labels.  They deliberately
contain no client values or personal data; filled documents remain input only.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
import json
from pathlib import Path
import re
import sys
from typing import Iterable, Mapping


def normalize_label(value: str) -> str:
    """Normalize a Word label before matching it against a profile."""

    return re.sub(
        r"[^а-яёa-z0-9]+", " ", value.lower().replace("ё", "е")
    ).strip()


def _resource_root() -> Path:
    bundle_root = getattr(sys, "_MEIPASS", None)
    if getattr(sys, "frozen", False) and bundle_root:
        return Path(bundle_root)
    return Path(__file__).resolve().parents[2]


DEFAULT_PROFILE_PATH = _resource_root() / "data" / "questionnaire_profiles.json"


@dataclass(frozen=True, slots=True)
class FieldRule:
    labels: tuple[str, ...]
    match: str = "contains"

    def matches(self, label: str) -> bool:
        normalized = normalize_label(label).rstrip(":")
        if self.match == "exact":
            return normalized in self.labels
        return any(candidate in normalized for candidate in self.labels)


@dataclass(frozen=True, slots=True)
class RowKindRule:
    any_terms: tuple[str, ...] = ()
    all_terms: tuple[str, ...] = ()

    def matches(self, label: str) -> bool:
        normalized = normalize_label(label)
        return (
            (not self.any_terms or any(term in normalized for term in self.any_terms))
            and all(term in normalized for term in self.all_terms)
        )


@dataclass(frozen=True, slots=True)
class QuestionnaireProfile:
    profile_id: str
    name: str
    priority: int
    match_all: tuple[str, ...] = ()
    match_any: tuple[str, ...] = ()
    match_none: tuple[str, ...] = ()
    fallback: bool = False
    field_overrides: Mapping[str, FieldRule] = field(default_factory=dict)
    row_kind_overrides: Mapping[str, RowKindRule] = field(default_factory=dict)
    section_start_any: tuple[str, ...] = ()
    section_stop_any: tuple[str, ...] = ()

    def score(self, labels: Iterable[str]) -> int | None:
        normalized = tuple(normalize_label(label) for label in labels)

        def present(anchor: str) -> bool:
            return any(anchor in label for label in normalized)

        if self.match_all and not all(present(anchor) for anchor in self.match_all):
            return None
        if self.match_any and not any(present(anchor) for anchor in self.match_any):
            return None
        if any(present(anchor) for anchor in self.match_none):
            return None
        if self.fallback:
            return self.priority
        return self.priority + len(self.match_all) * 10 + sum(
            present(anchor) for anchor in self.match_any
        )


@dataclass(frozen=True, slots=True)
class QuestionnaireProfileLibrary:
    schema_version: int
    section_start_any: tuple[str, ...]
    section_stop_any: tuple[str, ...]
    fields: Mapping[str, FieldRule]
    row_kinds: Mapping[str, RowKindRule]
    profiles: tuple[QuestionnaireProfile, ...]

    def select_profile(self, labels: Iterable[str]) -> QuestionnaireProfile:
        candidates = [
            (score, profile)
            for profile in self.profiles
            if (score := profile.score(labels)) is not None
        ]
        if not candidates:
            raise ValueError("В библиотеке нет универсального профиля анкеты.")
        return max(candidates, key=lambda item: item[0])[1]

    def field_rule(self, profile: QuestionnaireProfile, field_name: str) -> FieldRule:
        try:
            return profile.field_overrides.get(field_name, self.fields[field_name])
        except KeyError as exc:
            raise ValueError(f"В библиотеке отсутствует описание поля: {field_name}") from exc

    def row_matches(
        self, profile: QuestionnaireProfile, kind: str, label: str
    ) -> bool:
        try:
            return profile.row_kind_overrides.get(kind, self.row_kinds[kind]).matches(label)
        except KeyError as exc:
            raise ValueError(f"В библиотеке отсутствует тип строки: {kind}") from exc

    def starts_section(self, text: str, profile: QuestionnaireProfile | None = None) -> bool:
        normalized = normalize_label(text)
        anchors = (
            profile.section_start_any
            if profile is not None and profile.section_start_any
            else self.section_start_any
        )
        return any(anchor in normalized for anchor in anchors)

    def stops_section(self, text: str, profile: QuestionnaireProfile | None = None) -> bool:
        normalized = normalize_label(text)
        anchors = (
            profile.section_stop_any
            if profile is not None and profile.section_stop_any
            else self.section_stop_any
        )
        return any(anchor in normalized for anchor in anchors)


def _string_tuple(value: object, *, context: str) -> tuple[str, ...]:
    if value is None:
        return ()
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{context} должен быть списком строк.")
    return tuple(normalize_label(item) for item in value if normalize_label(item))


def _field_rule(value: object, *, context: str) -> FieldRule:
    if not isinstance(value, dict):
        raise ValueError(f"{context} должен быть объектом.")
    match = value.get("match", "contains")
    if match not in {"contains", "exact"}:
        raise ValueError(f"{context}.match должен быть contains или exact.")
    labels = _string_tuple(value.get("labels"), context=f"{context}.labels")
    if not labels:
        raise ValueError(f"{context}.labels не может быть пустым.")
    return FieldRule(labels=labels, match=match)


def load_profile_library(
    path: str | Path = DEFAULT_PROFILE_PATH,
) -> QuestionnaireProfileLibrary:
    """Load and validate a questionnaire-profile library from one JSON file."""

    source = Path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Не удалось прочитать библиотеку анкет {source}: {exc}") from exc
    if not isinstance(payload, dict) or payload.get("schema_version") != 1:
        raise ValueError("Поддерживается только schema_version 1 библиотеки анкет.")

    section = payload.get("section")
    if not isinstance(section, dict):
        raise ValueError("В библиотеке отсутствует раздел section.")
    section_start = _string_tuple(section.get("start_any"), context="section.start_any")
    section_stop = _string_tuple(section.get("stop_any"), context="section.stop_any")
    if not section_start or not section_stop:
        raise ValueError("Границы раздела I не могут быть пустыми.")

    raw_fields = payload.get("fields")
    if not isinstance(raw_fields, dict):
        raise ValueError("В библиотеке отсутствует раздел fields.")
    fields = {
        str(name): _field_rule(value, context=f"fields.{name}")
        for name, value in raw_fields.items()
    }

    raw_row_kinds = payload.get("row_kinds")
    if not isinstance(raw_row_kinds, dict):
        raise ValueError("В библиотеке отсутствует раздел row_kinds.")
    row_kinds: dict[str, RowKindRule] = {}
    for name, value in raw_row_kinds.items():
        if not isinstance(value, dict):
            raise ValueError(f"row_kinds.{name} должен быть объектом.")
        rule = RowKindRule(
            any_terms=_string_tuple(value.get("any"), context=f"row_kinds.{name}.any"),
            all_terms=_string_tuple(value.get("all"), context=f"row_kinds.{name}.all"),
        )
        if not rule.any_terms and not rule.all_terms:
            raise ValueError(f"row_kinds.{name} не содержит условий.")
        row_kinds[str(name)] = rule

    raw_profiles = payload.get("profiles")
    if not isinstance(raw_profiles, list) or not raw_profiles:
        raise ValueError("Библиотека должна содержать хотя бы один профиль.")
    profiles = []
    seen_ids = set()
    for index, value in enumerate(raw_profiles):
        context = f"profiles[{index}]"
        if not isinstance(value, dict):
            raise ValueError(f"{context} должен быть объектом.")
        profile_id = str(value.get("id", "")).strip()
        name = str(value.get("name", "")).strip()
        if not profile_id or not name or profile_id in seen_ids:
            raise ValueError(f"{context} содержит пустой или повторяющийся id/name.")
        seen_ids.add(profile_id)
        match = value.get("match") or {}
        if not isinstance(match, dict):
            raise ValueError(f"{context}.match должен быть объектом.")
        raw_overrides = value.get("field_overrides") or {}
        if not isinstance(raw_overrides, dict):
            raise ValueError(f"{context}.field_overrides должен быть объектом.")
        overrides = {
            str(field_name): _field_rule(
                rule, context=f"{context}.field_overrides.{field_name}"
            )
            for field_name, rule in raw_overrides.items()
        }
        unknown_fields = set(overrides) - set(fields)
        if unknown_fields:
            raise ValueError(
                f"{context}.field_overrides содержит неизвестные поля: "
                f"{', '.join(sorted(unknown_fields))}."
            )
        raw_row_overrides = value.get("row_kind_overrides") or {}
        if not isinstance(raw_row_overrides, dict):
            raise ValueError(f"{context}.row_kind_overrides должен быть объектом.")
        row_overrides = {}
        for kind, rule_value in raw_row_overrides.items():
            if not isinstance(rule_value, dict):
                raise ValueError(f"{context}.row_kind_overrides.{kind} должен быть объектом.")
            row_overrides[str(kind)] = RowKindRule(
                any_terms=_string_tuple(
                    rule_value.get("any"),
                    context=f"{context}.row_kind_overrides.{kind}.any",
                ),
                all_terms=_string_tuple(
                    rule_value.get("all"),
                    context=f"{context}.row_kind_overrides.{kind}.all",
                ),
            )
            if not (
                row_overrides[str(kind)].any_terms
                or row_overrides[str(kind)].all_terms
            ):
                raise ValueError(
                    f"{context}.row_kind_overrides.{kind} не содержит условий."
                )
        unknown_row_kinds = set(row_overrides) - set(row_kinds)
        if unknown_row_kinds:
            raise ValueError(
                f"{context}.row_kind_overrides содержит неизвестные типы: "
                f"{', '.join(sorted(unknown_row_kinds))}."
            )
        profile_section = value.get("section") or {}
        if not isinstance(profile_section, dict):
            raise ValueError(f"{context}.section должен быть объектом.")
        profiles.append(
            QuestionnaireProfile(
                profile_id=profile_id,
                name=name,
                priority=int(value.get("priority", 0)),
                match_all=_string_tuple(match.get("all"), context=f"{context}.match.all"),
                match_any=_string_tuple(match.get("any"), context=f"{context}.match.any"),
                match_none=_string_tuple(match.get("none"), context=f"{context}.match.none"),
                fallback=bool(value.get("fallback", False)),
                field_overrides=overrides,
                row_kind_overrides=row_overrides,
                section_start_any=_string_tuple(
                    profile_section.get("start_any"), context=f"{context}.section.start_any"
                ),
                section_stop_any=_string_tuple(
                    profile_section.get("stop_any"), context=f"{context}.section.stop_any"
                ),
            )
        )
    if not any(profile.fallback for profile in profiles):
        raise ValueError("Библиотека должна содержать fallback-профиль.")

    return QuestionnaireProfileLibrary(
        schema_version=1,
        section_start_any=section_start,
        section_stop_any=section_stop,
        fields=fields,
        row_kinds=row_kinds,
        profiles=tuple(profiles),
    )


@lru_cache(maxsize=1)
def default_profile_library() -> QuestionnaireProfileLibrary:
    return load_profile_library(DEFAULT_PROFILE_PATH)
