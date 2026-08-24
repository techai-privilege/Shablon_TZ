"""State fees from the approved 1–45 Nice-class reference table."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class FeeResult:
    filing: int
    registration: int
    total: int


# Each tuple is (fee for filing/examination, fee for registration/issuance).
# Values are transcribed from ``Пошлины.docx`` supplied for the application.
REFERENCE_FEES: tuple[tuple[int, int], ...] = (
    (17_000, 18_000),
    (20_500, 18_000),
    (24_000, 18_000),
    (27_500, 18_000),
    (31_000, 18_000),
    (34_500, 20_000),
    (38_000, 22_000),
    (41_500, 24_000),
    (45_000, 26_000),
    (48_500, 28_000),
    (52_000, 30_000),
    (55_500, 32_000),
    (59_000, 34_000),
    (62_500, 36_000),
    (66_000, 38_000),
    (69_500, 40_000),
    (73_000, 42_000),
    (76_500, 44_000),
    (80_000, 46_000),
    (83_500, 48_000),
    (87_000, 50_000),
    (90_500, 52_000),
    (94_000, 54_000),
    (97_500, 56_000),
    (101_000, 58_000),
    (104_500, 60_000),
    (108_000, 62_000),
    (111_500, 64_000),
    (115_000, 66_000),
    (118_500, 68_000),
    (122_000, 70_000),
    (125_500, 72_000),
    (129_000, 74_000),
    (132_500, 76_000),
    (136_000, 78_000),
    (139_500, 80_000),
    (143_000, 82_000),
    (146_500, 84_000),
    (150_000, 86_000),
    (153_500, 88_000),
    (157_000, 90_000),
    (160_500, 92_000),
    (164_000, 94_000),
    (167_500, 96_000),
    (171_000, 98_000),
)


def calculate_fees(class_count: int) -> FeeResult:
    class_count = int(class_count)
    if not 1 <= class_count <= len(REFERENCE_FEES):
        raise ValueError("Количество классов МКТУ должно быть от 1 до 45.")
    filing, registration = REFERENCE_FEES[class_count - 1]
    return FeeResult(filing=filing, registration=registration, total=filing + registration)


def format_rubles(value: int) -> str:
    return f"{value:,}".replace(",", " ")


def class_word(count: int) -> str:
    count = abs(int(count))
    last_two = count % 100
    last = count % 10
    if 11 <= last_two <= 14:
        return "классов"
    if last == 1:
        return "класс"
    if 2 <= last <= 4:
        return "класса"
    return "классов"
