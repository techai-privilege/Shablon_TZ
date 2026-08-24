import pytest

from trademark_report.fees import REFERENCE_FEES, calculate_fees


def test_one_class_fees_match_reference():
    result = calculate_fees(1)
    assert result.filing == 17_000
    assert result.registration == 18_000
    assert result.total == 35_000


def test_ten_class_fees_match_sauluk_reference():
    result = calculate_fees(10)
    assert result.filing == 48_500
    assert result.registration == 28_000
    assert result.total == 76_500


def test_all_reference_rows_cover_1_to_45_classes():
    assert len(REFERENCE_FEES) == 45
    assert calculate_fees(45).filing == 171_000
    assert calculate_fees(45).registration == 98_000


def test_rejects_class_count_outside_reference_table():
    with pytest.raises(ValueError):
        calculate_fees(46)
