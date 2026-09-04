from pathlib import Path

import pytest

from trademark_report.io_utils import atomic_output_path, atomic_write_bytes


def test_atomic_write_replaces_complete_file(tmp_path):
    destination = tmp_path / "report.docx"
    destination.write_bytes(b"old")

    atomic_write_bytes(destination, b"new")

    assert destination.read_bytes() == b"new"
    assert not list(tmp_path.glob(".report.docx.*.tmp"))


def test_atomic_output_keeps_old_file_on_generation_error(tmp_path):
    destination = tmp_path / "report.docx"
    destination.write_bytes(b"old")

    with pytest.raises(RuntimeError):
        with atomic_output_path(destination) as temporary:
            Path(temporary).write_bytes(b"partial")
            raise RuntimeError("generation failed")

    assert destination.read_bytes() == b"old"
    assert not list(tmp_path.glob(".report.docx.*.tmp"))
