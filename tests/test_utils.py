from __future__ import annotations

from pathlib import Path

from src import utils


def test_normalize_sap_code_trims_and_uppercases():
    assert utils.normalize_sap_code("  ab12 ") == "AB12"
    assert utils.normalize_sap_code("") == ""


def test_validate_directory_path_creates_directory(tmp_path: Path):
    target = tmp_path / "nested" / "dir"
    validated = utils.validate_directory_path(target)

    assert validated == target.resolve()
    assert target.exists()
    assert target.is_dir()


def test_sanitize_filename_removes_invalid_characters():
    name = 're:port*2024?.xlsx'
    sanitized = utils.sanitize_filename(name)
    assert sanitized == "re_port_2024_.xlsx"
