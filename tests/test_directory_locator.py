from __future__ import annotations

from pathlib import Path

import pytest

from src.config.app_config import AppConfig
from src.services.directory_locator import DirectoryLocator


def test_apply_defaults_populates_missing_fields(directory_config_stub):
    locator = DirectoryLocator()
    config = AppConfig()

    result = locator.apply_defaults(config)

    assert result.tests_folder == str(directory_config_stub["performance"])
    assert result.performance_test_dir == directory_config_stub["performance"]
    assert result.noise_dir == str(directory_config_stub["noise"])

    assert result.registry_path == str(directory_config_stub["lab_registry"])
    assert result.lab_registry_file == directory_config_stub["lab_registry"]

    assert result.noise_registry_path == str(directory_config_stub["noise_registry"])

    expected_output = directory_config_stub["output"] / DirectoryLocator.DEFAULT_OUTPUT_FILENAME
    assert result.output_path == str(expected_output)

    expected_log = directory_config_stub["logs"] / DirectoryLocator.DEFAULT_LOG_FILENAME
    assert result.log_path == str(expected_log)

    assert result.logo_path == str(directory_config_stub["logo"])
    assert result.project_root == directory_config_stub["project"]
    assert result.onedrive_root == directory_config_stub["onedrive"]


def test_apply_defaults_preserves_user_overrides(directory_config_stub, tmp_path: Path):
    locator = DirectoryLocator()

    custom_tests = tmp_path / "custom_tests"
    custom_tests.mkdir()
    custom_noise = tmp_path / "custom_noise"
    custom_noise.mkdir()
    custom_registry = tmp_path / "custom_registry.xlsx"
    custom_registry.write_text("data")
    custom_output = tmp_path / "reports" / "report.xlsx"
    custom_output.parent.mkdir(parents=True, exist_ok=True)
    custom_log = tmp_path / "logs" / "custom.log"
    custom_log.parent.mkdir(parents=True, exist_ok=True)
    custom_logo = tmp_path / "brand" / "logo.png"
    custom_logo.parent.mkdir(parents=True, exist_ok=True)
    custom_logo.write_bytes(b"logo")

    config = AppConfig(
        tests_folder=str(custom_tests),
        noise_dir=str(custom_noise),
        registry_path=str(custom_registry),
        output_path=str(custom_output),
        log_path=str(custom_log),
        logo_path=str(custom_logo),
    )

    locator.apply_defaults(config)

    assert config.tests_folder == str(custom_tests)
    assert config.noise_dir == str(custom_noise)
    assert config.registry_path == str(custom_registry)
    assert config.output_path == str(custom_output)
    assert config.log_path == str(custom_log)
    assert config.logo_path == str(custom_logo)


def test_snapshot_reflects_current_paths(directory_config_stub):
    locator = DirectoryLocator()

    snapshot = locator.snapshot()

    assert snapshot["performance_dir"] == directory_config_stub["performance"]
    assert snapshot["noise_dir"] == directory_config_stub["noise"]
    assert snapshot["lab_registry_file"] == directory_config_stub["lab_registry"]
    assert snapshot["noise_registry_file"] == directory_config_stub["noise_registry"]
    assert snapshot["output_dir"] == directory_config_stub["output"]
    assert snapshot["logs_dir"] == directory_config_stub["logs"]
    assert snapshot["logo_path"] == directory_config_stub["logo"]
