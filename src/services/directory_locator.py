"""Shared service for applying directory auto-detection and validation rules."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Optional, Union

from ..config import directory_config as config
from ..config.app_config import AppConfig
from ..utils.common import validate_directory_path


class DirectoryLocator:
    """Provides consolidated access to discovered directories and files."""

    DEFAULT_OUTPUT_FILENAME = "Motor_Performance_Report.xlsx"
    DEFAULT_LOG_FILENAME = "motor_report_refactored.log"

    def __init__(self, *, logger: Optional[logging.Logger] = None, auto_create_output: bool = True) -> None:
        self._logger = logger or logging.getLogger(__name__)
        self._auto_create_output = auto_create_output

    # ------------------------------------------------------------------
    def apply_defaults(self, app_config: AppConfig) -> AppConfig:
        """Populate *app_config* with detected paths while respecting manual overrides."""
        # Project roots
        if not app_config.project_root:
            app_config.project_root = self.project_root
        if not app_config.onedrive_root:
            app_config.onedrive_root = self.onedrive_root

        # Performance data
        app_config.tests_folder = self._resolve_directory(
            provided=app_config.tests_folder,
            default=self.performance_dir,
            label="performance tests folder",
        )
        app_config.performance_test_dir = self.performance_dir

        # Test lab workbook directory (CARICHI NOMINALI)
        app_config.test_lab_root = self._resolve_directory(
            provided=app_config.test_lab_root,
            default=self.test_lab_dir,
            label="test lab workbook folder",
        )

        # Noise data
        app_config.noise_dir = self._resolve_directory(
            provided=app_config.noise_dir,
            default=self.noise_dir,
            label="noise tests folder",
        )

        # Registry files
        app_config.registry_path = self._resolve_file(
            provided=app_config.registry_path,
            default=self.lab_registry_file,
            label="lab registry",
        )
        app_config.lab_registry_file = self.lab_registry_file

        app_config.noise_registry_path = self._resolve_file(
            provided=app_config.noise_registry_path,
            default=self.noise_registry_file,
            label="noise registry",
        )

        # Output / logs
        app_config.output_path = self._resolve_output_path(app_config.output_path)
        app_config.log_path = self._resolve_log_path(app_config.log_path)

        # Logo path (read-only optional)
        app_config.logo_path = self._resolve_file(
            provided=app_config.logo_path,
            default=config.LOGO_PATH,
            label="logo file",
            create_parent=False,
        )

        return app_config

    # ------------------------------------------------------------------
    def _resolve_directory(self, *, provided: Optional[str], default: Optional[Path], label: str) -> Optional[str]:
        if provided:
            validated = validate_directory_path(provided)
            if validated:
                return str(validated)
            self._logger.warning("Invalid %s '%s'; falling back to detected default", label, provided)

        default_path = self._existing_dir(default)
        if default_path:
            return str(default_path)

        if provided:
            self._logger.warning("No default available for %s", label)
        return None

    def _resolve_file(
        self,
        *,
        provided: Optional[str],
        default: Optional[Path],
        label: str,
        create_parent: bool = True,
    ) -> Optional[str]:
        provided_path = self._existing_path(provided)
        if provided_path:
            return str(provided_path)
        if provided and not default:
            self._logger.warning("Provided %s '%s' does not exist", label, provided)
            return None
        if provided and default:
            self._logger.warning("Provided %s '%s' missing – falling back to detected path", label, provided)

        default_path = self._existing_path(default)
        if default_path:
            if create_parent and default_path.parent:
                default_path.parent.mkdir(parents=True, exist_ok=True)
            return str(default_path)
        return None

    def _resolve_output_path(self, provided: Optional[str]) -> Optional[str]:
        if provided:
            target = Path(provided)
            target.parent.mkdir(parents=True, exist_ok=True)
            return str(target)

        base_dir = self.output_dir or config.OUTPUT_DIR
        return self._resolve_from_base_dir(base_dir, self.DEFAULT_OUTPUT_FILENAME)

    def _resolve_log_path(self, provided: Optional[str]) -> Optional[str]:
        if provided:
            target = Path(provided)
            target.parent.mkdir(parents=True, exist_ok=True)
            return str(target)

        base_dir = self.logs_dir or config.LOGS_DIR
        return self._resolve_from_base_dir(base_dir, self.DEFAULT_LOG_FILENAME)

    # ------------------------------------------------------------------
    @property
    def project_root(self) -> Optional[Path]:
        return self._existing_dir(config.PROJECT_ROOT)

    @property
    def onedrive_root(self) -> Optional[Path]:
        return self._existing_dir(config.ONEDRIVE_ROOT)

    @property
    def performance_dir(self) -> Optional[Path]:
        return self._existing_dir(config.PERFORMANCE_TEST_DIR)

    @property
    def noise_dir(self) -> Optional[Path]:
        return self._existing_dir(config.NOISE_TEST_DIR)

    @property
    def lab_registry_file(self) -> Optional[Path]:
        return self._existing_file(config.LAB_REGISTRY_FILE)

    @property
    def noise_registry_file(self) -> Optional[Path]:
        return self._existing_file(config.NOISE_REGISTRY_FILE)

    @property
    def test_lab_dir(self) -> Optional[Path]:
        return self._existing_dir(getattr(config, "TEST_LAB_CARICHI_DIR", None))

    @property
    def output_dir(self) -> Optional[Path]:
        existing = self._existing_dir(config.OUTPUT_DIR)
        if existing:
            return existing
        if self._auto_create_output:
            validated = validate_directory_path(config.OUTPUT_DIR) if config.OUTPUT_DIR else None
            return validated
        return None

    @property
    def logs_dir(self) -> Optional[Path]:
        existing = self._existing_dir(config.LOGS_DIR)
        if existing:
            return existing
        if self._auto_create_output:
            validated = validate_directory_path(config.LOGS_DIR) if config.LOGS_DIR else None
            return validated
        return None

    @property
    def logo_path(self) -> Optional[Path]:
        return self._existing_file(config.LOGO_PATH)

    def snapshot(self) -> Dict[str, Optional[Path]]:
        """Return a dictionary with the most important path values."""
        return {
            "project_root": self.project_root,
            "onedrive_root": self.onedrive_root,
            "performance_dir": self.performance_dir,
            "noise_dir": self.noise_dir,
            "lab_registry_file": self.lab_registry_file,
            "noise_registry_file": self.noise_registry_file,
            "test_lab_dir": self.test_lab_dir,
            "output_dir": self.output_dir,
            "logs_dir": self.logs_dir,
            "logo_path": self.logo_path,
        }

    # ------------------------------------------------------------------
    def _resolve_from_base_dir(self, base_dir: Optional[Union[str, Path]], filename: str) -> Optional[str]:
        base_path = self._ensure_path(base_dir)
        if not base_path:
            return None

        base_path.mkdir(parents=True, exist_ok=True)
        return str(base_path / filename)

    @staticmethod
    def _ensure_path(value: Optional[Union[str, Path]]) -> Optional[Path]:
        if not value:
            return None
        if isinstance(value, Path):
            return value
        try:
            return Path(value)
        except TypeError:
            return None

    def _existing_dir(self, value: Optional[Union[str, Path]]) -> Optional[Path]:
        path = self._ensure_path(value)
        if path and path.exists() and path.is_dir():
            return path
        return None

    def _existing_file(self, value: Optional[Union[str, Path]]) -> Optional[Path]:
        path = self._ensure_path(value)
        if path and path.exists() and path.is_file():
            return path
        return None

    def _existing_path(self, value: Optional[Union[str, Path]]) -> Optional[Path]:
        path = self._ensure_path(value)
        if path and path.exists():
            return path
        return None
