"""Shared registry loading service with thread-safe caching.

This module centralizes the Excel registry access used by both the CLI
workflow and the GUI. It replaces the ad hoc caching previously kept on
`MotorReportApp`, making the data available to any component that
receives a reference to this service.
"""

from __future__ import annotations

import logging
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    import pandas as pd

from ..config.app_config import AppConfig
from ..data.models import Test
from .noise_registry_reader import (
    ANNO_DATETIME_STD,
    ANNO_ORIGINAL_STD,
    CODICE_SAP_STD,
    N_PROVA_STD,
    NOTE_STD,
    VOLTAGE_STD,
    load_registry_dataframe,
)
from ..core.telemetry import log_duration

# Import pandas for runtime use
try:
    import pandas as pd
except ImportError:
    pd = None  # type: ignore

try:
    import pandas as pd  # type: ignore
except Exception:  # pragma: no cover - pandas is an application dependency
    pd = None


@dataclass(frozen=True)
class RegistryCacheKey:
    """Immutable cache key so we can safely compare cached state."""

    path: Path
    sheet_name: Optional[str]
    timestamp: float


class RegistryService:
    """Load and cache registry rows from the Excel data source."""

    def __init__(self, logger: Optional[logging.Logger] = None):
        self._logger = logger or logging.getLogger(__name__)
        self._cache_lock = threading.Lock()
        self._cache_key: Optional[RegistryCacheKey] = None
        self._cache_data: Optional[List[Test]] = None
        # Separate hook for unit tests to swap the dataframe loader.
        self.loader: Callable[..., Tuple[pd.DataFrame, dict]] = staticmethod(load_registry_dataframe)  # type: ignore

    def load_tests(
        self,
        config: AppConfig,
        *,
        sheet_override: Optional[str] = None,
    ) -> List[Test]:
        """Return registry rows as `Test` models with caching.

        Parameters
        ----------
        config:
            Application configuration containing the registry path and
            sheet name.
        sheet_override:
            When provided, forces the loader to use the given sheet name
            instead of the value inside `config`. This is primarily used
            by validators that need to inspect alternate sheets.
        """

        if pd is None:
            raise RuntimeError("pandas is required to load the registry")

        registry_path = self._resolve_registry_path(config)
        if registry_path is None:
            return []

        sheet_name = sheet_override if sheet_override is not None else (config.registry_sheet_name or None)
        timestamp = registry_path.stat().st_mtime

        cached = self._get_cached_data(registry_path, sheet_name, timestamp)
        if cached is not None:
            return cached

        dataframe = self._read_dataframe(registry_path, sheet_name)
        tests = self._convert_dataframe_to_tests(dataframe)

        self._store_cache(registry_path, sheet_name, timestamp, tests)
        return tests

    # ------------------------------------------------------------------
    # Internal helpers

    def _resolve_registry_path(self, config: AppConfig) -> Optional[Path]:
        path_value = config.registry_path
        if not path_value:
            self._logger.error("Registry path is not configured")
            return None

        registry_path = Path(path_value)
        if not registry_path.exists():
            self._logger.error("Registry file not found: %s", registry_path)
            return None
        return registry_path

    def _get_cached_data(
        self,
        path: Path,
        sheet_name: Optional[str],
        timestamp: float,
    ) -> Optional[List[Test]]:
        with self._cache_lock:
            if self._cache_key == RegistryCacheKey(path=path, sheet_name=sheet_name, timestamp=timestamp):
                self._logger.debug("Using cached registry data")
                # Return a shallow copy so callers cannot mutate the cache
                return list(self._cache_data or [])
        return None

    def _store_cache(
        self,
        path: Path,
        sheet_name: Optional[str],
        timestamp: float,
        tests: Iterable[Test],
    ) -> None:
        with self._cache_lock:
            self._cache_key = RegistryCacheKey(path=path, sheet_name=sheet_name, timestamp=timestamp)
            self._cache_data = list(tests)
            self._logger.info("Loaded and cached %d tests from registry", len(self._cache_data))

    def _read_dataframe(self, registry_path: Path, sheet_name: Optional[str]):
        self._logger.info("Loading registry from: %s", registry_path)
        try:
            with log_duration(
                self._logger,
                "registry_dataframe_read",
                level=logging.DEBUG,
                extra={"registry_path": str(registry_path), "sheet": sheet_name},
            ):
                dataframe, _ = self.loader(
                    registry_path,
                    sheet_name=sheet_name,
                    log=self._logger,
                )
        except ValueError as exc:
            self._logger.warning(
                "Sheet '%s' not found in registry (%s); falling back to first sheet",
                sheet_name,
                exc,
            )
            with log_duration(
                self._logger,
                "registry_dataframe_read_fallback",
                level=logging.DEBUG,
                extra={"registry_path": str(registry_path)},
            ):
                dataframe, _ = self.loader(
                    registry_path,
                    sheet_name=0,
                    log=self._logger,
                )
        return dataframe

    def _convert_dataframe_to_tests(self, dataframe) -> List[Test]:
        if dataframe.empty:
            self._logger.warning("Registry file contained no usable data after cleaning")
            return []

        tests: List[Test] = []
        for _, row in dataframe.iterrows():
            test_lab_number = str(row.get(N_PROVA_STD, "")).strip()
            if not test_lab_number or test_lab_number == "nan":
                continue

            sap_code = str(row.get(CODICE_SAP_STD, "")).strip()
            voltage = str(row.get(VOLTAGE_STD, "")).strip() if VOLTAGE_STD in row else ""
            notes = str(row.get(NOTE_STD, "")).strip() if NOTE_STD in row else ""

            date_val = row.get(ANNO_DATETIME_STD) or row.get(ANNO_ORIGINAL_STD, "")
            if pd is not None and pd.isna(date_val):
                date_info = ""
            elif hasattr(date_val, "strftime"):
                date_info = date_val.strftime("%Y-%m-%d")
            else:
                date_info = str(date_val)

            tests.append(
                Test(
                    test_lab_number=test_lab_number,
                    sap_code=sap_code,
                    voltage=voltage,
                    notes=notes,
                    date=date_info,
                )
            )

        return tests
