"""Utilities to resolve Carichi Nominali ("A" test) workbooks."""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional

from .test_lab_summary_loader import TestLabSummaryLoader, TestLabWorkbookMatch

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class CarichiTestInfo:
    """Information about a resolved Carichi Nominali workbook."""

    requested_test_number: str
    matched_test_number: str
    path: Path
    year_folder: Optional[str]
    match_strategy: str


class CarichiLocator:
    """Resolve Carichi Nominali workbooks that mirror performance tests."""

    __test__ = False  # prevent pytest from discovering this as a test case

    def __init__(self, base_path: str | Path | None, *, log: Optional[logging.Logger] = None) -> None:
        self._logger = log or logger
        self._loader = TestLabSummaryLoader(str(base_path) if base_path else None, logger_=self._logger)
        self._cache: Dict[str, Optional[CarichiTestInfo]] = {}

    # ------------------------------------------------------------------
    @property
    def available(self) -> bool:
        return self._loader.available

    # ------------------------------------------------------------------
    def find(self, test_number: str) -> Optional[CarichiTestInfo]:
        """Return Carichi info for ``test_number`` (with or without trailing "A")."""
        normalized = (test_number or "").strip().upper()
        if not normalized:
            return None

        cache_key = normalized
        if cache_key in self._cache:
            return self._cache[cache_key]

        match = self._loader.locate_workbook(normalized)
        # Strict matching: do not try to append "A" if not found
        # if not match and not normalized.endswith("A"):
        #     match = self._loader.locate_workbook(f"{normalized}A")

        if not match:
            self._cache[cache_key] = None
            return None

        info = self._to_info(match)
        self._cache[cache_key] = info
        return info

    # ------------------------------------------------------------------
    def find_for_performance_test(self, performance_test_number: str) -> Optional[CarichiTestInfo]:
        """Resolve the mirrored "A" workbook for a base performance test."""
        normalized = (performance_test_number or "").strip().upper()
        if not normalized:
            return None
        # Strict matching: Do not automatically append 'A'.
        # Only find if the test number explicitly matches a Carichi workbook.
        return self.find(normalized)

    # ------------------------------------------------------------------
    def bulk_lookup(self, test_numbers: Iterable[str]) -> Dict[str, Optional[CarichiTestInfo]]:
        """Resolve multiple test numbers in one pass, leveraging the cache."""
        result: Dict[str, Optional[CarichiTestInfo]] = {}
        for number in test_numbers:
            info = self.find(number)
            result[number] = info
        return result

    # ------------------------------------------------------------------
    def _to_info(self, match: TestLabWorkbookMatch) -> CarichiTestInfo:
        return CarichiTestInfo(
            requested_test_number=match.requested_test_number,
            matched_test_number=match.matched_test_number,
            path=match.path,
            year_folder=match.year_folder,
            match_strategy=match.match_strategy,
        )
