"""Caching helpers for expensive noise directory lookups."""
from __future__ import annotations

import logging
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

logger = logging.getLogger(__name__)

_IMAGE_EXTENSIONS: Tuple[str, ...] = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".tif")


@dataclass
class _FolderListing:
    mtime: float
    entries: List[Path]


class NoiseDirectoryCache:
    """Caches noisy directory resolution and file listings for noise assets."""

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._folder_cache: Dict[Tuple[str, str, str], Optional[str]] = {}
        self._year_cache: Dict[Tuple[str, str], _FolderListing] = {}
        self._file_cache: Dict[str, _FolderListing] = {}

    def clear(self) -> None:
        """Clear all cached entries."""
        with self._lock:
            self._folder_cache.clear()
            self._year_cache.clear()
            self._file_cache.clear()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def resolve_test_folder(self, noise_dir: Path, year: str, test_number: str) -> Optional[Path]:
        """Resolve the folder containing a specific noise test, with caching."""
        root_key = self._root_key(noise_dir)
        cache_key = (root_key, year, test_number)

        with self._lock:
            if cache_key in self._folder_cache:
                cached_path_str = self._folder_cache[cache_key]
                if cached_path_str is None:
                    return None
                cached_path = Path(cached_path_str)
                if cached_path.exists():
                    return cached_path
                # Stale entry – drop and recompute below
                del self._folder_cache[cache_key]

        year_dir = noise_dir / year
        if not year_dir.exists():
            self._store_folder_cache(cache_key, None)
            return None

        # Direct match first
        target_folder = year_dir / test_number
        if target_folder.exists():
            self._store_folder_cache(cache_key, target_folder)
            return target_folder

        # Otherwise search within the year directory (often thousands of entries)
        for candidate in self._list_year_subdirectories(root_key, year_dir):
            if test_number in candidate.name:
                self._store_folder_cache(cache_key, candidate)
                return candidate

        self._store_folder_cache(cache_key, None)
        return None

    def list_image_files(self, folder: Path) -> List[Path]:
        """Return cached list of image files within a folder, refreshing on change."""
        try:
            folder_mtime = folder.stat().st_mtime
        except OSError as exc:
            logger.debug("Cannot stat folder '%s': %s", folder, exc)
            return []

        folder_key = self._folder_key(folder)

        with self._lock:
            cached = self._file_cache.get(folder_key)
            if cached and cached.mtime == folder_mtime:
                return list(cached.entries)

        try:
            entries = [
                item
                for item in folder.iterdir()
                if item.is_file() and item.suffix.lower() in _IMAGE_EXTENSIONS
            ]
        except (OSError, PermissionError) as exc:
            logger.debug("Failed to list files in '%s': %s", folder, exc)
            return []

        with self._lock:
            self._file_cache[folder_key] = _FolderListing(folder_mtime, entries)

        return list(entries)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _list_year_subdirectories(self, root_key: str, year_dir: Path) -> Iterable[Path]:
        try:
            year_mtime = year_dir.stat().st_mtime
        except OSError as exc:
            logger.debug("Cannot stat year directory '%s': %s", year_dir, exc)
            return []

        cache_key = (root_key, year_dir.name)

        with self._lock:
            cached_listing = self._year_cache.get(cache_key)
            if cached_listing and cached_listing.mtime == year_mtime:
                return list(cached_listing.entries)

        try:
            entries = [entry for entry in year_dir.iterdir() if entry.is_dir()]
        except (OSError, PermissionError) as exc:
            logger.debug("Failed to enumerate '%s': %s", year_dir, exc)
            return []

        with self._lock:
            self._year_cache[cache_key] = _FolderListing(year_mtime, entries)

        return list(entries)

    def _store_folder_cache(self, cache_key: Tuple[str, str, str], folder: Optional[Path]) -> None:
        with self._lock:
            self._folder_cache[cache_key] = None if folder is None else str(folder)

    @staticmethod
    def _root_key(noise_dir: Path) -> str:
        try:
            return str(noise_dir.resolve())
        except OSError:
            return str(noise_dir)

    @staticmethod
    def _folder_key(folder: Path) -> str:
        try:
            return str(folder.resolve())
        except OSError:
            return str(folder)


_cache = NoiseDirectoryCache()


def get_noise_directory_cache() -> NoiseDirectoryCache:
    """Return the singleton noise directory cache."""
    return _cache
