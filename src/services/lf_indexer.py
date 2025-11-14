"""
Lightweight LF file indexer.

Builds a cache of LF files under the LF base directory and maps probable
test IDs (normalized to 3 digits) to file paths. The index is stored in a
JSON cache file and refreshed automatically based on TTL or on-demand.

This makes repeated lookups (from the GUI) fast and robust as the
LF repository grows.
"""
from __future__ import annotations

import json
import logging
import re
import time
from pathlib import Path
from typing import Dict, List, Optional

DEFAULT_EXTENSIONS = {'.xls', '.xlsx', '.xlsm', '.xlsb'}


class LFIndex:
    """Index LF files under a base directory for fast lookup.

    Basic behavior:
    - Scan base_dir for LF files and extract probable test IDs and years
      from filenames.
    - Store mapping test_id -> list of candidate entries
    - Cache the index to disk (.lf_index_cache.json by default)
    - Provide methods to query the best candidate for a (test_id, year)
      pair and to refresh the cache on demand.
    """

    def __init__(self, base_dir: Path, cache_file: Optional[Path] = None,
                 rebuild_interval: int = 3600, logger: Optional[logging.Logger] = None,
                 background: bool = False):
        self.base_dir = Path(base_dir) if base_dir else None
        self.logger = logger or logging.getLogger(__name__)
        self.rebuild_interval = rebuild_interval

        if cache_file:
            self.cache_file = Path(cache_file)
        else:
            # Write cache inside base_dir when possible so it's colocated
            self.cache_file = (self.base_dir / '.lf_index_cache.json') if self.base_dir else Path('.lf_index_cache.json')

        self._index: Dict[str, List[Dict]] = {}
        self._metadata: Dict = {}
        self._last_built = 0.0

        # Lazy load cache; optionally start a background build if required
        self._load_cache()
        # If index appears stale, optionally kick off a background build
        if background and self.is_stale():
            try:
                import threading
                t = threading.Thread(target=self.build_index, kwargs={'force': True}, daemon=True)
                t.start()
                self.logger.debug("Started background LF index build")
            except Exception as e:
                self.logger.debug(f"Failed to start background LF index build: {e}")

    # --- Internal helpers ---
    def _load_cache(self) -> None:
        if not self.cache_file.exists():
            self.logger.debug("LF index cache not found, will build on demand")
            return
        try:
            data = json.loads(self.cache_file.read_text(encoding='utf-8'))
            self._index = data.get('index', {})
            self._metadata = data.get('metadata', {})
            self._last_built = self._metadata.get('generated_on', 0)
            self.logger.debug(f"Loaded LF index cache: {self.cache_file} (entries={sum(len(v) for v in self._index.values())})")
        except Exception as e:
            self.logger.warning(f"Failed to load LF index cache: {e}")

    def _save_cache(self) -> None:
        try:
            payload = {'metadata': self._metadata, 'index': self._index}
            self.cache_file.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding='utf-8')
            self.logger.debug(f"Saved LF index cache: {self.cache_file}")
        except Exception as e:
            self.logger.warning(f"Failed to save LF index cache: {e}")

    @staticmethod
    def _normalize_id(test_id: str) -> str:
        try:
            return str(int(test_id)).zfill(3)
        except Exception:
            # Fall back to zero-padded digits if possible
            digits = re.sub(r'\D', '', str(test_id) or '')
            if digits:
                return digits.zfill(3)
            return str(test_id)

    @staticmethod
    def _normalize_year(year_token: Optional[str]) -> Optional[str]:
        if not year_token:
            return None
        try:
            y = int(year_token)
            if y < 100:  # 2-digit
                return f"19{y:02d}" if y >= 97 else f"20{y:02d}"
            return str(y)
        except Exception:
            return None

    def _extract_possible_ids(self, filename: str) -> List[Dict]:
        """Extract possible (test_id, year) tuples from a filename.

        Returns a list of candidate dicts: {'id': '053', 'year': '2018'}
        """
        stem = Path(filename).stem
        candidates = []

        # Strategy 1: explicit LF pattern: LF 053-18, LF053/18, LF_053-18
        m = re.search(r'LF\D*(\d{1,4})(?:\D+[/-]?(\d{2,4}))?', stem, flags=re.IGNORECASE)
        if m:
            test_id = self._normalize_id(m.group(1))
            year = self._normalize_year(m.group(2)) if m.group(2) else None
            candidates.append({'id': test_id, 'year': year})

        # Strategy 2: generic digits/yy pattern anywhere (e.g., '056-23')
        if not candidates:
            m2 = re.search(r'(\d{1,4})[\s_\-]*/?[\s_\-]?(\d{2,4})', stem)
            if m2:
                test_id = self._normalize_id(m2.group(1))
                year = self._normalize_year(m2.group(2))
                candidates.append({'id': test_id, 'year': year})

        # Strategy 3: fallback - any digit group likely to be an id
        if not candidates:
            for d in re.findall(r'\d{1,4}', stem):
                candidates.append({'id': self._normalize_id(d), 'year': None})

        return candidates

    # --- Public operations ---
    def build_index(self, force: bool = False) -> None:
        """Scan the base_dir recursively and build the mapping.

        If force=False the method will still rebuild if the cache is stale.
        """
        if not self.base_dir or not self.base_dir.exists():
            self.logger.warning(f"LF base dir not available for indexing: {self.base_dir}")
            return

        # Quick skip if recently built
        now = time.time()
        if not force and (now - self._last_built) < self.rebuild_interval and self._index:
            self.logger.debug("LF index is fresh; skipping rebuild")
            return

        self.logger.info(f"Building LF index for: {self.base_dir}")
        index: Dict[str, List[Dict]] = {}
        file_count = 0
        max_mtime = 0.0

        for p in self.base_dir.rglob('*'):
            if not p.is_file():
                continue
            if p.suffix.lower() not in DEFAULT_EXTENSIONS:
                continue
            file_count += 1
            mlist = self._extract_possible_ids(p.name)
            for cand in mlist:
                entry = {
                    'path': str(p.resolve()),
                    'name': p.name,
                    'year': cand.get('year'),
                    'mtime': p.stat().st_mtime,
                }
                try:
                    mtime_val = p.stat().st_mtime
                    if mtime_val and mtime_val > max_mtime:
                        max_mtime = mtime_val
                except Exception:
                    pass
                key = cand.get('id')
                if not key:
                    continue
                index.setdefault(key, []).append(entry)

        self._index = index
        self._last_built = now
        self._metadata = {'generated_on': now, 'file_count': file_count, 'max_mtime': max_mtime}
        self._save_cache()
        self.logger.info(f"LF index built: {file_count} files indexed, {len(self._index)} unique ids")

    def _scan_metrics(self) -> tuple[int, float]:
        """Scan base_dir to compute quick metrics used to detect changes.

        Returns (file_count, max_mtime) for known extensions.
        """
        # Guard against missing base dir
        if not self.base_dir or not self.base_dir.exists():
            return 0, 0.0

        file_count = 0
        max_mtime = 0.0
        try:
            for p in self.base_dir.rglob('*'):
                if not p.is_file():
                    continue
                if p.suffix.lower() not in DEFAULT_EXTENSIONS:
                    continue
                file_count += 1
                try:
                    m = p.stat().st_mtime
                    if m > max_mtime:
                        max_mtime = m
                except Exception:
                    continue
        except Exception as e:
            self.logger.debug(f"Failed to scan metrics for LF index: {e}")
        return file_count, max_mtime

    def is_stale(self) -> bool:
        """Return True if the index appears stale (files changed or TTL expired)."""
        # If no metadata present consider stale
        if not self._metadata:
            return True
        now = time.time()
        generated = self._metadata.get('generated_on', 0)
        if (now - generated) > self.rebuild_interval:
            return True
        # Do a quick metrics scan to detect added/removed/modified files
        file_count, max_mtime = self._scan_metrics()
        if file_count != self._metadata.get('file_count', 0):
            return True
        if max_mtime > self._metadata.get('max_mtime', 0):
            return True
        return False

    def refresh_if_stale(self) -> None:
        """Rebuild index when stale using file metrics and TTL."""
        if self.is_stale():
            self.build_index(force=True)

    def get_candidates(self, test_id: str) -> List[Dict]:
        if not test_id:
            return []
        key = self._normalize_id(test_id)
        return self._index.get(key, [])

    def get_best_file(self, test_id: str, year: Optional[str] = None) -> Optional[Path]:
        """Return the best candidate Path for a (test_id, year) pair, or None."""
        if not test_id:
            return None
        # ensure index is reasonably fresh
        self.refresh_if_stale()

        key = self._normalize_id(test_id)
        candidates = self._index.get(key, [])
        if not candidates:
            return None

        # Prefer exact matches that include 'LF' and the id in the filename
        for c in candidates:
            name_upper = c.get('name', '').upper()
            if re.search(rf'\bLF\D*{re.escape(key)}\b', name_upper):
                return Path(c['path'])

        # If year specified prefer same-year entries
        if year:
            for c in candidates:
                if c.get('year') == year:
                    return Path(c['path'])

        # Otherwise return the newest candidate (by mtime)
        candidates_sorted = sorted(candidates, key=lambda x: x.get('mtime', 0), reverse=True)
        return Path(candidates_sorted[0]['path']) if candidates_sorted else None
