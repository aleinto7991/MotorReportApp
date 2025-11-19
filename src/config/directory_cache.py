"""
Directory cache system to avoid repeated slow directory searches.
Stores found directories in a JSON file for quick retrieval.
"""

import json
import os
import tempfile
import time
from pathlib import Path
from typing import Dict, List, Optional
import logging

from .runtime import get_user_data_dir

logger = logging.getLogger(__name__)

class DirectoryCache:
    """Manages cached directory paths to avoid repeated searches.

    By default the cache is stored in a per-user application data directory
    (see `get_user_data_dir`). An explicit absolute `cache_file` path or the
    environment variable `MOTOR_REPORT_CACHE_FILE` can override this behavior.
    """

    def __init__(self, cache_file: Optional[str] = None, user_data_dir: Optional[Path] = None):
        """Initialize directory cache.

        Args:
            cache_file: Name or path of the cache file. If None, defaults to
                        'directory_cache.json' under the per-user data dir.
            user_data_dir: Optional Path to use as the base directory for the cache.
        """
        # Determine effective cache file path with precedence:
        # 1. Explicit absolute path passed in `cache_file`
        # 2. Environment variable MOTOR_REPORT_CACHE_FILE
        # 3. Per-user data dir (get_user_data_dir) + 'directory_cache.json'
        # 4. Fallback to project-relative file (legacy behavior)
        default_name = "directory_cache.json"
        env_override = os.environ.get('MOTOR_REPORT_CACHE_FILE')

        if cache_file and Path(cache_file).is_absolute():
            self.cache_file = Path(cache_file)
        elif env_override:
            self.cache_file = Path(env_override)
        else:
            # Choose base user data dir
            base_dir = Path(user_data_dir) if user_data_dir else get_user_data_dir()
            try:
                base_dir.mkdir(parents=True, exist_ok=True)
            except Exception:
                # If we cannot create the user dir, fall back to project-relative
                base_dir = Path(__file__).parent.parent
            name = cache_file if cache_file else default_name
            self.cache_file = Path(base_dir) / name

        # Ensure parent directory exists
        try:
            self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            logger.debug(f"Could not ensure parent dir for cache file: {self.cache_file.parent}")

        self._cache = self._load_cache()
    
    def _load_cache(self) -> Dict:
        """Load cache from file."""
        try:
            if self.cache_file.exists():
                try:
                    with open(self.cache_file, 'r', encoding='utf-8') as f:
                        cache = json.load(f)
                        logger.info(f"Loaded directory cache with {len(cache)} entries from {self.cache_file}")
                        return cache
                except (json.JSONDecodeError, ValueError) as jde:
                    # Corrupt cache file - move it aside and start fresh
                    ts = int(time.time())
                    corrupt_name = f"{self.cache_file.name}.corrupt-{ts}"
                    corrupt_path = self.cache_file.parent / corrupt_name
                    try:
                        self.cache_file.replace(corrupt_path)
                        logger.warning(f"Corrupt directory cache renamed to {corrupt_path}")
                    except Exception:
                        logger.warning(f"Could not rename corrupt cache file: {self.cache_file}")
                    return {}
        except Exception as e:
            logger.warning(f"Could not load directory cache: {e}")

        return {}
    
    def _save_cache(self):
        """Save cache to file."""
        try:
            # Atomic write: write to temp file in same directory then replace
            parent = self.cache_file.parent
            parent.mkdir(parents=True, exist_ok=True)
            fd, tmp_path = tempfile.mkstemp(prefix=f"{self.cache_file.stem}.", suffix=".tmp", dir=str(parent))
            try:
                with os.fdopen(fd, 'w', encoding='utf-8') as tmpf:
                    json.dump(self._cache, tmpf, indent=2, ensure_ascii=False)
                # Replace target atomically
                os.replace(tmp_path, str(self.cache_file))
                logger.info(f"Saved directory cache with {len(self._cache)} entries to {self.cache_file}")
            finally:
                # If tmp file still exists, attempt to remove
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
        except Exception as e:
            logger.error(f"Could not save directory cache: {e}")
    
    def get_registry_directories(self) -> List[str]:
        """Get cached registry directories."""
        return self._cache.get('registry_directories', [])
    
    def get_inf_directories(self) -> List[str]:
        """Get cached INF directories."""
        return self._cache.get('inf_directories', [])
    
    def set_registry_directories(self, directories: List[str]):
        """Set registry directories in cache."""
        # Filter out non-existent directories
        valid_dirs = [d for d in directories if os.path.exists(d)]
        self._cache['registry_directories'] = valid_dirs
        self._save_cache()
        logger.info(f"Cached {len(valid_dirs)} registry directories")
    
    def set_inf_directories(self, directories: List[str]):
        """Set INF directories in cache."""
        # Filter out non-existent directories
        valid_dirs = [d for d in directories if os.path.exists(d)]
        self._cache['inf_directories'] = valid_dirs
        self._save_cache()
        logger.info(f"Cached {len(valid_dirs)} INF directories")
    
    def is_cache_valid(self) -> bool:
        """Check if cache has valid entries and directories still exist."""
        registry_dirs = self.get_registry_directories()
        inf_dirs = self.get_inf_directories()
        
        # Check if we have cached directories
        if not registry_dirs and not inf_dirs:
            return False
        
        # Check if at least some cached directories still exist
        valid_registry = any(os.path.exists(d) for d in registry_dirs)
        valid_inf = any(os.path.exists(d) for d in inf_dirs)
        
        return valid_registry or valid_inf
    
    def invalidate_cache(self):
        """Clear the cache and delete cache file."""
        self._cache.clear()
        try:
            if self.cache_file.exists():
                self.cache_file.unlink()
                logger.info("Directory cache invalidated")
        except Exception as e:
            logger.error(f"Could not delete cache file: {e}")
    
    def add_directory(self, directory: str, dir_type: str):
        """Add a single directory to cache.
        
        Args:
            directory: Directory path to add
            dir_type: 'registry' or 'inf'
        """
        if not os.path.exists(directory):
            return
        
        cache_key = f"{dir_type}_directories"
        if cache_key not in self._cache:
            self._cache[cache_key] = []
        
        if directory not in self._cache[cache_key]:
            self._cache[cache_key].append(directory)
            self._save_cache()
            logger.info(f"Added {dir_type} directory to cache: {directory}")
    
    def get_cache_info(self) -> Dict:
        """Get information about the cache."""
        registry_dirs = self.get_registry_directories()
        inf_dirs = self.get_inf_directories()
        
        return {
            'cache_file': str(self.cache_file),
            'cache_exists': self.cache_file.exists(),
            'registry_directories': len(registry_dirs),
            'inf_directories': len(inf_dirs),
            'valid_registry_dirs': sum(1 for d in registry_dirs if os.path.exists(d)),
            'valid_inf_dirs': sum(1 for d in inf_dirs if os.path.exists(d)),
            'is_valid': self.is_cache_valid()
        }
    
    def get_exact_path(self, target_name: str) -> Optional[str]:
        """Get exact cached path for a target file/directory.
        
        Args:
            target_name: Name of the target file/directory
            
        Returns:
            Cached path if found, None otherwise
        """
        return self._cache.get('exact_paths', {}).get(target_name)
    
    def cache_exact_path(self, target_name: str, path: str):
        """Cache an exact path for a target file/directory.
        
        Args:
            target_name: Name of the target file/directory
            path: Full path to cache
        """
        if 'exact_paths' not in self._cache:
            self._cache['exact_paths'] = {}
        
        self._cache['exact_paths'][target_name] = path
        self._save_cache()
        logger.info(f"Cached exact path for '{target_name}': {path}")

    def ensure_exists(self) -> bool:
        """Ensure the cache file exists on disk.

        Creates parent directories if necessary and writes the current in-memory
        cache to disk using the existing atomic write implementation. Returns
        True if the cache file exists after the operation, False otherwise.
        """
        try:
            # Ensure parent dir
            try:
                self.cache_file.parent.mkdir(parents=True, exist_ok=True)
            except Exception:
                logger.debug(f"Could not create cache parent dir: {self.cache_file.parent}")

            if not self.cache_file.exists():
                # Use the existing atomic save
                self._save_cache()

            return self.cache_file.exists()
        except Exception as e:
            logger.debug(f"DirectoryCache.ensure_exists error: {e}")
            return False

# Global cache instance
_cache_instance = None

def get_directory_cache() -> DirectoryCache:
    """Get the global directory cache instance."""
    global _cache_instance
    if _cache_instance is None:
        _cache_instance = DirectoryCache()
    return _cache_instance
