"""
Directory cache system to avoid repeated slow directory searches.
Stores found directories in a JSON file for quick retrieval.
"""

import json
import os
from pathlib import Path
from typing import Dict, List, Optional
import logging

logger = logging.getLogger(__name__)

class DirectoryCache:
    """Manages cached directory paths to avoid repeated searches."""
    
    def __init__(self, cache_file: str = "directory_cache.json"):
        """Initialize directory cache.
        
        Args:
            cache_file: Name of the cache file (stored in project root)
        """
        # Store cache file in project root
        self.cache_file = Path(__file__).parent.parent / cache_file
        self._cache = self._load_cache()
    
    def _load_cache(self) -> Dict:
        """Load cache from file."""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    cache = json.load(f)
                    logger.info(f"Loaded directory cache with {len(cache)} entries")
                    return cache
        except Exception as e:
            logger.warning(f"Could not load directory cache: {e}")
        
        return {}
    
    def _save_cache(self):
        """Save cache to file."""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self._cache, f, indent=2, ensure_ascii=False)
                logger.info(f"Saved directory cache with {len(self._cache)} entries")
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

# Global cache instance
_cache_instance = None

def get_directory_cache() -> DirectoryCache:
    """Get the global directory cache instance."""
    global _cache_instance
    if _cache_instance is None:
        _cache_instance = DirectoryCache()
    return _cache_instance
