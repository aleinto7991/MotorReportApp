"""
Noise Registry Loader Service - Handles loading and caching of noise registry data.

This service provides centralized noise registry management with:
- SAP code loading with caching (10-minute validity)
- Thread-safe loading with timeout protection
- Background preloading for better performance
- Stale cache handling on errors
- Performance profiling with telemetry

Extracted from src/gui/tabs/config_tab.py for better separation of concerns
and reusability across the application.
"""

import threading
import time
import logging
from typing import List, Optional
from pathlib import Path

from ..core.telemetry import log_duration

logger = logging.getLogger(__name__)


class NoiseRegistryLoader:
    """
    Manages loading and caching of noise registry data.
    
    Responsibilities:
    - Load SAP codes from noise registry with caching
    - Thread-safe loading with timeout protection (15 seconds)
    - Cache management with 10-minute validity period
    - Background preloading for improved responsiveness
    - Stale cache fallback on errors
    
    Thread Safety:
        All methods are thread-safe. Loading operations use a flag to prevent
        concurrent loads. Timeouts prevent UI blocking on slow operations.
    """
    
    # Cache validity period (10 minutes)
    CACHE_VALIDITY_SECONDS = 600
    
    # Loading timeout (15 seconds) for UI responsiveness
    LOAD_TIMEOUT_SECONDS = 15.0
    
    # Maximum rows to process for performance
    MAX_ROWS = 1000
    
    def __init__(self):
        """Initialize the noise registry loader with empty cache."""
        self._cached_sap_codes: Optional[List[str]] = None
        self._cache_timestamp: Optional[float] = None
        self._loading: bool = False
        self._load_lock = threading.Lock()
        
        logger.debug("NoiseRegistryLoader initialized")
    
    def get_sap_codes(
        self, 
        registry_path: Optional[str] = None,
        noise_test_dir: Optional[str] = None,
        force_reload: bool = False
    ) -> List[str]:
        """
        Get SAP codes from noise registry with caching.
        
        Args:
            registry_path: Path to noise registry file (REGISTRO RUMORE.xlsx)
            noise_test_dir: Path to noise test directory
            force_reload: If True, bypass cache and force reload
        
        Returns:
            List of SAP codes found in noise registry, or empty list if unavailable
        
        Thread Safety:
            Safe to call from multiple threads. Uses lock to prevent concurrent loads.
        """
        # Prevent multiple simultaneous cache loads
        if self._loading and not force_reload:
            # If cache is being loaded, return current cache (even if empty) to avoid blocking
            logger.debug("Cache is currently loading, returning cached data")
            return self._cached_sap_codes or []
        
        # Check if cache is valid
        if not force_reload and self._is_cache_valid():
            logger.debug(f"Returning cached SAP codes ({len(self._cached_sap_codes or [])} codes)")
            return self._cached_sap_codes or []
        
        # Cache is expired or empty, reload
        return self._load_sap_codes(registry_path, noise_test_dir)
    
    def _is_cache_valid(self) -> bool:
        """
        Check if the current cache is still valid.
        
        Returns:
            True if cache exists and is not expired, False otherwise
        """
        if self._cached_sap_codes is None or self._cache_timestamp is None:
            return False
        
        current_time = time.time()
        age = current_time - self._cache_timestamp
        
        return age < self.CACHE_VALIDITY_SECONDS
    
    def _load_sap_codes(
        self, 
        registry_path: Optional[str],
        noise_test_dir: Optional[str]
    ) -> List[str]:
        """
        Load SAP codes from noise registry with timeout protection.
        
        Args:
            registry_path: Path to noise registry file
            noise_test_dir: Path to noise test directory
        
        Returns:
            List of SAP codes, or empty list on error/timeout
        
        Thread Safety:
            Uses lock to ensure only one load operation at a time.
        """
        with self._load_lock:
            # Double-check: another thread might have loaded while we waited
            if self._is_cache_valid():
                return self._cached_sap_codes or []
            
            self._loading = True
            try:
                # Determine cache status for logging
                if self._cached_sap_codes is None:
                    logger.info("Loading noise registry for the first time...")
                else:
                    logger.info("Refreshing noise registry cache...")
                
                # Validate inputs
                if not registry_path:
                    logger.warning("No noise registry path provided")
                    self._cached_sap_codes = []
                    self._cache_timestamp = time.time()
                    return []
                
                if not noise_test_dir:
                    logger.warning("No noise test directory provided")
                    self._cached_sap_codes = []
                    self._cache_timestamp = time.time()
                    return []
                
                # Import validator
                from ..validators.noise_test_validator import NoiseTestValidator
                
                # Setup validator with correct sheet name
                sheet_name = "Registro"  # Correct sheet name for REGISTRO RUMORE.xlsx
                validator = NoiseTestValidator(str(noise_test_dir), sheet_name)
                
                logger.debug(f"Using noise registry: {registry_path}")
                logger.debug(f"Using sheet name: {sheet_name}")
                logger.debug(f"Using noise test directory: {noise_test_dir}")
                
                # Load with timeout protection and profiling
                with log_duration(logger, f"Noise registry load from {registry_path}"):
                    result = self._load_with_timeout(validator, registry_path)
                
                # Update cache
                self._cached_sap_codes = result
                self._cache_timestamp = time.time()
                
                logger.info(f"Loaded {len(result)} noise SAP codes successfully")
                return result
                
            except Exception as e:
                logger.error(f"Error loading noise SAP codes: {e}", exc_info=True)
                
                # Return stale cache if available
                if self._cached_sap_codes is not None:
                    logger.warning("Using stale cache due to error")
                    return self._cached_sap_codes
                
                # Cache empty result to avoid repeated failures
                self._cached_sap_codes = []
                self._cache_timestamp = time.time()
                return []
                
            finally:
                self._loading = False
    
    def _load_with_timeout(
        self, 
        validator,
        registry_path: str
    ) -> List[str]:
        """
        Load SAP codes with timeout protection for UI responsiveness.
        
        Args:
            validator: NoiseTestValidator instance
            registry_path: Path to noise registry file
        
        Returns:
            List of SAP codes, or empty list on timeout/error
        """
        start_time = time.time()
        
        # Container for thread results
        result_container: List[Optional[List[str]]] = [None]
        error_container: List[Optional[Exception]] = [None]
        
        def load_task():
            """Background loading task."""
            try:
                # Use lightweight SAP code extraction for fast UI pre-check
                result_container[0] = validator.get_sap_codes_from_registry(
                    registry_path, 
                    max_rows=self.MAX_ROWS
                )
            except Exception as e:
                error_container[0] = e
        
        # Start loading in background thread with timeout
        load_thread = threading.Thread(target=load_task, daemon=True)
        load_thread.start()
        
        # Wait for completion with timeout
        load_thread.join(timeout=self.LOAD_TIMEOUT_SECONDS)
        
        load_time = time.time() - start_time
        
        if load_thread.is_alive():
            # Thread is still running - timeout occurred
            logger.warning(f"Noise registry loading timed out after {self.LOAD_TIMEOUT_SECONDS}s")
            return []
        
        if error_container[0]:
            # Error occurred during loading
            logger.error(f"Error during noise registry loading: {error_container[0]}")
            return []
        
        if result_container[0] is not None:
            # Successfully loaded
            logger.debug(f"Noise registry loaded in {load_time:.2f}s")
            return result_container[0]
        
        # Unexpected state
        logger.warning("Unexpected state during noise registry loading")
        return []
    
    def preload_async(
        self,
        registry_path: Optional[str] = None,
        noise_test_dir: Optional[str] = None
    ):
        """
        Preload SAP codes in background for improved responsiveness.
        
        Starts a daemon thread to load noise registry SAP codes without
        blocking the UI. Useful for loading data while the user is still
        interacting with previous steps.
        
        Args:
            registry_path: Path to noise registry file (REGISTRO RUMORE.xlsx)
            noise_test_dir: Path to noise test directory for validation
        
        Side Effects:
            - Spawns background daemon thread if not already loading
            - Updates cache when loading completes
            - Logs completion time and any errors
        
        Thread Safety:
            Safe to call multiple times. Subsequent calls are ignored if
            a load is already in progress.
        
        Example:
            >>> loader = NoiseRegistryLoader()
            >>> loader.preload_async("/path/to/registry.xlsx", "/path/to/tests")
            >>> # User continues working...
            >>> saps = loader.get_sap_codes()  # Returns quickly from cache
        """
        def preload_task():
            """Background preload task."""
            try:
                start_time = time.time()
                self.get_sap_codes(registry_path, noise_test_dir)
                load_time = time.time() - start_time
                logger.info(f"Background noise SAP codes preload completed in {load_time:.2f}s")
            except Exception as e:
                logger.error(f"Error preloading noise SAP codes: {e}", exc_info=True)
        
        # Only start background loading if not already in progress
        if not self._loading:
            threading.Thread(target=preload_task, daemon=True).start()
            logger.debug("Started background noise registry preload")
    
    def clear_cache(self):
        """
        Clear the cache to force reload on next request.
        
        Useful when noise registry file has been updated and you want to
        ensure fresh data is loaded on the next get_sap_codes() call.
        
        Side Effects:
            - Sets _cached_sap_codes to None
            - Sets _cache_timestamp to None
            - Forces cache validation to fail
            - Next get_sap_codes() will reload from file
        
        Thread Safety:
            Safe to call anytime. Uses lock to prevent race conditions.
            Does not interrupt ongoing loads - they will complete but
            their results won't be cached.
        
        Example:
            >>> loader.clear_cache()
            >>> saps = loader.get_sap_codes(registry_path)  # Reloads from file
        """
        with self._load_lock:
            self._cached_sap_codes = None
            self._cache_timestamp = None
            logger.debug("Noise registry cache cleared")
    
    def get_cache_age(self) -> Optional[float]:
        """
        Get the age of the current cache in seconds.
        
        Useful for displaying cache freshness to users or for custom
        cache invalidation logic.
        
        Returns:
            float: Age in seconds if cache exists
            None: If cache doesn't exist (never loaded or cleared)
        
        Thread Safety:
            Safe to call anytime. Read-only operation.
        
        Example:
            >>> age = loader.get_cache_age()
            >>> if age and age > 300:
            >>>     print(f"Cache is {age:.0f}s old, consider refreshing")
        """
        if self._cache_timestamp is None:
            return None
        
        return time.time() - self._cache_timestamp
    
    def is_loading(self) -> bool:
        """
        Check if a load operation is currently in progress.
        
        Useful for UI feedback (showing loading indicators) or for
        deciding whether to wait for ongoing loads vs. using stale cache.
        
        Returns:
            True: A load operation is in progress
            False: No load operation running
        
        Thread Safety:
            Safe to call anytime. Read-only operation on atomic flag.
        
        Note:
            This is a best-effort indicator. Due to threading, the state
            may change immediately after this call returns.
        
        Example:
            >>> if loader.is_loading():
            >>>     show_spinner()
            >>> else:
            >>>     hide_spinner()
        """
        return self._loading
