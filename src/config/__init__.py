"""Configuration module for Motor Report Application."""

from .app_config import AppConfig
from .directory_config import (
    OUTPUT_DIR,
    LOGS_DIR,
    ASSETS_DIR,
    NOISE_REGISTRY_FILE,
    NOISE_TEST_DIR,
    TEST_LAB_CARICHI_DIR,
    PROJECT_ROOT,
    LOGO_PATH,
)
from .runtime import is_bundled, get_bundle_root, log_runtime_info
from .measurement_units import apply_unit_preferences
from .directory_cache import get_directory_cache

__all__ = [
    "AppConfig",
    "OUTPUT_DIR",
    "LOGS_DIR",
    "ASSETS_DIR",
    "NOISE_REGISTRY_FILE",
    "NOISE_TEST_DIR",
    "TEST_LAB_CARICHI_DIR",
    "PROJECT_ROOT",
    "LOGO_PATH",
    "is_bundled",
    "get_bundle_root",
    "log_runtime_info",
    "apply_unit_preferences",
    "get_directory_cache",
]
