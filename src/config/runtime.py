"""
Centralized runtime bootstrapping for Motor Report Application.

This module handles:
- PyInstaller bundle detection
- stdin/stdout/stderr initialization (for bundled mode)
- Python path configuration
- Environment variable setup
- Asset path detection
"""

import sys
import os
import io
from pathlib import Path
from typing import Tuple, Optional
import logging

logger = logging.getLogger(__name__)


def is_bundled() -> bool:
    """
    Check if running in PyInstaller bundled mode.
    
    Returns:
        bool: True if running as PyInstaller executable, False otherwise
    """
    return hasattr(sys, '_MEIPASS')


def get_bundle_root() -> Optional[Path]:
    """
    Get the PyInstaller bundle root directory.
    
    Returns:
        Path if bundled, None otherwise
    """
    if is_bundled():
        return Path(getattr(sys, '_MEIPASS'))
    return None


def setup_stdio_for_bundle():
    """
    Fix stdin/stdout/stderr for PyInstaller bundles.
    
    PyInstaller can create executables where stdin/stdout/stderr are None,
    which causes crashes in some libraries (especially uvicorn).
    
    This function creates StringIO replacements to prevent NoneType errors.
    """
    if not is_bundled():
        return
    
    if sys.stdin is None:
        logger.debug("sys.stdin is None, creating StringIO replacement")
        sys.stdin = io.StringIO()
    
    if sys.stdout is None:
        logger.debug("sys.stdout is None, creating StringIO replacement")
        sys.stdout = io.StringIO()
    
    if sys.stderr is None:
        logger.debug("sys.stderr is None, creating StringIO replacement")
        sys.stderr = io.StringIO()


def configure_logging_for_bundle():
    """
    Configure logging to prevent crashes in bundled mode.
    
    Patches logging.config.dictConfig to handle configuration errors gracefully
    in PyInstaller environment where some logging handlers may not be available.
    """
    if not is_bundled():
        return
    
    # Suppress uvicorn logging
    os.environ['UVICORN_LOG_LEVEL'] = 'critical'
    
    # Patch logging configuration to prevent errors
    import logging.config
    original_dictConfig = logging.config.dictConfig
    
    def patched_dictConfig(config):
        try:
            return original_dictConfig(config)
        except (ValueError, AttributeError, TypeError) as e:
            logger.debug(f"Ignoring logging configuration error in bundled mode: {e}")
            # Ignore logging configuration errors in packaged app
            pass
    
    logging.config.dictConfig = patched_dictConfig


def setup_python_path() -> Path:
    """
    Configure Python path for imports.
    
    Returns:
        Path: The project root directory that was added to sys.path
    """
    if is_bundled():
        # Running in PyInstaller bundle - use the bundle directory
        bundle_root = get_bundle_root()
        if bundle_root is None:
            raise RuntimeError("Bundle root is None in bundled mode - this should never happen")
        current_dir = bundle_root
        logger.debug(f"Running in bundle, _MEIPASS: {current_dir}")
    else:
        # Running in development mode - use script directory
        current_dir = Path(__file__).parent
        logger.debug("Running in development mode")
    
    # Add to Python path if not already present
    current_dir_str = str(current_dir)
    if current_dir_str not in sys.path:
        sys.path.insert(0, current_dir_str)
        logger.debug(f"Added to Python path: {current_dir_str}")
    
    return current_dir


def setup_environment_variables() -> Tuple[Path, Path]:
    """
    Set up environment variables for runtime and assets.
    
    Returns:
        Tuple[Path, Path]: (runtime_root, assets_dir)
    """
    if is_bundled():
        bundle_root = get_bundle_root()
        if bundle_root is None:
            raise RuntimeError("Bundle root is None in bundled mode - this should never happen")
        runtime_root = bundle_root
        assets_dir = bundle_root / "assets"
    else:
        # Development mode - project root is parent of src/
        project_root = Path(__file__).resolve().parent.parent
        runtime_root = project_root
        assets_dir = project_root / "assets"
    
    # Set environment variables (don't overwrite user overrides)
    os.environ.setdefault("MOTOR_REPORT_APP_RUNTIME_ROOT", str(runtime_root))
    os.environ.setdefault("MOTOR_REPORT_APP_ASSETS", str(assets_dir))
    
    logger.debug(f"Runtime root: {runtime_root}")
    logger.debug(f"Assets directory: {assets_dir}")
    
    return runtime_root, assets_dir


def bootstrap_runtime() -> dict:
    """
    Complete runtime initialization.
    
    Performs all necessary bootstrapping steps:
    1. Detect bundle mode
    2. Fix stdio for bundles
    3. Configure logging
    4. Setup Python path
    5. Setup environment variables
    
    Returns:
        dict: Runtime configuration with keys:
            - 'bundled': bool
            - 'bundle_root': Optional[Path]
            - 'project_root': Path
            - 'assets_dir': Path
            - 'python_path_configured': bool
    """
    bundled = is_bundled()
    bundle_root = get_bundle_root()
    
    # Perform bootstrapping
    if bundled:
        logger.info("Initializing runtime for bundled executable")
        setup_stdio_for_bundle()
        configure_logging_for_bundle()
    else:
        logger.info("Initializing runtime for development mode")
    
    project_root = setup_python_path()
    runtime_root, assets_dir = setup_environment_variables()
    
    config = {
        'bundled': bundled,
        'bundle_root': bundle_root,
        'project_root': project_root,
        'runtime_root': runtime_root,
        'assets_dir': assets_dir,
        'python_path_configured': True,
    }
    
    logger.info(f"Runtime bootstrap complete: bundled={bundled}, root={runtime_root}")
    return config


def log_runtime_info():
    """
    Log detailed runtime information for debugging.
    
    Prints/logs:
    - Python executable path
    - Current working directory
    - First few sys.path entries
    - Bundle status
    - Directory contents (if in development mode)
    """
    logger.info("=== Runtime Information ===")
    logger.info(f"Python executable: {sys.executable}")
    logger.info(f"Current working directory: {os.getcwd()}")
    logger.info(f"Python path (first 3): {sys.path[:3]}")
    logger.info(f"Bundled mode: {is_bundled()}")
    
    if is_bundled():
        bundle_root = get_bundle_root()
        logger.info(f"Bundle root (_MEIPASS): {bundle_root}")
    else:
        # List directory contents in development mode
        current_dir = Path(__file__).parent
        logger.info(f"\nDirectory contents of {current_dir}:")
        try:
            for item in current_dir.iterdir():
                if item.is_dir():
                    logger.debug(f"  DIR:  {item.name}")
                else:
                    logger.debug(f"  FILE: {item.name}")
        except Exception as e:
            logger.error(f"Error listing directory: {e}")
    
    logger.info("=" * 30)


def get_user_data_dir(app_name: str = "MotorReportApp") -> Path:
    """
    Get a platform-appropriate per-user application data directory.

    Precedence:
    - Environment variable `MOTOR_REPORT_USER_DATA_DIR` if set
    - On Windows: `%LOCALAPPDATA%/<app_name>`
    - On macOS: `~/Library/Application Support/<app_name>`
    - On Linux / XDG: `$XDG_DATA_HOME/<app_name>` or `~/.local/share/<app_name>`

    Ensures the directory exists and is writable (attempts to create it).
    """
    env_override = os.environ.get('MOTOR_REPORT_USER_DATA_DIR')
    if env_override:
        p = Path(env_override)
        try:
            p.mkdir(parents=True, exist_ok=True)
        except Exception:
            logger.debug(f"Could not create env override user data dir: {p}")
        return p

    # Windows
    if os.name == 'nt':
        local_app = os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA')
        if local_app:
            p = Path(local_app) / app_name
            try:
                p.mkdir(parents=True, exist_ok=True)
            except Exception:
                logger.debug(f"Could not create user data dir under LOCALAPPDATA: {p}")
            return p

    # macOS
    if sys.platform == 'darwin':
        p = Path.home() / 'Library' / 'Application Support' / app_name
        try:
            p.mkdir(parents=True, exist_ok=True)
        except Exception:
            logger.debug(f"Could not create macOS user data dir: {p}")
        return p

    # Linux / XDG
    xdg = os.environ.get('XDG_DATA_HOME')
    if xdg:
        p = Path(xdg) / app_name
    else:
        p = Path.home() / '.local' / 'share' / app_name

    try:
        p.mkdir(parents=True, exist_ok=True)
    except Exception:
        logger.debug(f"Could not create XDG user data dir: {p}")

    return p


# Auto-bootstrap when imported (can be disabled by checking a flag)
if os.getenv('MOTOR_REPORT_NO_AUTO_BOOTSTRAP') != '1':
    _runtime_config = bootstrap_runtime()
else:
    _runtime_config = None
