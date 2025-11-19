import os
import sys
from pathlib import Path
import logging
import signal
import time
from .directory_cache import get_directory_cache
from ..utils.common import validate_directory_path

# Configure logging
logger = logging.getLogger(__name__)

def _search_directory_limited(dir_path: str, target_lower: str, max_depth: int = 3, timeout_seconds: int = 2) -> Path | None:
    """
    Search for a target file/directory with limited depth and timeout protection.
    
    Args:
        dir_path: Directory to search in
        target_lower: Target name in lowercase
        max_depth: Maximum search depth (default: 3)
        timeout_seconds: Maximum time to spend searching (default: 2)
    
    Returns:
        Path to found item or None
    """
    start_time = time.time()
    
    def timeout_handler(signum, frame):
        raise TimeoutError("Directory search timeout")
    
    try:
        # Set up timeout (Windows compatible)
        if os.name != 'nt':  # Unix-like systems
            signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(timeout_seconds)
        
        # Limited depth search
        for root, dirs, files in os.walk(dir_path):
            # Check timeout manually for Windows compatibility
            if time.time() - start_time > timeout_seconds:
                logger.warning(f"Search timeout in {dir_path}")
                break
                
            # Calculate current depth
            current_depth = len(Path(root).relative_to(Path(dir_path)).parts)
            if current_depth >= max_depth:
                dirs.clear()  # Don't go deeper
                continue
            
            # Check files and directories
            for item in dirs + files:
                if item.lower() == target_lower and not item.startswith('~'):
                    return Path(root) / item
        
        return None
        
    except (TimeoutError, OSError) as e:
        logger.warning(f"Search interrupted in {dir_path}: {e}")
        return None
    finally:
        if os.name != 'nt':  # Unix-like systems
            signal.alarm(0)  # Cancel alarm

def _bundled_project_root() -> Path | None:
    """Resolve the project root when running from a bundled executable."""
    # Prefer explicit overrides set by the bootstrap script
    env_root = os.environ.get("MOTOR_REPORT_APP_RUNTIME_ROOT")
    if env_root:
        candidate = Path(env_root)
        assets_dir = os.environ.get("MOTOR_REPORT_APP_ASSETS")
        assets_path = Path(assets_dir) if assets_dir else candidate / "assets"
        if assets_path.exists():
            return candidate

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundle_root = Path(getattr(sys, "_MEIPASS"))  # type: ignore[attr-defined]
        if (bundle_root / "assets").exists():
            return bundle_root
    return None


def find_project_root() -> Path | None:
    """
    Searches for the project root directory by traversing up from the current file's location.
    The project root is identified by containing both 'src' and 'assets' subdirectories.
    """
    bundled_root = _bundled_project_root()
    if bundled_root:
        logger.info(f"Project root resolved from bundled runtime: {bundled_root}")
        return bundled_root
    try:
        # Start from the directory of the current file (__file__) and resolve to an absolute path
        current_dir = Path(__file__).parent.resolve()
    except NameError:
        # Fallback for environments where __file__ is not defined
        current_dir = Path.cwd()

    # Validate the starting directory
    if not validate_directory_path(current_dir):
        logger.error(f"Invalid starting directory: {current_dir}")
        return None

    # Traverse up the directory tree from the current directory
    for parent in [current_dir] + list(current_dir.parents):
        # Validate each parent directory
        validated_parent = validate_directory_path(parent)
        if not validated_parent:
            continue
            
        if (validated_parent / 'src').is_dir():
            logger.info(f"Project root found at: {validated_parent}")
            return validated_parent
            
    logger.error(f"Could not find a valid project root. Traversed up from {current_dir}.")
    return None

def find_onedrive_root() -> Path | None:
    """
    Finds the user's OneDrive directory. It prioritizes OneDrive for Business/School
    (e.g., 'OneDrive - YourOrganization') over personal OneDrive folders if both exist.
    Enhanced for desktop app compatibility.
    """
    import os
    
    # First check environment variables (works better in desktop apps)
    onedrive_env_vars = [
        os.environ.get('OneDrive'),
        os.environ.get('OneDriveCommercial'), 
        os.environ.get('OneDriveConsumer'),
    ]
    
    for env_path in onedrive_env_vars:
        if env_path:
            path = Path(env_path)
            validated = validate_directory_path(path)
            if validated:
                logger.info(f"Found OneDrive via environment variable: {validated}")
                return validated
    
    # Fallback to home directory search
    home_dir = Path.home()
    
    # Validate home directory
    validated_home = validate_directory_path(home_dir)
    if not validated_home:
        logger.error(f"Invalid home directory: {home_dir}")
        return None
    
    try:
        onedrive_folders = [d for d in validated_home.iterdir() 
                          if d.is_dir() and d.name.lower().startswith('onedrive') 
                          and validate_directory_path(d)]
    except (PermissionError, OSError) as e:
        logger.warning(f"Cannot access home directory {validated_home}: {e}")
        return None

    if not onedrive_folders:
        logger.warning(f"Could not find any OneDrive directory in {validated_home}.")
        # Try hardcoded common paths as last resort
        username = os.environ.get('USERNAME', os.environ.get('USER', ''))
        if username:
            common_paths = [
                Path(f"C:/Users/{username}/OneDrive - AMETEK Inc"),
                Path(f"C:/Users/{username}/OneDrive"),
                Path(f"C:/Users/{username}/OneDrive for Business"),
            ]
            for path in common_paths:
                if validate_directory_path(path):
                    logger.info(f"Found OneDrive at hardcoded path: {path}")
                    return path
        return None

    # Prioritize corporate/university OneDrive folders
    for folder in onedrive_folders:
        if " - " in folder.name:
            logger.info(f"Prioritizing OneDrive for Business/School: {folder}")
            return folder

    # If no corporate folder is found, return the first one found (likely personal)
    logger.info(f"Found standard OneDrive directory at: {onedrive_folders[0]}")
    return onedrive_folders[0]

def find_all_paths(base_path: Path, targets: dict[str, str]) -> dict[str, Path | None]:
    """
    Performs a single, efficient walk through the directory tree to find multiple target files/directories.
    This is much faster than searching for each file individually.
    Uses directory cache to avoid repeated searches.

    Args:
        base_path: The root directory to start the search from (e.g., OneDrive root).
        targets: A dictionary where keys are logical names (e.g., "PERFORMANCE_DIR")
                 and values are the actual file/directory names to find (e.g., "ProveEffettuate").

    Returns:
        A dictionary with the same keys as the input `targets` dict, where each value
        is the found Path object or None if not found.
    """
    cache = get_directory_cache()
    results: dict[str, Path | None] = {key: None for key in targets}
    
    # Check cache first
    if cache.is_cache_valid():
        logger.info("Using cached directories (fast path)")
        
        # Get cached directories
        registry_dirs = cache.get_registry_directories()
        inf_dirs = cache.get_inf_directories()
        
        # Look for target files in cached directories (OPTIMIZED for performance)
        for key, target_name in targets.items():
            target_lower = target_name.lower()
            
            # OPTIMIZATION: Check if we already have this exact target cached with full path
            cached_result = cache.get_exact_path(target_name)
            if cached_result and os.path.exists(cached_result):
                results[key] = Path(cached_result)
                logger.info(f"Found '{target_name}' in exact path cache: {cached_result}")
                continue
            
            # OPTIMIZATION: Use limited-depth search to avoid performance issues
            found = False
            
            # Check in registry directories (LIMITED DEPTH search for performance)
            for dir_path in registry_dirs:
                if not os.path.exists(dir_path):
                    continue
                    
                # PERFORMANCE: Limited depth search with timeout protection
                found_path = _search_directory_limited(dir_path, target_lower, max_depth=3, timeout_seconds=2)
                if found_path:
                    results[key] = found_path
                    logger.info(f"Found '{target_name}' in cached registry search: {found_path}")
                    # OPTIMIZATION: Cache the exact path for future use
                    cache.cache_exact_path(target_name, str(found_path))
                    found = True
                    break
            
            # Check in INF directories if not found yet (LIMITED DEPTH search for performance)
            if not found:
                for dir_path in inf_dirs:
                    if not os.path.exists(dir_path):
                        continue
                        
                    # PERFORMANCE: Limited depth search with timeout protection
                    found_path = _search_directory_limited(dir_path, target_lower, max_depth=3, timeout_seconds=2)
                    if found_path:
                        results[key] = found_path
                        logger.info(f"Found '{target_name}' in cached INF search: {found_path}")
                        # OPTIMIZATION: Cache the exact path for future use
                        cache.cache_exact_path(target_name, str(found_path))
                        found = True
                        break
                    
                if found:
                    break
        
        # If we found everything in cache, return early
        if all(results.values()):
            logger.info("All targets found in cache!")
            return results
        
        # If we found some but not all, fall through to full search
        missing = [k for k, v in results.items() if v is None]
        logger.info(f"Cache provided partial results. Still need to search for: {missing}")
    
    # Perform full search (either cache was invalid or incomplete)
    logger.info(f"Starting full directory search for {len(targets)} targets in '{base_path}'...")
    
    # Create a copy of the values to search for, converting to lowercase for case-insensitive matching
    remaining_targets_map = {v.lower(): k for k, v in targets.items() if results[k] is None}
    
    # Track found directories for caching
    registry_dirs_found = set()
    inf_dirs_found = set()

    for root, dirs, files in os.walk(base_path):
        if not remaining_targets_map:
            logger.info("All targets found. Halting search early.")
            break

        # Check both directories and files
        for name in dirs + files:
            name_lower = name.lower()
            if name_lower in remaining_targets_map:
                # Skip temporary office files (e.g., ~$Registro LAB.xlsx)
                if name.startswith('~'):
                    continue

                key = remaining_targets_map[name_lower]
                found_path = Path(root) / name
                results[key] = found_path
                logger.info(f"Found '{name}' for key '{key}' at: {found_path}")
                
                # Track directory for caching
                parent_dir = str(Path(root))
                if 'registro' in name_lower or 'lab' in name_lower:
                    registry_dirs_found.add(parent_dir)
                elif 'prove' in name_lower or 'test' in name_lower or '.inf' in name_lower:
                    inf_dirs_found.add(parent_dir)
                
                # Remove from remaining targets so we don't find it again
                del remaining_targets_map[name_lower]

                # If we found everything, we can break out of the inner loop
                if not remaining_targets_map:
                    break
    
    # Update cache with newly found directories
    if registry_dirs_found:
        cache.set_registry_directories(list(registry_dirs_found))
    if inf_dirs_found:
        cache.set_inf_directories(list(inf_dirs_found))
    
    if remaining_targets_map:
        for target_name_lower, key in remaining_targets_map.items():
             logger.warning(f"Could not find a path for target '{targets[key]}' (key: {key}) within {base_path}")

    logger.info("Directory search complete.")
    return results

# --- Manual Path Management Functions ---

def update_manual_paths(performance_dir: str | None = None, noise_dir: str | None = None, 
                       lab_registry: str | None = None, noise_registry: str | None = None) -> dict:
    """
    Update paths manually and cache them for future use.
    
    Args:
        performance_dir: Manual path to ProveEffettuate directory
        noise_dir: Manual path to Tests Rumore directory
        lab_registry: Manual path to Registro LAB.xlsx file
        noise_registry: Manual path to REGISTRO RUMORE.xlsx file
    
    Returns:
        Dictionary with validation results for each path
    """
    global PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE, TEST_LAB_CARICHI_DIR
    
    results = {}
    cache = get_directory_cache()
    
    # Update performance directory
    if performance_dir:
        path = Path(performance_dir)
        if validate_directory_path(path):
            PERFORMANCE_TEST_DIR = path
            cache.cache_exact_path("ProveEffettuate", str(path))
            results['performance_dir'] = {'status': 'success', 'path': str(path)}
            logger.info(f"Manually updated Performance Test Dir: {path}")
        else:
            results['performance_dir'] = {'status': 'error', 'message': 'Invalid directory path'}
    
    # Update noise directory
    if noise_dir:
        path = Path(noise_dir)
        if validate_directory_path(path):
            NOISE_TEST_DIR = path
            cache.cache_exact_path("Tests Rumore", str(path))
            results['noise_dir'] = {'status': 'success', 'path': str(path)}
            logger.info(f"Manually updated Noise Test Dir: {path}")
        else:
            results['noise_dir'] = {'status': 'error', 'message': 'Invalid directory path'}
    
    # Update lab registry file
    if lab_registry:
        path = Path(lab_registry)
        if path.exists() and path.is_file() and path.suffix.lower() == '.xlsx':
            LAB_REGISTRY_FILE = path
            cache.cache_exact_path("Registro LAB.xlsx", str(path))
            results['lab_registry'] = {'status': 'success', 'path': str(path)}
            logger.info(f"Manually updated Lab Registry File: {path}")
        else:
            results['lab_registry'] = {'status': 'error', 'message': 'Invalid Excel file path'}
    
    # Update noise registry file
    if noise_registry:
        path = Path(noise_registry)
        if path.exists() and path.is_file() and path.suffix.lower() == '.xlsx':
            NOISE_REGISTRY_FILE = path
            cache.cache_exact_path("REGISTRO RUMORE.xlsx", str(path))
            results['noise_registry'] = {'status': 'success', 'path': str(path)}
            logger.info(f"Manually updated Noise Registry File: {path}")
        else:
            results['noise_registry'] = {'status': 'error', 'message': 'Invalid Excel file path'}
    
    return results

def refresh_directory_cache() -> dict:
    """
    Force refresh the directory cache by re-scanning for all target directories.
    
    Returns:
        Dictionary with refresh results
    """
    global PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE
    
    logger.info("Force refreshing directory cache...")
    
    # Invalidate current cache
    invalidate_directory_cache()
    
    # Re-run the search
    if ONEDRIVE_ROOT:
        targets_to_find = {
            "PERFORMANCE_TEST_DIR": "ProveEffettuate",
            "NOISE_TEST_DIR": "Tests Rumore", 
            "LAB_REGISTRY_FILE": "Registro LAB.xlsx",
            "NOISE_REGISTRY_FILE": "REGISTRO RUMORE.xlsx",
            "TEST_LAB_CARICHI_DIR": "CARICHI NOMINALI",
        }
        
        logger.info("Re-scanning OneDrive for all target directories...")
        found_paths = find_all_paths(ONEDRIVE_ROOT, targets_to_find)
        
        # Update global variables
        PERFORMANCE_TEST_DIR = found_paths.get("PERFORMANCE_TEST_DIR")
        NOISE_TEST_DIR = found_paths.get("NOISE_TEST_DIR")
        LAB_REGISTRY_FILE = found_paths.get("LAB_REGISTRY_FILE")
        NOISE_REGISTRY_FILE = found_paths.get("NOISE_REGISTRY_FILE")
        TEST_LAB_CARICHI_DIR = found_paths.get("TEST_LAB_CARICHI_DIR")
        
        # Return results
        results = {
            'status': 'success',
            'found_paths': {
                'performance_dir': str(PERFORMANCE_TEST_DIR) if PERFORMANCE_TEST_DIR else None,
                'noise_dir': str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None,
                'lab_registry': str(LAB_REGISTRY_FILE) if LAB_REGISTRY_FILE else None,
                'noise_registry': str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
                'test_lab_dir': str(TEST_LAB_CARICHI_DIR) if TEST_LAB_CARICHI_DIR else None,
            },
            'success_count': sum(1 for p in [PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE, TEST_LAB_CARICHI_DIR] if p)
        }
        
        logger.info(f"Cache refresh complete. Found {results['success_count']}/{len(targets_to_find)} targets")
        return results
    else:
        return {'status': 'error', 'message': 'OneDrive root not available'}

def get_current_paths() -> dict:
    """
    Get current configured paths.
    
    Returns:
        Dictionary with all current paths
    """
    return {
        'project_root': str(PROJECT_ROOT) if PROJECT_ROOT else None,
        'onedrive_root': str(ONEDRIVE_ROOT) if ONEDRIVE_ROOT else None,
        'performance_dir': str(PERFORMANCE_TEST_DIR) if PERFORMANCE_TEST_DIR else None,
        'noise_dir': str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None,
        'lab_registry': str(LAB_REGISTRY_FILE) if LAB_REGISTRY_FILE else None,
        'noise_registry': str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
        'lf_registry': str(LF_REGISTRY_FILE) if LF_REGISTRY_FILE else None,
        'lf_base_dir': str(LF_BASE_DIR) if LF_BASE_DIR else None,
        'test_lab_dir': str(TEST_LAB_CARICHI_DIR) if TEST_LAB_CARICHI_DIR else None,
        'assets_dir': str(ASSETS_DIR) if ASSETS_DIR else None,
        'logs_dir': str(LOGS_DIR) if LOGS_DIR else None,
        'output_dir': str(OUTPUT_DIR) if OUTPUT_DIR else None
    }

# --- Cache Management Functions ---

def invalidate_directory_cache():
    """Invalidate the directory cache. Call this when directory-related operations fail."""
    cache = get_directory_cache()
    cache.invalidate_cache()
    logger.info("Directory cache invalidated. Next startup will perform fresh directory search.")

def get_cache_status():
    """Get information about the directory cache status."""
    cache = get_directory_cache()
    return cache.get_cache_info()

# --- Configuration Setup ---

PROJECT_ROOT = find_project_root()
ONEDRIVE_ROOT = find_onedrive_root()

# --- Project-Relative Paths ---
ASSETS_DIR = None
LOGS_DIR = None
SRC_DIR = None
OUTPUT_DIR = Path.home() / 'Desktop' # Default output, always available
LOGO_PATH = None
IS_BUNDLED = getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")

if PROJECT_ROOT:
    assets_override = os.environ.get("MOTOR_REPORT_APP_ASSETS")
    ASSETS_DIR = Path(assets_override) if assets_override else PROJECT_ROOT / 'assets'
    SRC_DIR = PROJECT_ROOT / 'src'
    LOGO_PATH = ASSETS_DIR / 'logo.png'

    if IS_BUNDLED:
        logs_base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
        LOGS_DIR = logs_base / "MotorReportApp" / "logs"
    else:
        LOGS_DIR = PROJECT_ROOT / 'logs'

    # Ensure directories we own exist
    if LOGS_DIR:
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
    if ASSETS_DIR and not IS_BUNDLED:
        ASSETS_DIR.mkdir(parents=True, exist_ok=True)
else:
    logger.critical("Fatal: Project root could not be determined. Application might not find assets or source files.")

# Always ensure output dir exists
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# --- OneDrive-Relative Data Paths ---
# These will be initialized at runtime, not at import time
PERFORMANCE_TEST_DIR = None
NOISE_TEST_DIR = None
LAB_REGISTRY_FILE = None
NOISE_REGISTRY_FILE = None
LF_REGISTRY_FILE = None
LF_BASE_DIR = None
TEST_LAB_CARICHI_DIR = None

# Flag to track if directories have been initialized
_directories_initialized = False

def _initialize_data_directories():
    """
    Initialize data directories at runtime (not at import time).
    This ensures that the executable works for any user, not just the build machine.
    """
    global PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE
    global LF_REGISTRY_FILE, LF_BASE_DIR, TEST_LAB_CARICHI_DIR, _directories_initialized
    
    # Only initialize once
    if _directories_initialized:
        return
    
    _directories_initialized = True
    
    # Get current user's OneDrive root
    onedrive_root = find_onedrive_root()
    if not onedrive_root:
        logger.warning("OneDrive root not found - data directories will not be initialized")
        return
    # Define the logical names and the actual file/folder names to find.
    # This makes the code cleaner and the results dictionary easy to use.
    targets_to_find = {
        "PERFORMANCE_TEST_DIR": "ProveEffettuate",
        "PERFORMANCE_TEST_DIR_ALT": "Prove Effettuate SalaProve",  # Alternative name
        "NOISE_TEST_DIR": "Tests Rumore", 
        "LAB_REGISTRY_FILE": "Registro LAB.xlsx",
        "NOISE_REGISTRY_FILE": "REGISTRO RUMORE.xlsx",
        "LF_REGISTRY_FILE": "REGISTRO LF .xlsx",
        "LF_BASE_DIR": "RELIABIL",  # Directory containing year folders with LF tests
        "TEST_LAB_CARICHI_DIR": "CARICHI NOMINALI",
    }

    # Initialize results dictionary
    found_paths: dict[str, Path | None] = {key: None for key in targets_to_find}
    
    # Search for UTE_wrk directory first, then search within it
    logger.info("Searching for UTE_wrk directory under OneDrive...")
    ute_wrk_path = None
    
    # Improved UTE_wrk search with early termination and better error handling
    try:
        # First check common locations (faster)
        common_ute_wrk_paths = [
            onedrive_root / "ENG & Quality" / "UTE_wrk",
            onedrive_root / "Engineering" / "UTE_wrk", 
            onedrive_root / "UTE_wrk",
        ]
        
        for potential_path in common_ute_wrk_paths:
            if potential_path.exists() and potential_path.is_dir():
                ute_wrk_path = potential_path
                logger.info(f"Found UTE_wrk at common location: {ute_wrk_path}")
                break
        
        # If not found in common locations, do a broader search with depth limit
        if not ute_wrk_path:
            logger.info("UTE_wrk not found in common locations, performing broader search...")
            max_search_depth = 3  # Limit search depth for performance
            
            for root, dirs, files in os.walk(onedrive_root):
                # Calculate current depth to avoid going too deep
                try:
                    current_depth = len(Path(root).relative_to(onedrive_root).parts)
                    if current_depth > max_search_depth:
                        dirs.clear()  # Don't go deeper
                        continue
                except ValueError:
                    # Skip if we can't calculate relative path
                    continue
                
                # Look for UTE_wrk directory (case-insensitive)
                for dir_name in dirs[:]:  # Create a copy to safely modify during iteration
                    if dir_name.lower() == "ute_wrk":
                        ute_wrk_path = Path(root) / dir_name
                        logger.info(f"Found UTE_wrk directory at: {ute_wrk_path}")
                        # Clear dirs to stop deeper searching in this branch
                        dirs.clear()
                        break
                
                if ute_wrk_path:
                    break
    
    except (PermissionError, OSError) as e:
        logger.warning(f"Error during UTE_wrk search: {e}")
    
    if not ute_wrk_path:
        logger.warning("UTE_wrk directory not found under OneDrive")
    
    if ute_wrk_path and validate_directory_path(ute_wrk_path):
        logger.info(f"Using UTE_wrk path: {ute_wrk_path}")
        
        # Search for all targets within the UTE_wrk directory structure
        logger.info("Searching for targets within UTE_wrk directory...")
        found_paths = find_all_paths(ute_wrk_path, targets_to_find)
        
        # Validate and log results
        success_count = 0
        for key, path in found_paths.items():
            if path and path.exists():
                logger.info(f"✅ Found {key}: {path}")
                success_count += 1
            else:
                logger.warning(f"❌ Could not find {key} in UTE_wrk directory")
        
        logger.info(f"Found {success_count}/{len(targets_to_find)} targets in UTE_wrk")
        
        # If we didn't find everything, try a secondary search in common subdirectories
        if success_count < len(targets_to_find):
            logger.info("Performing secondary search in common subdirectories...")
            
            common_subdirs = ["LAB", "RUMORE", "REPORTS", "PROJECTS", "UTE"]
            for subdir_name in common_subdirs:
                subdir_path = ute_wrk_path / subdir_name
                if subdir_path.exists() and subdir_path.is_dir():
                    logger.info(f"Searching in subdirectory: {subdir_path}")
                    
                    # Search for missing targets in this subdirectory
                    missing_targets = {}
                    for k, v in targets_to_find.items():
                        found_path = found_paths.get(k)
                        if not found_path or not found_path.exists():
                            missing_targets[k] = v
                    
                    if missing_targets:
                        subdir_results = find_all_paths(subdir_path, missing_targets)
                        
                        # Update found_paths with any newly found items
                        for key, path in subdir_results.items():
                            if path and path.exists() and not found_paths.get(key):
                                found_paths[key] = path
                                logger.info(f"✅ Found {key} in {subdir_name}: {path}")
                                success_count += 1
            
            logger.info(f"After secondary search: {success_count}/{len(targets_to_find)} targets found")
    
    else:
        logger.warning("UTE_wrk directory not found or not accessible, falling back to full OneDrive search")
        # Fallback to searching the entire OneDrive (with progress indication)
        logger.info("Starting full OneDrive search (this may take longer)...")
        found_paths = find_all_paths(onedrive_root, targets_to_find)

    # Assign the results to our configuration variables
    PERFORMANCE_TEST_DIR = found_paths.get("PERFORMANCE_TEST_DIR")
    
    # Check if we found the alternative path and prefer it
    alt_performance_dir = found_paths.get("PERFORMANCE_TEST_DIR_ALT")
    if alt_performance_dir:
        # Look for ProveEffettuate inside the alternative directory
        prove_effettuate_path = alt_performance_dir / "ProveEffettuate"
        if prove_effettuate_path.exists():
            PERFORMANCE_TEST_DIR = prove_effettuate_path
            logger.info(f"Using preferred ProveEffettuate path: {PERFORMANCE_TEST_DIR}")
        else:
            PERFORMANCE_TEST_DIR = alt_performance_dir
            logger.info(f"Using alternative performance directory: {PERFORMANCE_TEST_DIR}")
    elif not PERFORMANCE_TEST_DIR:
        logger.warning("Could not find ProveEffettuate directory in any location")
    
    NOISE_TEST_DIR = found_paths.get("NOISE_TEST_DIR")
    LAB_REGISTRY_FILE = found_paths.get("LAB_REGISTRY_FILE")
    NOISE_REGISTRY_FILE = found_paths.get("NOISE_REGISTRY_FILE")
    LF_REGISTRY_FILE = found_paths.get("LF_REGISTRY_FILE")
    LF_BASE_DIR = found_paths.get("LF_BASE_DIR")
    TEST_LAB_CARICHI_DIR = found_paths.get("TEST_LAB_CARICHI_DIR")
    
    # --- Final Log of All Configured Paths ---
    logger.info("--- Final Directory Configuration ---")
    logger.info(f"  Project Root: {PROJECT_ROOT}")
    logger.info(f"  OneDrive Root: {onedrive_root}")
    logger.info(f"  Assets Dir:   {ASSETS_DIR}")
    logger.info(f"  Logs Dir:     {LOGS_DIR}")
    logger.info(f"  Output Dir:   {OUTPUT_DIR}")
    logger.info(f"  Logo Path:    {LOGO_PATH}")
    logger.info(f"  Performance Data Dir: {PERFORMANCE_TEST_DIR}")
    logger.info(f"  Noise Data Dir:       {NOISE_TEST_DIR}")
    logger.info(f"  Lab Registry File:    {LAB_REGISTRY_FILE}")
    logger.info(f"  Noise Registry File:  {NOISE_REGISTRY_FILE}")
    logger.info(f"  LF Registry File:     {LF_REGISTRY_FILE}")
    logger.info(f"  LF Base Dir:          {LF_BASE_DIR}")
    logger.info(f"  Test Lab Dir:         {TEST_LAB_CARICHI_DIR}")
    logger.info("-----------------------------------")

    if not all([PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE]):
        logger.warning("One or more data/registry paths could not be determined. The application may not function correctly.")


# DO NOT call _initialize_data_directories() here!
# It must be called by the application at runtime, not during module import.
# This prevents PyInstaller from baking in build-time paths.
# The application should call ensure_directories_initialized() before using any paths.


def ensure_directories_initialized():
    """
    Ensure data directories are initialized. Safe to call multiple times.
    This should be called by the application before accessing any data paths.
    """
    if not _directories_initialized:
        _initialize_data_directories()


def update_cached_paths(paths: dict) -> dict:
    """
    Update cached paths with manually provided paths.
    
    Args:
        paths: Dictionary with path updates
               Keys can be: 'performance_dir', 'noise_dir', 'lab_registry', 'noise_registry', 
                           'lf_registry', 'lf_base_dir', 'test_lab_dir'
    
    Returns:
        Dictionary with update results
    """
    global PERFORMANCE_TEST_DIR, NOISE_TEST_DIR, LAB_REGISTRY_FILE, NOISE_REGISTRY_FILE, LF_REGISTRY_FILE, LF_BASE_DIR, TEST_LAB_CARICHI_DIR
    
    logger.info("Updating cached paths with manual entries...")
    
    results = {'updated': [], 'errors': []}
    
    # Update Performance Test Directory
    if 'performance_dir' in paths and paths['performance_dir']:
        try:
            path = Path(paths['performance_dir'])
            if path.exists():
                PERFORMANCE_TEST_DIR = path
                results['updated'].append(f"Performance directory: {path}")
                logger.info(f"Updated PERFORMANCE_TEST_DIR to: {path}")
            else:
                results['errors'].append(f"Performance directory not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating performance directory: {e}")
    
    # Update Noise Test Directory
    if 'noise_dir' in paths and paths['noise_dir']:
        try:
            path = Path(paths['noise_dir'])
            if path.exists():
                NOISE_TEST_DIR = path
                results['updated'].append(f"Noise directory: {path}")
                logger.info(f"Updated NOISE_TEST_DIR to: {path}")
            else:
                results['errors'].append(f"Noise directory not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating noise directory: {e}")
    
    # Update Lab Registry File
    if 'lab_registry' in paths and paths['lab_registry']:
        try:
            path = Path(paths['lab_registry'])
            if path.exists() and path.is_file():
                LAB_REGISTRY_FILE = path
                results['updated'].append(f"Lab registry file: {path}")
                logger.info(f"Updated LAB_REGISTRY_FILE to: {path}")
            else:
                results['errors'].append(f"Lab registry file not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating lab registry file: {e}")
    
    # Update Noise Registry File
    if 'noise_registry' in paths and paths['noise_registry']:
        try:
            path = Path(paths['noise_registry'])
            if path.exists() and path.is_file():
                NOISE_REGISTRY_FILE = path
                results['updated'].append(f"Noise registry file: {path}")
                logger.info(f"Updated NOISE_REGISTRY_FILE to: {path}")
            else:
                results['errors'].append(f"Noise registry file not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating noise registry file: {e}")
    
    # Update LF Registry File
    if 'lf_registry' in paths and paths['lf_registry']:
        try:
            path = Path(paths['lf_registry'])
            if path.exists() and path.is_file():
                LF_REGISTRY_FILE = path
                results['updated'].append(f"LF registry file: {path}")
                logger.info(f"✅ Updated LF_REGISTRY_FILE to: {path}")
            else:
                results['errors'].append(f"LF registry file not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating LF registry file: {e}")
    
    # Update LF Base Directory
    if 'lf_base_dir' in paths and paths['lf_base_dir']:
        try:
            path = Path(paths['lf_base_dir'])
            if path.exists() and path.is_dir():
                LF_BASE_DIR = path
                results['updated'].append(f"LF base directory: {path}")
                logger.info(f"✅ Updated LF_BASE_DIR to: {path}")
            else:
                results['errors'].append(f"LF base directory not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating LF base directory: {e}")
    
    # Update Test Lab Directory (CARICHI NOMINALI)
    if 'test_lab_dir' in paths and paths['test_lab_dir']:
        try:
            path = Path(paths['test_lab_dir'])
            if path.exists() and path.is_dir():
                TEST_LAB_CARICHI_DIR = path
                results['updated'].append(f"Test Lab directory: {path}")
                logger.info(f"✅ Updated TEST_LAB_CARICHI_DIR to: {path}")
            else:
                results['errors'].append(f"Test Lab directory not found: {path}")
        except Exception as e:
            results['errors'].append(f"Error updating Test Lab directory: {e}")
    
    # Update cache with new paths
    if results['updated']:
        try:
            cache = get_directory_cache()
            # Cache the exact paths for future use
            if PERFORMANCE_TEST_DIR:
                cache.cache_exact_path("ProveEffettuate", str(PERFORMANCE_TEST_DIR))
            if NOISE_TEST_DIR:
                cache.cache_exact_path("Tests Rumore", str(NOISE_TEST_DIR))
            if LAB_REGISTRY_FILE:
                cache.cache_exact_path("Registro LAB.xlsx", str(LAB_REGISTRY_FILE))
            if NOISE_REGISTRY_FILE:
                cache.cache_exact_path("REGISTRO RUMORE.xlsx", str(NOISE_REGISTRY_FILE))
            if LF_REGISTRY_FILE:
                cache.cache_exact_path("REGISTRO LF .xlsx", str(LF_REGISTRY_FILE))
            if LF_BASE_DIR:
                cache.cache_exact_path("RELIABIL", str(LF_BASE_DIR))
            if TEST_LAB_CARICHI_DIR:
                cache.cache_exact_path("CARICHI NOMINALI", str(TEST_LAB_CARICHI_DIR))
            
            logger.info("✅ Cache updated with new manual paths (including LF and Test Lab)")
        except Exception as e:
            logger.warning(f"Could not update cache: {e}")
    
    return results
