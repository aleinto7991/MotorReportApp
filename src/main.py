#!/usr/bin/env python3
"""Motor Report Application - Web entry point with structured logging."""

import argparse
import logging
import sys
from pathlib import Path

import flet as ft


logger = logging.getLogger(__name__)


def _configure_logging() -> None:
    """Ensure structured logging is configured once for the web entry point."""
    root_logger = logging.getLogger()
    if root_logger.handlers:
        # Respect existing configuration (e.g., when invoked from tests)
        if root_logger.level > logging.DEBUG:
            root_logger.setLevel(logging.DEBUG)
        return

    logs_dir = Path(__file__).resolve().parent.parent / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    file_handler = logging.FileHandler(
        logs_dir / "motor_report_main.log", encoding="utf-8"
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    root_logger.setLevel(logging.DEBUG)
    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)


_configure_logging()


# Add project root to sys.path for imports
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

# Bootstrap runtime environment (handles PyInstaller, paths, stdio)
try:
    from .config.runtime import log_runtime_info, is_bundled
except ImportError as e:
    logger.critical(f"Failed to import runtime module: {e}")
    sys.exit(1)

# Log runtime information for debugging
log_runtime_info()

# Initialize directories at runtime (not at import time during build)
try:
    logger.info("Initializing data directories at runtime...")
    from .config.directory_config import ensure_directories_initialized
    ensure_directories_initialized()
    logger.info("Data directories initialized successfully")
except Exception as e:
    logger.error(f"Failed to initialize data directories: {e}")
    import traceback
    traceback.print_exc()
    # Continue anyway - the GUI will show path configuration options

# Import the GUI module
try:
    logger.info("Importing MotorReportAppGUI...")
    from .ui.main_gui import MotorReportAppGUI
    logger.info("MotorReportAppGUI import complete")
except ImportError as e:
    logger.exception(f"Unable to import MotorReportAppGUI: {e}")
    logger.error("Make sure the application is run from the project root directory")
    sys.exit(1)


def main(page: ft.Page):
    """Main entry point for the web application."""
    try:
        # Configure page for web
        page.title = "Motor Report Generator - Web App"
        page.padding = 10
        page.scroll = ft.ScrollMode.AUTO
        
        logger.info("Initializing MotorReportAppGUI for Flet page")
        # Initialize and run the GUI
        app = MotorReportAppGUI(page)
        logger.debug("MotorReportAppGUI instantiated: %s", app)
        # The app is already initialized and built in the constructor
    except Exception as e:
        logger.exception("Error while initializing the GUI page")
        # Add a simple error message to the page
        page.add(ft.Text(f"Error loading application: {e}", color="red"))
        page.update()


def kill_all_processes_on_port(port):
    """Kill ALL processes using the specified port (aggressive cleanup)"""
    killed_count = 0
    try:
        import subprocess
        logger.info("Scanning for processes on port %s", port)
        
        # Find all processes using the port
        result = subprocess.run(
            f'netstat -ano | findstr :{port}',
            shell=True, capture_output=True, text=True
        )
        
        if result.stdout:
            lines = result.stdout.strip().split('\n')
            pids_to_kill = set()  # Use set to avoid duplicates
            
            for line in lines:
                if f':{port}' in line:
                    parts = line.split()
                    if len(parts) >= 5:
                        pid = parts[-1]
                        if pid.isdigit():
                            pids_to_kill.add(pid)
            
            if pids_to_kill:
                logger.info(
                    "Found %d process(es) using port %s: %s",
                    len(pids_to_kill),
                    port,
                    sorted(pids_to_kill),
                )
                
                for pid in pids_to_kill:
                    try:
                        logger.debug("Killing process %s", pid)
                        subprocess.run(f'taskkill /F /PID {pid}', shell=True, check=True)
                        logger.debug("Successfully killed process %s", pid)
                        killed_count += 1
                    except subprocess.CalledProcessError as e:
                        logger.warning("Failed to kill process %s directly: %s", pid, e)
                        # Try alternative method
                        try:
                            subprocess.run(f'wmic process where ProcessId={pid} delete', shell=True, check=True)
                            logger.debug("Successfully killed process %s via WMIC", pid)
                            killed_count += 1
                        except subprocess.CalledProcessError:
                            logger.error("Could not kill process %s with any method", pid)
            else:
                logger.info("No processes found using port %s", port)
        else:
            logger.info("Port %s is already free", port)
            
    except Exception as e:
        logger.exception("Error scanning/killing processes on port %s", port)
    
    if killed_count > 0:
        logger.info("Killed %d process(es) on port %s", killed_count, port)
        # Wait a moment for processes to fully terminate
        import time
        time.sleep(2)
    
    return killed_count


if __name__ == "__main__":
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Motor Report Generator')
    parser.add_argument('--port', type=int, default=8080, help='Port to run on (default: 8080)')
    parser.add_argument('--no-browser', action='store_true', help='Run in headless mode without browser')
    args = parser.parse_args()
    
    port = args.port
    
    # Run as web app that opens in browser
    logger.info("Starting Motor Report Web Application")
    logger.info(f"Application will open in the default web browser at http://localhost:{port}")
    
    # Kill any existing processes on the specified port before starting
    logger.info(f"Cleaning up any existing processes on port {port}")
    kill_all_processes_on_port(port)
    
    try:
        # Disable uvicorn logging in packaged environment (handled by runtime.py)
        if is_bundled():
            import uvicorn.config
            try:
                # Minimal logging config for bundled mode
                uvicorn.config.LOGGING_CONFIG = {
                    "version": 1,
                    "disable_existing_loggers": False,
                    "loggers": {
                        "uvicorn": {"level": "CRITICAL"},
                        "uvicorn.error": {"level": "CRITICAL"},
                        "uvicorn.access": {"level": "CRITICAL"},
                    },
                }
            except Exception:
                pass  # Ignore if we can't modify logging config
        
        # Determine view mode based on --no-browser flag
        view_mode = ft.AppView.FLET_APP if args.no_browser else ft.AppView.WEB_BROWSER
        
        ft.app(
            target=main, 
            port=port, 
            view=view_mode, 
            assets_dir="assets"
        )
        logger.info(f"Motor Report Web Application started on port {port}")
    except Exception:
        logger.exception("Error starting web application")
        # Don't use input() in packaged environment as it causes issues
        if not is_bundled():
            input("Press Enter to exit...")
        else:
            import time
            time.sleep(5)  # Give user time to see error
