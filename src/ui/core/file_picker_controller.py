"""
File Picker Controller for Motor Report GUI

Handles all file and folder selection events, including:
- Folder selection (tests, noise)
- File selection (registry files)
- Context-aware path extraction
- UI updates after selection
"""

import flet as ft
import logging
import threading
import os
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI
    from .state_manager import StateManager

logger = logging.getLogger(__name__)


class FilePickerController:
    """
    Controller for file picker events and path management.
    
    Handles:
    - Folder selection with context awareness
    - File selection with automatic folder extraction
    - Path validation and UI updates
    - Backend configuration synchronization
    """
    
    def __init__(self, gui: 'MotorReportAppGUI', state_manager: 'StateManager'):
        """
        Initialize the file picker controller.
        
        Args:
            gui: Reference to the main GUI instance
            state_manager: Reference to the state manager
        """
        self.gui = gui
        self.state_manager = state_manager
    
    @property
    def state(self):
        """Get the application state."""
        return self.state_manager.state
    
    def on_folder_picked(self, e: ft.FilePickerResultEvent):
        """
        Handle folder selection event.
        
        Uses picker_context from state to determine which folder type was selected:
        - 'noise_folder': Noise test data folder
        - 'test_folder': Performance test folder
        - default: Performance test folder (for backward compatibility)
        
        Args:
            e: File picker result event containing selected path
        """
        logger.debug(f"on_folder_picked called with path: {e.path}")
        
        if not e.path:
            logger.debug("No path selected in folder picker")
            return
        
        # Determine which folder was selected based on context
        context = getattr(self.state, 'picker_context', '')
        logger.debug(f"picker_context is: {context}")
        
        if context == "noise_folder":
            self._update_noise_folder(e.path)
        elif context == "test_folder":
            self._update_tests_folder(e.path)
        else:
            # Default to performance test folder for backward compatibility
            logger.debug("Updating test folder path (default context)")
            self._update_tests_folder(e.path)
        
        # Clear context and update UI
        self._finalize_folder_selection()
    
    def on_registry_file_picked(self, e: ft.FilePickerResultEvent):
        """
        Handle registry file selection event.
        
        Uses picker_context to determine handling:
        - 'noise_registry': Noise registry file
        - 'test_folder_via_file': Extract folder from file path for tests
        - 'noise_folder_via_file': Extract folder from file path for noise
        - 'save_report_via_file': Extract folder for report saving
        - default: Performance registry file
        
        Args:
            e: File picker result event containing selected file path
        """
        logger.debug(f"on_registry_file_picked called with path: {e.path}")
        
        if not e.path:
            logger.debug("No file selected")
            return
        
        # Determine file type based on context
        context = getattr(self.state, 'picker_context', '')
        logger.debug(f"picker_context is: {context}")
        
        if context == "noise_registry":
            self._update_noise_registry(e.path)
        elif context == "test_folder_via_file":
            self._extract_and_update_tests_folder(e.path)
        elif context == "noise_folder_via_file":
            self._extract_and_update_noise_folder(e.path)
        elif context == "save_report_via_file":
            self._extract_and_save_report(e.path)
        else:
            # Default to performance registry file
            self._update_registry_file(e.path)
        
        # Clear context and update UI
        self._finalize_file_selection()
    
    # Private helper methods for specific path types
    
    def _update_noise_folder(self, path: str):
        """
        Update noise folder path in state and refresh UI.
        
        Args:
            path: Absolute path to the noise test folder
            
        Side Effects:
            - Updates state manager with new noise folder path
            - Updates UI text to show selected path (green, non-italic)
        """
        logger.debug("Updating noise folder path")
        self.state_manager.update_paths(noise_folder=path)
        self.gui.noise_folder_path_text.value = f"Selected: {path}"
        self.gui.noise_folder_path_text.color = "green"
        self.gui.noise_folder_path_text.italic = False
    
    def _update_tests_folder(self, path: str):
        """
        Update tests folder path in state and refresh UI.
        
        Args:
            path: Absolute path to the performance test folder
            
        Side Effects:
            - Updates state manager with new tests folder path
            - Updates UI text to show selected path (green, non-italic)
            - May trigger test scanning workflow
        """
        logger.debug("Updating test folder path")
        self.state_manager.update_paths(tests_folder=path)
        self.gui.tests_folder_path_text.value = f"Selected: {path}"
        self.gui.tests_folder_path_text.color = "green"
        self.gui.tests_folder_path_text.italic = False
    
    def _update_noise_registry(self, path: str):
        """
        Update noise registry file path in state and refresh UI.
        
        Args:
            path: Absolute path to the noise registry Excel file (REGISTRO RUMORE.xlsx)
            
        Side Effects:
            - Updates state manager with new noise registry path
            - Updates UI text to show selected path (green, non-italic)
        """
        logger.debug("Updating noise registry path")
        self.state_manager.update_paths(noise_registry=path)
        self.gui.noise_registry_path_text.value = f"Selected: {path}"
        self.gui.noise_registry_path_text.color = "green"
        self.gui.noise_registry_path_text.italic = False
    
    def _update_registry_file(self, path: str):
        """
        Update performance registry file path in state and refresh UI.
        
        Args:
            path: Absolute path to the performance registry Excel file
            
        Side Effects:
            - Updates state manager with new registry file path
            - Updates UI text to show selected path (green, non-italic)
        """
        logger.debug("Updating registry file path")
        self.state_manager.update_paths(registry_file=path)
        self.gui.registry_file_path_text.value = f"Selected: {path}"
        self.gui.registry_file_path_text.color = "green"
        self.gui.registry_file_path_text.italic = False
    
    def _extract_and_update_tests_folder(self, file_path: str):
        """
        Extract parent folder from a file path and update tests folder.
        
        Useful when user selects a file within the tests folder rather than
        the folder itself (e.g., via file picker dialog).
        
        Args:
            file_path: Absolute path to a file within the tests folder
            
        Side Effects:
            - Extracts parent directory using os.path.dirname()
            - Updates state manager with extracted folder path
            - Updates UI to show extracted folder path
        """
        folder_path = os.path.dirname(file_path)
        logger.debug(f"Extracted folder path from file: {folder_path}")
        self.state_manager.update_paths(tests_folder=folder_path)
        self.gui.tests_folder_path_text.value = f"Selected: {folder_path}"
        self.gui.tests_folder_path_text.color = "green"
        self.gui.tests_folder_path_text.italic = False
    
    def _extract_and_update_noise_folder(self, file_path: str):
        """
        Extract parent folder from a file path and update noise folder.
        
        Useful when user selects a file within the noise folder rather than
        the folder itself (e.g., via file picker dialog).
        
        Args:
            file_path: Absolute path to a file within the noise folder
            
        Side Effects:
            - Extracts parent directory using os.path.dirname()
            - Updates state manager with extracted folder path
            - Updates UI to show extracted folder path
        """
        folder_path = os.path.dirname(file_path)
        logger.debug(f"Extracted folder path from file: {folder_path}")
        self.state_manager.update_paths(noise_folder=folder_path)
        self.gui.noise_folder_path_text.value = f"Selected: {folder_path}"
        self.gui.noise_folder_path_text.color = "green"
        self.gui.noise_folder_path_text.italic = False
    
    def _extract_and_save_report(self, file_path: str):
        """
        Extract folder from file path and trigger report save operation.
        
        Used when user selects a file path to determine the save location
        for the generated report.
        
        Args:
            file_path: Absolute path to the file/location for report saving
            
        Side Effects:
            - Extracts parent directory
            - Delegates to report generation controller's save_report_to_folder()
            - Logs warning if report generation controller unavailable
        """
        folder_path = os.path.dirname(file_path)
        logger.debug(f"Extracted save folder path: {folder_path}")
        
        # Delegate to report generation controller
        if hasattr(self.gui.event_handlers, 'report_generation_controller'):
            self.gui.event_handlers.report_generation_controller.save_report_to_folder(folder_path)
        else:
            logger.warning("Report generation controller not available")
    
    def _finalize_folder_selection(self):
        """
        Finalize folder selection workflow.
        
        Side Effects:
            - Clears picker_context from state
            - Triggers UI update via _safe_page_update()
            - Initiates backend configuration update in background thread
        """
        self.state.picker_context = ""
        self.gui._safe_page_update()
        self._trigger_backend_config_update()
    
    def _finalize_file_selection(self):
        """
        Finalize file selection workflow.
        
        Side Effects:
            - Clears picker_context from state
            - Triggers UI update via _safe_page_update()
            - Initiates backend configuration update in background thread
        """
        self.state.picker_context = ""
        self.gui._safe_page_update()
        self._trigger_backend_config_update()
    
    def _trigger_backend_config_update(self):
        """
        Trigger backend configuration update in a background thread.
        
        Ensures the backend MotorReportApp instance is updated with new
        file/folder paths without blocking the UI thread.
        
        Side Effects:
            - Spawns daemon thread running _update_backend_config()
            - Logs debug message on successful thread start
            - Does nothing if GUI app instance is unavailable
        """
        if hasattr(self.gui, 'app') and self.gui.app:
            run_in_background(self.gui._update_backend_config)
            logger.debug("Triggered backend config update in background thread")

