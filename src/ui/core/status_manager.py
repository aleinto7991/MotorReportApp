"""
Status Manager - Handles status text and progress display for the Motor Report App GUI.

This manager provides centralized status and progress management with:
- Status message updates with color coding
- Progress bar visibility control
- Progress text updates
- Safe callback execution with shutdown handling

Relocated from src/gui/utils/helpers.py to src/gui/core/ for architectural
consistency with other managers (StateManager, WorkflowManager, SearchManager, etc.)
"""

import flet as ft
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)


class StatusManager:
    """
    Manages status text and progress display for the GUI.
    
    Responsibilities:
    - Update status messages with color coding
    - Show/hide progress bar
    - Update progress text during long operations
    - Safe callback execution with shutdown handling
    
    Thread Safety:
        All methods include exception handling for shutdown scenarios.
        Callbacks are executed safely with proper error handling.
    """
    
    def __init__(
        self,
        status_text: ft.Text,
        progress_bar: ft.ProgressBar,
        update_callback: Optional[Callable] = None,
        progress_text: Optional[ft.Text] = None,
        color_resolver: Optional[Callable[[Optional[str]], Optional[str]]] = None,
    ):
        """
        Initialize the status manager.
        
        Args:
            status_text: Flet Text control for status messages
            progress_bar: Flet ProgressBar control for progress indication
            update_callback: Optional callback to trigger page updates (typically _safe_page_update)
            progress_text: Optional Text control for progress messages
        """
        self.status_text = status_text
        self.progress_bar = progress_bar
        self.progress_text = progress_text
        self.update_callback = update_callback or (lambda: None)
        self._color_resolver = color_resolver
    
    def update_status(self, message: str, color: str = 'black'):
        """
        Update the status text with a message and color.
        
        Args:
            message: Status message to display
            color: Color for the status text (default: 'black')
                   Common values: 'green' (success), 'red' (error), 
                   'blue' (info), 'orange' (warning)
        
        Thread Safety:
            Handles RuntimeError and AttributeError during shutdown gracefully.
        """
        try:
            self.status_text.value = message
            self.status_text.color = self._resolve_color(color)
            logger.info(f"GUI Status Update: {message}")
            
            # Safe callback execution
            if self.update_callback:
                try:
                    self.update_callback()
                except (RuntimeError, AttributeError) as e:
                    if "shutdown" in str(e).lower() or "session" in str(e).lower():
                        logger.debug(f"Status update callback skipped due to shutdown: {e}")
                    else:
                        logger.warning(f"Status update callback failed: {e}")
        except Exception as e:
            logger.warning(f"Failed to update status: {e}")

    def _resolve_color(self, color: Optional[str]) -> Optional[str]:
        if callable(self._color_resolver):
            try:
                resolved = self._color_resolver(color)
                if resolved:
                    return resolved
            except Exception:
                pass
        return color
    
    def show_progress(self, message: str = ""):
        """
        Show the progress bar and optionally update progress text.
        
        Args:
            message: Optional progress message to display
        
        Thread Safety:
            Handles RuntimeError and AttributeError during shutdown gracefully.
        """
        try:
            self.progress_bar.visible = True
            if message and self.progress_text:
                self.progress_text.value = message
                self.progress_text.visible = True
            
            # Safe callback execution
            if self.update_callback:
                try:
                    result = self.update_callback()
                    if result is False:
                        logger.debug("Progress show update callback returned False - UI may not reflect changes yet")
                except (RuntimeError, AttributeError) as e:
                    if "shutdown" in str(e).lower() or "session" in str(e).lower():
                        logger.debug(f"Progress show callback skipped due to shutdown: {e}")
                    else:
                        logger.warning(f"Progress show callback failed: {e}")
        except Exception as e:
            logger.warning(f"Failed to show progress: {e}")
    
    def hide_progress(self):
        """
        Hide the progress bar and progress text.
        
        Thread Safety:
            Handles RuntimeError and AttributeError during shutdown gracefully.
        """
        try:
            self.progress_bar.visible = False
            if self.progress_text:
                self.progress_text.visible = False
            
            # Safe callback execution
            if self.update_callback:
                try:
                    result = self.update_callback()
                    if result is False:
                        logger.debug("Progress hide update callback returned False - UI may not reflect changes yet")
                except (RuntimeError, AttributeError) as e:
                    if "shutdown" in str(e).lower() or "session" in str(e).lower():
                        logger.debug(f"Progress hide callback skipped due to shutdown: {e}")
                    else:
                        logger.warning(f"Progress hide callback failed: {e}")
        except Exception as e:
            logger.warning(f"Failed to hide progress: {e}")
    
    def update_progress_text(self, message: str):
        """
        Update the progress text message.
        
        Args:
            message: Progress message to display
        
        Thread Safety:
            Handles RuntimeError and AttributeError during shutdown gracefully.
        """
        try:
            if self.progress_text:
                self.progress_text.value = message
                if not self.progress_text.visible:
                    self.progress_text.visible = True
            
            # Safe callback execution
            if self.update_callback:
                try:
                    self.update_callback()
                except (RuntimeError, AttributeError) as e:
                    if "shutdown" in str(e).lower() or "session" in str(e).lower():
                        logger.debug(f"Progress text update callback skipped due to shutdown: {e}")
                    else:
                        logger.warning(f"Progress text update callback failed: {e}")
        except Exception as e:
            logger.warning(f"Failed to update progress text: {e}")
    
    def set_progress_value(self, value: float, message: str = ""):
        """
        Set progress bar value (0.0 to 1.0) and optionally update message.
        
        Args:
            value: Progress value from 0.0 (0%) to 1.0 (100%)
            message: Optional progress message to display
        
        Thread Safety:
            Handles RuntimeError and AttributeError during shutdown gracefully.
        """
        try:
            # Clamp value between 0 and 1
            clamped_value = max(0.0, min(1.0, value))
            
            # Update progress bar
            self.progress_bar.value = clamped_value
            if not self.progress_bar.visible:
                self.progress_bar.visible = True
            
            # Update message if provided
            if message and self.progress_text:
                self.progress_text.value = message
                if not self.progress_text.visible:
                    self.progress_text.visible = True
            
            # Safe callback execution
            if self.update_callback:
                try:
                    self.update_callback()
                except (RuntimeError, AttributeError) as e:
                    if "shutdown" in str(e).lower() or "session" in str(e).lower():
                        logger.debug(f"Progress value update callback skipped due to shutdown: {e}")
                    else:
                        logger.warning(f"Progress value update callback failed: {e}")
        except Exception as e:
            logger.warning(f"Failed to set progress value: {e}")

