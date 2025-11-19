"""
Base classes for GUI components
"""
import flet as ft
import logging
from typing import Dict, Any, Optional, Callable
from abc import ABC, abstractmethod

try:
    from ..theme import resolve_token
except Exception:  # pragma: no cover - fallback when theme not available
    resolve_token = lambda page, token, fallback=None: fallback  # type: ignore

logger = logging.getLogger(__name__)


class BaseComponent(ABC):
    """Base class for all GUI components"""
    
    def __init__(self, parent_gui=None):
        self.parent_gui = parent_gui
        self._component = None
    
    @abstractmethod
    def build(self) -> ft.Control:
        """Build and return the Flet control for this component"""
        pass
    
    @property
    def component(self) -> ft.Control:
        """Get the built component, building it if necessary"""
        if self._component is None:
            self._component = self.build()
        return self._component
    
    def safe_update(self):
        """Safely update the component if parent GUI allows it"""
        if self.parent_gui and hasattr(self.parent_gui, '_safe_page_update'):
            self.parent_gui._safe_page_update()
        elif self.parent_gui and hasattr(self.parent_gui, 'page'):
            try:
                self.parent_gui.page.update()
            except Exception as e:
                logger.warning(f"Failed to update page: {e}")

    # Shared theme helpers -------------------------------------------------
    def theme_color(self, token: str, fallback: str) -> str:
        """Resolve a semantic theme color with graceful fallback."""
        try:
            if self.parent_gui and hasattr(self.parent_gui, '_themed_color'):
                return self.parent_gui._themed_color(token, fallback)
            page = getattr(self.parent_gui, 'page', None) if self.parent_gui else None
            return resolve_token(page, token, fallback) or fallback
        except Exception:
            return fallback


class BaseTab(BaseComponent):
    """Base class for tab implementations"""
    
    def __init__(self, parent_gui=None):
        super().__init__(parent_gui)
        self.tab_name = "Base Tab"
        self.tab_icon = ft.Icons.TAB
    
    @abstractmethod
    def get_tab_content(self) -> ft.Control:
        """Get the content for this tab"""
        pass
    
    def build(self) -> ft.Tab:
        """Build the tab with content"""
        return ft.Tab(
            text=self.tab_name,
            icon=self.tab_icon,
            content=self.get_tab_content()
        )


class ProgressIndicators:
    """Manages progress indicators for workflow steps"""
    
    def __init__(self):
        self.progress_rings = {}
        self.status_icons = {}
        
    def create_indicators(self, step_count: int):
        """Create progress rings and status icons for the given number of steps"""
        for i in range(1, step_count + 1):
            self.progress_rings[i] = ft.ProgressRing(
                width=16, 
                height=16, 
                stroke_width=2, 
                visible=False
            )
            self.status_icons[i] = ft.Icon(
                ft.Icons.CHECK_CIRCLE, 
                color="green", 
                size=16, 
                visible=False
            )
    
    def show_progress(self, step: int):
        """Show progress indicator for a step"""
        if step in self.progress_rings:
            self.progress_rings[step].visible = True
            self.status_icons[step].visible = False
    
    def hide_progress(self, step: int):
        """Hide progress indicator for a step"""
        if step in self.progress_rings:
            self.progress_rings[step].visible = False
    
    def show_success(self, step: int):
        """Show success indicator for a step"""
        if step in self.progress_rings and step in self.status_icons:
            self.progress_rings[step].visible = False
            self.status_icons[step].visible = True
    
    def get_indicators_for_step(self, step: int) -> tuple:
        """Get progress ring and status icon for a step"""
        return (
            self.progress_rings.get(step), 
            self.status_icons.get(step)
        )

