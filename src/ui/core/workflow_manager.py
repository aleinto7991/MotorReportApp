"""
Simplified workflow manager for the Motor Report GUI
Handles basic tab navigation and content refresh.
"""
import flet as ft
import logging
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI

logger = logging.getLogger(__name__)


class WorkflowManager:
    """Simplified workflow manager for tab navigation"""
    
    def __init__(self, gui: 'MotorReportAppGUI'):
        self.gui = gui
        self.current_step = 0
        self.step_names = {
            0: "Setup",
            1: "Search & Select", 
            2: "Configure",
            3: "Generate"
        }
    
    @property
    def state(self):
        """Get state manager from GUI"""
        return self.gui.state_manager
    
    def handle_tab_change(self, selected_index: int):
        """Handle tab change - simplified with no validation"""
        self.current_step = selected_index
        self._update_tab_titles()
        self._update_status_message()
        
        # Special handling for Configure tab - refresh when activated to pick up search results
        if selected_index == 2:  # Config tab
            self.refresh_tab('config')
        # Special handling for Generate tab - auto-apply config and force refresh when activated
        elif selected_index == 3:  # Generate tab
            self._auto_apply_config_and_refresh_generate_tab()
    
    def _update_tab_titles(self):
        """Update tab titles to show current progress"""
        if hasattr(self.gui, 'tabs') and self.gui.tabs and self.gui.tabs.tabs:
            for i, tab in enumerate(self.gui.tabs.tabs):
                if i < self.current_step:
                    # Completed step
                    tab.text = f"✓ {self.step_names[i]}"
                elif i == self.current_step:
                    # Current step
                    tab.text = f"→ {self.step_names[i]}"
                else:
                    # Available step
                    tab.text = f"{i + 1}. {self.step_names[i]}"
    
    def _update_status_message(self):
        """Update status message based on current step"""
        step_messages = {
            0: "Configure your input paths, then click the 'Search & Select' tab to continue.",
            1: "Search for tests and select the ones you want. Click the 'Configure' tab when ready.",
            2: "Configure report settings as needed. Click the 'Generate' tab when ready.",
            3: "Review settings and generate your motor performance report."
        }
        
        message = step_messages.get(self.current_step, "Follow the workflow steps.")
        if hasattr(self.gui, 'status_manager') and self.gui.status_manager:
            self.gui.status_manager.update_status(message, "blue")
    
    def go_to_step(self, step: int):
        """Navigate to a specific step"""
        self.current_step = step
        self.gui.tabs.selected_index = step
        
        # Refresh tab content if navigating to Configure tab (step 2)
        if step == 2:
            self.refresh_tab('config')
        # Auto-apply config and refresh if navigating to Generate tab (step 3)
        elif step == 3:
            self._auto_apply_config_and_refresh_generate_tab()
        
        self._update_tab_titles()
        self._update_status_message()
        self.gui._safe_page_update()
    
    def update_workflow_state(self):
        """Public method to update workflow state"""
        self._update_tab_titles()
        self._update_status_message()

    def refresh_tab(self, tab_name):
        """
        Unified tab refresh method.
        Args:
            tab_name: 'config' or 'generate'
        """
        try:
            logger.info(f"Refreshing {tab_name} tab...")
            
            if tab_name == 'config':
                self._refresh_config_tab_content()
            elif tab_name == 'generate':
                self._refresh_generate_tab_content()
            else:
                logger.warning(f"Unknown tab name: {tab_name}")
                
        except Exception as e:
            logger.error(f"Error refreshing {tab_name} tab: {e}")

    def _refresh_config_tab_content(self):
        """Internal method to refresh config tab content"""
        if hasattr(self.gui, 'tabs') and self.gui.tabs and len(self.gui.tabs.tabs) > 2:
            from ..tabs.config_tab import ConfigTab
            new_config_tab = ConfigTab(self.gui)
            new_config_content = new_config_tab.get_tab_content()
            self.gui.config_tab = new_config_tab
            self.gui.tabs.tabs[2].content = new_config_content
            self.gui._safe_page_update()
            logger.info("Config tab refreshed")
        else:
            logger.warning("Could not refresh Configure tab - tabs not initialized")

    def _refresh_generate_tab_on_activation(self):
        """Force refresh Generate tab content when tab becomes active"""
        try:
            logger.info("🔄 Generate tab activated - forcing content refresh...")
            
            if hasattr(self.gui, 'generate_tab') and self.gui.generate_tab:
                # Try multiple refresh strategies
                success = False
                
                # Strategy 1: Normal refresh
                if self.gui.generate_tab.refresh_content():
                    success = True
                    logger.info("✅ Generate tab refresh successful (normal)")
                else:
                    # Strategy 2: Force visibility
                    if self.gui.generate_tab.force_content_visibility():
                        success = True
                        logger.info("✅ Generate tab refresh successful (visibility forced)")
                
                if success:
                    # Apply extra page update for good measure
                    self.gui._safe_page_update()
                    
                    import time
                    time.sleep(0.02)  # Brief pause
                    self.gui._safe_page_update()
                else:
                    logger.warning("⚠️ Generate tab refresh failed on activation")
            else:
                logger.warning("⚠️ Generate tab not available for refresh")
                
        except Exception as e:
            logger.error(f"❌ Error refreshing Generate tab on activation: {e}")

    def _refresh_generate_tab_content(self):
        """Internal method to refresh generate tab content"""
        if hasattr(self.gui, 'tabs') and self.gui.tabs and len(self.gui.tabs.tabs) > 3:
            if hasattr(self.gui, 'generate_tab') and self.gui.generate_tab:
                # Use the generate tab's own refresh method instead of replacing content
                success = self.gui.generate_tab.refresh_content()
                if success:
                    self.gui._safe_page_update()
                    logger.info("Generate tab refreshed")
                else:
                    logger.warning("Generate tab refresh failed")
    
    def _auto_apply_config_and_refresh_generate_tab(self):
        """Auto-apply configuration and refresh Generate tab for seamless workflow"""
        try:
            logger.info("🔄 Auto-applying configuration and refreshing Generate tab...")
            
            # First, auto-apply the configuration (replaces the removed Apply button)
            if hasattr(self.gui, 'event_handlers') and self.gui.event_handlers:
                logger.info("📝 Auto-applying configuration settings...")
                self.gui.event_handlers.on_apply_config_selection()
                logger.info("✅ Configuration auto-applied successfully")
            else:
                logger.warning("⚠️ Event handlers not available for auto-apply")
            
            # Then refresh the Generate tab to show the updated configuration
            self._refresh_generate_tab_on_activation()
            
        except Exception as e:
            logger.error(f"❌ Error in auto-apply config and refresh: {e}")
            # Fallback to just refreshing the generate tab
            self._refresh_generate_tab_on_activation()

