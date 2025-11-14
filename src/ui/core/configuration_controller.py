"""
Configuration Controller - Handles SAP checkbox selection and test lab visibility management.

This controller manages:
- SAP code checkbox selections (comparison/noise)
- Test lab container visibility toggling
- Configuration state synchronization
- Auto-refresh and rebuild of missing containers

Extracted from EventHandlers as part of controller separation pattern.
"""

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI
    from .state_manager import StateManager

logger = logging.getLogger(__name__)


class ConfigurationController:
    """
    Handles configuration-related event handlers for SAP selections.
    
    Responsibilities:
    - Handle SAP checkbox selections (comparison and noise)
    - Manage test lab container visibility
    - Update test lab checkboxes when SAPs are selected
    - Auto-refresh config tab if containers are missing
    
    Thread Safety:
        All methods are called from the Flet UI thread. Background updates
        use daemon threads and update GUI via _safe_page_update().
    """
    
    def __init__(self, gui: 'MotorReportAppGUI', state: 'StateManager'):
        """
        Initialize configuration controller.
        
        Args:
            gui: Reference to main GUI for UI updates
            state: Reference to StateManager for configuration state
        """
        self.gui = gui
        self.state = state
        
        logger.debug("ConfigurationController initialized")
    
    def on_sap_checked(self, sap_type: str):
        """
        Create SAP checkbox handler for specific type (comparison or noise).
        
        This is a handler factory that returns a closure for handling checkbox events.
        When a SAP checkbox is checked/unchecked:
        1. Updates state via StateManager
        2. Shows/hides corresponding test lab selection container
        3. Updates test lab checkboxes if SAP is newly selected
        4. Auto-refreshes config tab if containers are missing
        
        Args:
            sap_type: Type of SAP - either "comparison" or "noise"
        
        Returns:
            Event handler function for checkbox state changes
        
        Thread Safety:
            Handlers run on Flet UI thread. Uses _safe_page_update for UI updates.
        """
        def handler(e):
            sap_code = e.control.data
            selected = e.control.value
            
            logger.debug(f"SAP {sap_code} ({sap_type}) checked: {selected}")
            
            # Update state
            self.state.update_sap_selection(sap_type, sap_code, selected)
            
            # Handle visibility and checkboxes based on SAP type
            if sap_type == "comparison":
                self._handle_comparison_sap_visibility(sap_code, selected)
            elif sap_type == "noise":
                self._handle_noise_sap_visibility(sap_code, selected)
        
        return handler
    
    def _handle_comparison_sap_visibility(self, sap_code: str, selected: bool):
        """
        Handle visibility of comparison test lab containers.
        
        When a comparison SAP is selected/deselected:
        1. Shows/hides the test lab selection container for that SAP
        2. Updates test lab checkboxes if newly selected
        3. Refreshes config tab if container is missing
        
        Args:
            sap_code: SAP code for comparison
            selected: Whether the SAP checkbox is checked
        """
        logger.debug(f"Handling comparison SAP {sap_code}, selected: {selected}")
        
        if not hasattr(self.gui, 'config_tab'):
            logger.debug("No config_tab attribute on GUI")
            return
        
        config_tab = self.gui.config_tab
        logger.debug(f"Config tab found: {config_tab is not None}")
        
        if not hasattr(config_tab, 'comparison_test_lab_containers'):
            logger.debug("Config tab does not have comparison_test_lab_containers attribute")
            return
        
        containers = config_tab.comparison_test_lab_containers
        logger.debug(f"Available containers: {list(containers.keys())}")
        
        if sap_code in containers:
            # Container exists - update visibility
            container = containers[sap_code]
            logger.debug(f"Container found for {sap_code}, setting visible={selected}")
            container.visible = selected
            
            # If the SAP is now selected, update the test lab checkboxes
            if selected:
                logger.debug(f"Updating test lab checkboxes for {sap_code}")
                config_tab._update_test_lab_checkboxes(sap_code)
            
            # Update the page
            logger.debug("Updating page")
            self.gui._safe_page_update()
        else:
            # Container missing - attempt to rebuild
            logger.debug(f"Container not found for {sap_code}")
            logger.debug(f"Expected containers: {self.state.state.found_sap_codes}")
            logger.debug("Will try to build content again to create missing container")
            
            self._rebuild_missing_container(sap_code, selected, containers, config_tab, 'comparison')
    
    def _handle_noise_sap_visibility(self, sap_code: str, selected: bool):
        """
        Handle visibility of noise test lab containers.
        
        When a noise SAP is selected/deselected:
        1. Shows/hides the noise test selection container for that SAP
        2. Updates noise test checkboxes if newly selected
        3. Refreshes config tab if container is missing
        
        Args:
            sap_code: SAP code for noise testing
            selected: Whether the SAP checkbox is checked
        """
        logger.debug(f"Handling noise SAP {sap_code}, selected: {selected}")
        
        if not hasattr(self.gui, 'config_tab'):
            logger.debug("No config_tab attribute on GUI")
            return
        
        config_tab = self.gui.config_tab
        logger.debug(f"Config tab found: {config_tab is not None}")
        
        if not hasattr(config_tab, 'noise_test_lab_containers'):
            logger.debug("Config tab does not have noise_test_lab_containers attribute")
            return
        
        containers = config_tab.noise_test_lab_containers
        logger.debug(f"Available noise containers: {list(containers.keys())}")
        
        if sap_code in containers:
            # Container exists - update visibility
            container = containers[sap_code]
            logger.debug(f"Noise container found for {sap_code}, setting visible={selected}")
            container.visible = selected
            
            # If the SAP is now selected, update the noise test checkboxes
            if selected:
                logger.debug(f"Updating noise test checkboxes for {sap_code}")
                config_tab._update_noise_test_checkboxes(sap_code)
            
            # Update the page
            logger.debug("Updating page")
            self.gui._safe_page_update()
        else:
            # Container missing - attempt to rebuild
            logger.debug(f"Noise container not found for {sap_code}")
            self._rebuild_missing_container(sap_code, selected, containers, config_tab, 'noise')
    
    def _rebuild_missing_container(self, sap_code: str, selected: bool, 
                                   containers: dict, config_tab, container_type: str):
        """
        Attempt to rebuild missing container by refreshing config tab.
        
        If a container is expected but not found, this triggers a refresh
        of the config tab to recreate all containers, then retries the
        visibility update.
        
        Args:
            sap_code: SAP code for the missing container
            selected: Desired visibility state
            containers: Current containers dictionary
            config_tab: Reference to config tab instance
            container_type: Either 'comparison' or 'noise'
        """
        try:
            if not hasattr(self.gui, 'workflow_manager'):
                logger.warning("No workflow_manager available for refresh")
                return
            
            logger.debug("Refreshing config tab to recreate containers...")
            self.gui.workflow_manager.refresh_tab('config')
            
            # After refresh, try again
            new_config_tab = self.gui.config_tab
            if not new_config_tab:
                logger.warning("Config tab not available after refresh")
                return
            
            # Get the appropriate containers based on type
            container_attr = f'{container_type}_test_lab_containers'
            if not hasattr(new_config_tab, container_attr):
                logger.warning(f"Config tab still missing {container_attr} after refresh")
                return
            
            new_containers = getattr(new_config_tab, container_attr)
            logger.debug(f"After refresh, available containers: {list(new_containers.keys())}")
            
            if sap_code in new_containers:
                # Successfully recreated - update visibility
                container = new_containers[sap_code]
                container.visible = selected
                
                if selected:
                    # Update checkboxes based on type
                    if container_type == 'comparison':
                        new_config_tab._update_test_lab_checkboxes(sap_code)
                    else:  # noise
                        new_config_tab._update_noise_test_checkboxes(sap_code)
                
                self.gui._safe_page_update()
                logger.debug(f"Successfully handled {sap_code} after refresh")
            else:
                logger.warning(f"Container for {sap_code} still not found after refresh")
                
        except Exception as refresh_error:
            logger.error(f"Error during config tab refresh: {refresh_error}", exc_info=True)

