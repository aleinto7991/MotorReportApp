"""
Configuration Tab - Third tab for configuring report options
"""
from typing import List
import logging

import flet as ft
from ..components.base import BaseTab
from ...services.noise_registry_loader import NoiseRegistryLoader
from ..utils.thread_pool import run_in_background

logger = logging.getLogger(__name__)


class ConfigTab(BaseTab):
    """Tab for configuring report options"""
    
    def __init__(self, parent_gui=None):
        super().__init__(parent_gui)
        self.tab_name = "3. Configure"
        self.tab_icon = ft.Icons.TUNE
        self.comparison_test_lab_containers = {}  # Track containers for each SAP code
        self.noise_test_lab_containers = {}  # Track containers for noise test selection
        self._content_built = False  # Flag to track if content has been built
        
        # Initialize noise registry loader service
        self.noise_registry_loader = NoiseRegistryLoader()
        
        # Legacy cache variables (kept for backward compatibility during migration)
        self._noise_registry_cache = {}  # Cache noise registry data to avoid repeated reads
        self._cached_noise_tests = None  # Cache full noise tests for detailed UI (loaded when needed)
        
        # Multiple comparison groups management
        self.comparison_groups = []  # List of comparison group containers
        self.comparison_counter = 0  # Counter for generating unique group IDs
        self.comparison_groups_container = None  # Container that holds all comparison groups
        
        # PERFORMANCE OPTIMIZATION: Start preloading noise cache immediately for better responsiveness
        self._preload_noise_registry_async()

    def _color(self, token: str, fallback: str) -> str:
        """Convenience wrapper around BaseComponent.theme_color."""
        return self.theme_color(token, fallback)
    
    def _clear_noise_cache(self):
        """Clear noise data cache to force reload"""
        self.noise_registry_loader.clear_cache()
        self._cached_noise_tests = None
        self._cache_timestamp = None
    
    def get_tab_content(self) -> ft.Control:
        """Build the configuration tab content (Step 3) - Optimized for performance"""
        try:
            logger.debug("ConfigTab: Starting to build tab content (fast mode)...")
            
            # Get progress indicators from parent
            step2_progress = None
            step2_status = None
            if self.parent_gui and hasattr(self.parent_gui, 'progress_indicators'):
                step2_progress, step2_status = self.parent_gui.progress_indicators.get_indicators_for_step(2)

            # Get ALL SAP codes that were found in the search (not just from selected tests)
            found_sap_codes = []
            if self.parent_gui and hasattr(self.parent_gui, 'state_manager'):
                # Use all SAP codes found during search
                found_sap_codes = self.parent_gui.state_manager.state.found_sap_codes.copy()
                logger.debug(f"ConfigTab: Found {len(found_sap_codes)} SAP codes: {found_sap_codes}")
                
                # Also check selected tests for additional context
                selected_tests = self.parent_gui.state_manager.state.selected_tests
                logger.debug(f"ConfigTab: {len(selected_tests)} selected tests")
                if selected_tests:
                    selected_sap_codes = list(set(test.sap_code for test in selected_tests.values() if test.sap_code))
                    logger.debug(f"ConfigTab: Selected tests span SAP codes: {selected_sap_codes}")
            else:
                logger.debug("ConfigTab: No state manager or parent GUI found")
                
            # Local helper for semantic colors
            color = self._color

            # Initialize sections (fast generation without heavy computation)
            noise_sap_sections = []
            lf_sap_sections = []
            comparison_sap_sections = []
            noise_info_msgs = []
            
            logger.info("=" * 70)
            logger.info("🔧 CONFIG TAB INITIALIZATION")
            logger.info("=" * 70)
            logger.info(f"📋 parent_gui exists: {bool(self.parent_gui)}")
            logger.info(f"📋 found_sap_codes: {found_sap_codes}")
            logger.info(f"📋 Number of SAP codes: {len(found_sap_codes)}")
            
            if self.parent_gui and found_sap_codes:
                logger.info("✅ Conditions met - generating sections...")
                # OPTIMIZATION 1: Generate UI quickly without heavy noise registry checks
                # We'll load noise data lazily when sections are expanded
                
                # Fast comparison SAP sections generation
                comparison_sap_sections = self._generate_comparison_sections_fast(found_sap_codes)
                
                # Fast noise SAP sections generation (without registry lookup)
                noise_sap_sections = self._generate_noise_sections_fast(found_sap_codes)
                
                # Fast LF SAP sections generation
                lf_sap_sections = self._generate_lf_sections_fast(found_sap_codes)
                
                # Start background noise registry preloading for later use
                self._preload_noise_registry_async()
                
                logger.info(f"✅ Generated sections - Noise: {len(noise_sap_sections)}, LF: {len(lf_sap_sections)}, Comparison: {len(comparison_sap_sections)}")
                            
            else:
                logger.warning("⚠️ Conditions NOT met for section generation:")
                logger.warning(f"   parent_gui: {bool(self.parent_gui)}")
                logger.warning(f"   found_sap_codes: {found_sap_codes}")
                logger.debug("ConfigTab: No parent GUI or no SAP codes found")
                # Initialize empty sections
                noise_sap_sections = []
                lf_sap_sections = []
                comparison_sap_sections = []

            # Create the fast UI (no heavy operations during initial render)
            none_noise_btn = ft.ElevatedButton(
                "None (Disable Noise)",
                icon=ft.Icons.CLOSE,
                bgcolor=color('surface_variant', '#bdbdbd'),
                color=color('on_surface', 'black'),
                on_click=lambda e: self._disable_noise_feature()
            )
            none_comparison_btn = ft.ElevatedButton(
                "None (Disable Comparison)",
                icon=ft.Icons.CLOSE,
                bgcolor=color('surface_variant', '#bdbdbd'),
                color=color('on_surface', 'black'),
                on_click=lambda e: self._disable_comparison_feature()
            )
            
            # Create reference for comparison groups container
            comparison_groups_ref = ft.Ref[ft.Container]()
            self.comparison_groups_container = comparison_groups_ref

            logger.debug("ConfigTab: Returning container with content")
            return ft.Container(
                content=ft.Column([
                    ft.Row([
                        ft.Text("Configure SAP Codes for Each Feature", size=20, weight=ft.FontWeight.BOLD),
                        step2_progress or ft.Container(),
                        step2_status or ft.Container(),
                    ], spacing=10),
                    ft.Text(
                        "Select which SAP codes to use for each feature. Only SAP codes with test labs selected in the previous step are shown. Performance reports will include all selected SAP codes automatically.",
                        color=color('text_muted', 'white'),
                        size=14
                    ),
                    self._build_unit_selection_section(),
                    ft.Divider(),
                    ft.Text("Noise SAP Codes:", color=color('success', '#388e3c'), size=16, weight=ft.FontWeight.W_500),
                    ft.Text(
                        "Select SAP codes for noise tests. When checked, you can choose specific tests from the noise registry. Tests are unchecked by default.",
                        color=color('text_muted', 'grey'),
                        size=12
                    ),
                    ft.Column(noise_sap_sections, spacing=5),
                    *noise_info_msgs,
                    none_noise_btn,
                    ft.Divider(),
                    ft.Text("🔬 Life Test (LF) SAP Codes:", color=color('primary', '#1976d2'), size=16, weight=ft.FontWeight.W_500),
                    ft.Text(
                        "Select SAP codes for Life Test data. When checked, you can choose specific LF tests from the registry. The report will include hyperlinks to the test files.",
                        color=color('text_muted', 'grey'),
                        size=12
                    ),
                    ft.Column(lf_sap_sections, spacing=5),
                    ft.Divider(),
                    ft.Text("Comparison SAP Codes:", color=color('warning', '#fbc02d'), size=16, weight=ft.FontWeight.W_500),
                    ft.Text(
                        "Create multiple comparison groups. Each group can include tests from any SAP codes. Test labs can be reused across different comparisons.",
                        color=color('text_muted', 'grey'),
                        size=12
                    ),
                    
                    # Container for all comparison groups
                    ft.Container(
                        content=ft.Column([], spacing=10),
                        ref=comparison_groups_ref
                    ),
                    
                    # Add Comparison button at the bottom
                    ft.Container(
                        content=ft.ElevatedButton(
                            "Add Comparison",
                            icon=ft.Icons.ADD,
                            bgcolor=color('primary', 'blue'),
                            color=color('on_primary', 'white'),
                            on_click=self._add_new_comparison_group
                        ),
                        padding=ft.padding.symmetric(vertical=10)
                    ),
                    
                    # Note: Removed Apply/Clear buttons for seamless workflow
                    # Config navigation controls are no longer needed with the new workflow
                    
                    # Dynamic navigation controls will be added by main GUI
                    ft.Container(ref=ft.Ref[ft.Container](), height=20),  # Reduced spacing for cleaner look
                    ft.Container(
                        content=ft.Row([]),
                        padding=ft.padding.only(top=20)
                    )
                ], spacing=15),
                padding=ft.padding.all(20),
                expand=True
            )
            
        except Exception as e:
            logger.error(f"Error building ConfigTab content: {e}")
            import traceback
            traceback.print_exc()
            
            # Return a simple error message container
            return ft.Container(
                content=ft.Column([
                    ft.Text("❌ Error loading configuration tab", size=20, color=self.theme_color('error', 'red')),
                    ft.Text(f"Error: {str(e)}", size=14, color=self.theme_color('error', 'red')),
                    ft.Text("Please check the console for details.", size=12, color=self.theme_color('text_muted', 'grey'))
                ], spacing=10),
                padding=ft.padding.all(20),
                expand=True
            )

    def _build_unit_selection_section(self) -> ft.Control:
        """Create the measurement unit selection controls."""
        state = None
        if self.parent_gui and hasattr(self.parent_gui, "state_manager"):
            state = self.parent_gui.state_manager.state

        pressure_dropdown = self._create_unit_dropdown(
            label="Pressure Unit",
            value=state.pressure_unit if state else "kPa",
            config_key="pressure_unit",
            options=["kPa", "mmH2O", "psi"]
        )

        flow_dropdown = self._create_unit_dropdown(
            label="Air Flow Unit",
            value=state.flow_unit if state else "m³/h",
            config_key="flow_unit",
            options=["m³/h", "l/s", "CFM"]
        )

        speed_dropdown = self._create_unit_dropdown(
            label="Speed Unit",
            value=state.speed_unit if state else "rpm",
            config_key="speed_unit",
            options=["rpm", "rps"]
        )

        power_dropdown = self._create_unit_dropdown(
            label="Power Unit",
            value=state.power_unit if state else "W",
            config_key="power_unit",
            options=["W", "kW", "HP"]
        )

        description = ft.Text(
            "Choose how measurements are displayed in charts and Excel exports. Changes apply immediately to new reports.",
            color=self.theme_color('text_muted', 'grey'),
            size=12,
        )

        return ft.Container(
            content=ft.Column([
                ft.Text(
                    "Measurement Units",
                    color=self.theme_color('primary', '#1976d2'),
                    size=16,
                    weight=ft.FontWeight.W_500
                ),
                description,
                ft.ResponsiveRow([
                    ft.Container(content=pressure_dropdown, col={'xs': 12, 'md': 6, 'lg': 3}),
                    ft.Container(content=flow_dropdown, col={'xs': 12, 'md': 6, 'lg': 3}),
                    ft.Container(content=speed_dropdown, col={'xs': 12, 'md': 6, 'lg': 3}),
                    ft.Container(content=power_dropdown, col={'xs': 12, 'md': 6, 'lg': 3}),
                ], spacing=10, run_spacing=10),
            ], spacing=10),
            padding=ft.padding.symmetric(vertical=10),
        )

    def _create_unit_dropdown(self, label: str, value: str, config_key: str, options: List[str]) -> ft.Dropdown:
        return ft.Dropdown(
            label=label,
            value=value,
            options=[ft.dropdown.Option(opt) for opt in options],
            on_change=lambda e, key=config_key: self._on_unit_changed(key, e.control.value),
            width=220,
        )

    def _on_unit_changed(self, config_key: str, value: str):
        if not self.parent_gui or not hasattr(self.parent_gui, "state_manager"):
            return

        self.parent_gui.state_manager.update_configuration(**{config_key: value})
        if hasattr(self.parent_gui, "_safe_page_update"):
            self.parent_gui._safe_page_update()

    def _update_test_lab_checkboxes(self, sap_code: str):
        """Update test lab checkboxes for a specific SAP code with voltage and notes"""
        if not self.parent_gui or sap_code not in self.comparison_test_lab_containers:
            return
            
        container = self.comparison_test_lab_containers[sap_code]
        
        # Get ONLY the tests that were selected in the previous step for this SAP code
        selected_tests_from_step2 = []
        if hasattr(self.parent_gui.state_manager.state, 'selected_tests'):
            selected_tests_from_step2 = [test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                        if test.sap_code == sap_code]
        
        # Get currently selected test labs for this SAP (from fine-grained selection)
        selected_test_labs = self.parent_gui.state_manager.state.selected_comparison_test_labs.get(sap_code, set())
        
        # Create checkboxes for each test lab
        test_lab_checkboxes = []
        if selected_tests_from_step2:
            test_lab_checkboxes.append(
                ft.Text(
                    f"Select test lab numbers for {sap_code} (from your previous selection):",
                    size=12,
                    color=self.theme_color('text_muted', 'grey'),
                    weight=ft.FontWeight.W_500
                )
            )
            
            # Add Select All / Select None buttons
            select_all_btn = ft.ElevatedButton(
                "Select All",
                icon=ft.Icons.SELECT_ALL,
                on_click=lambda e: self._select_all_test_labs(sap_code),
                scale=0.8,
                bgcolor=self.theme_color('primary_container', '#e3f2fd'),
                color=self.theme_color('on_primary_container', '#0d1a2b')
            )
            select_none_btn = ft.ElevatedButton(
                "Select None", 
                icon=ft.Icons.DESELECT,
                on_click=lambda e: self._select_none_test_labs(sap_code),
                scale=0.8,
                bgcolor=self.theme_color('surface_variant', '#f5f5f5'),
                color=self.theme_color('on_surface_variant', '#1f2933')
            )
            
            test_lab_checkboxes.append(
                ft.Row([select_all_btn, select_none_btn], spacing=5)
            )
            
            # Sort tests by test lab number for consistent display
            sorted_tests = sorted(selected_tests_from_step2, key=lambda t: t.test_lab_number)
            
            # Create enhanced checkboxes for each test with voltage and notes
            for test in sorted_tests:
                # Format voltage display
                voltage_display = f"{test.voltage}V" if test.voltage and test.voltage.strip() else "N/A"
                
                # Format notes display (truncate if too long)
                notes_display = test.notes if test.notes and test.notes.strip() else "No notes"
                if len(notes_display) > 50:
                    notes_display = notes_display[:47] + "..."
                
                # Create the checkbox with enhanced label (unchecked by default, but respects current state)
                is_test_lab_selected = test.test_lab_number in selected_test_labs
                checkbox = ft.Checkbox(
                    value=is_test_lab_selected,  # Respect current state but default to unchecked for new selections
                    on_change=lambda e, tl=test.test_lab_number, sap=sap_code: self._on_test_lab_checked(sap, tl, e.control.value),
                    scale=0.9
                )
                
                # Create a row with test info
                test_info_row = ft.Container(
                    content=ft.Row([
                        checkbox,
                        ft.Container(
                            content=ft.Column([
                                ft.Text(
                                    f"Test: {test.test_lab_number}",
                                    size=12,
                                    weight=ft.FontWeight.W_500,
                                    color=self.theme_color('primary', 'blue')
                                ),
                                ft.Row([
                                    ft.Text(
                                        f"Voltage: {voltage_display}",
                                        size=11,
                                        color=self.theme_color('success', 'darkgreen')
                                    ),
                                    ft.Text("•", size=11, color=self.theme_color('text_muted', 'grey')),
                                    ft.Text(
                                        f"Notes: {notes_display}",
                                        size=11,
                                        color=self.theme_color('text_muted', 'grey'), 
                                           tooltip=test.notes if test.notes else "No notes available")
                                ], spacing=5)
                            ], spacing=2),
                            expand=True
                        )
                    ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.START),
                    padding=ft.padding.symmetric(horizontal=5, vertical=3),
                    margin=ft.margin.only(bottom=3),
                    bgcolor=None,  # No highlighting by default since unchecked
                    border_radius=3,
                    border=None  # No border by default since unchecked
                )
                
                test_lab_checkboxes.append(test_info_row)
        else:
            test_lab_checkboxes.append(
                ft.Text(
                    "No test lab numbers were selected for this SAP code in the previous step",
                    size=12,
                    color=self.theme_color('warning', 'orange')
                )
            )
        
        # Update the container content
        container.content = ft.Column(test_lab_checkboxes, spacing=2)
        if self.parent_gui:
            self.parent_gui._safe_page_update()
    
    def _on_test_lab_checked(self, sap_code: str, test_lab: str, checked: bool):
        """Handle test lab checkbox changes"""
        if not self.parent_gui:
            return
            
        state = self.parent_gui.state_manager.state
        
        # Initialize the set if it doesn't exist
        if sap_code not in state.selected_comparison_test_labs:
            state.selected_comparison_test_labs[sap_code] = set()
        
        # Add or remove the test lab
        if checked:
            state.selected_comparison_test_labs[sap_code].add(test_lab)
        else:
            state.selected_comparison_test_labs[sap_code].discard(test_lab)
            
        self.parent_gui.state_manager.notify_observers("test_lab_selection_changed", {"sap": sap_code, "test_lab": test_lab, "selected": checked})
    
    def _select_all_test_labs(self, sap_code: str):
        """Select all test labs for a SAP code (only those selected in step 2)"""
        if not self.parent_gui:
            return
            
        # Get only the tests that were selected in step 2 for this SAP
        selected_tests_from_step2 = []
        if self.parent_gui.state_manager and self.parent_gui.state_manager.state.selected_tests:
            selected_tests_from_step2 = [test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                        if test.sap_code == sap_code]
        
        # Select all test lab numbers from step 2
        state = self.parent_gui.state_manager.state
        state.selected_comparison_test_labs[sap_code] = set(test.test_lab_number for test in selected_tests_from_step2)
        
        # Refresh the UI
        self._update_test_lab_checkboxes(sap_code)
        self.parent_gui.state_manager.notify_observers("test_lab_selection_changed", {"sap": sap_code, "action": "select_all"})
    
    def _select_none_test_labs(self, sap_code: str):
        """Deselect all test labs for a SAP code"""
        if not self.parent_gui:
            return
            
        # Clear selections
        state = self.parent_gui.state_manager.state
        state.selected_comparison_test_labs[sap_code] = set()
        
        # Refresh the UI
        self._update_test_lab_checkboxes(sap_code)
        self.parent_gui.state_manager.notify_observers("test_lab_selection_changed", {"sap": sap_code, "action": "select_none"})

    def _update_noise_test_checkboxes(self, sap_code: str):
        """Update noise test checkboxes for a specific SAP code - Load full data only when user selects"""
        if not self.parent_gui or sap_code not in self.noise_test_lab_containers:
            return

        color = self._color
            
        container = self.noise_test_lab_containers[sap_code]
        
        # Show loading indicator
        container.content = ft.Column([
            ft.Row([
                ft.ProgressRing(width=16, height=16, stroke_width=2),
                ft.Text(
                    "Loading noise tests from registry...",
                    size=12,
                    color=color('text_muted', 'grey')
                )
            ], spacing=10)
        ])
        
        # Update UI to show loading
        if (hasattr(self.parent_gui, '_safe_page_update') and 
            callable(getattr(self.parent_gui, '_safe_page_update', None))):
            self.parent_gui._safe_page_update()
        
        # Get currently selected noise tests for this SAP (from fine-grained selection)
        selected_noise_tests = set()
        if self._has_selected_noise_labs():
            selected_noise_tests = self.parent_gui.state_manager.state.selected_noise_test_labs.get(sap_code, set())
        
        # Load noise tests in background to avoid blocking UI
        def load_noise_tests_for_sap():
            try:
                # FIRST: Fast pre-check to see if SAP exists in registry
                noise_sap_codes = self._get_cached_noise_sap_codes()
                
                if sap_code not in noise_sap_codes:
                    # SAP code not found in noise registry
                    def update_no_tests():
                        container.content = ft.Column([
                            ft.Text(
                                f"SAP code {sap_code} not found in noise registry",
                                size=12,
                                color=color('warning', 'orange'),
                                weight=ft.FontWeight.W_500
                            ),
                            ft.Text(
                                "This SAP code is not available for noise testing",
                                size=11,
                                color=color('text_muted', 'grey'),
                                italic=True
                            )
                        ], spacing=5)
                        if self.parent_gui:
                            self.parent_gui._safe_page_update()
                    
                    update_no_tests()
                    return
                
                # SAP code exists, now load full test data
                logger.debug(f"ConfigTab: Loading full noise test data for SAP {sap_code}...")
                noise_tests = self._get_cached_noise_tests()
                
                # Filter tests for this specific SAP code
                sap_noise_tests = [test for test in noise_tests if test.sap_code == sap_code]
                
                def update_with_tests():
                    noise_test_checkboxes = []
                    
                    if sap_noise_tests:
                        # Show count and info
                        test_count = len(sap_noise_tests)
                        noise_test_checkboxes.append(
                            ft.Text(f"Found {test_count} noise test{'s' if test_count != 1 else ''} for {sap_code}", 
                                   size=14, color=color('success', 'darkgreen'), weight=ft.FontWeight.BOLD)
                        )
                        
                        # Add Select All / Select None buttons
                        buttons_row = ft.Row([
                            ft.ElevatedButton(
                                "Select All",
                                icon=ft.Icons.SELECT_ALL,
                                on_click=lambda e: self._select_all_noise_tests(sap_code),
                                scale=0.8,
                                bgcolor=color('success_container', '#e8f5e8')
                            ),
                            ft.ElevatedButton(
                                "Select None", 
                                icon=ft.Icons.DESELECT,
                                on_click=lambda e: self._select_none_noise_tests(sap_code),
                                scale=0.8,
                                bgcolor=color('surface_variant', '#f5f5f5')
                            )
                        ], spacing=5)
                        
                        noise_test_checkboxes.append(buttons_row)
                        
                        # Sort tests by test number for consistent display
                        sorted_tests = sorted(sap_noise_tests, key=lambda t: t.test_no if t.test_no else "")
                        
                        # Create checkboxes for each test with full details
                        for i, test in enumerate(sorted_tests):
                            # Prepare test display info
                            test_display = test.test_no if test.test_no and test.test_no.strip() else f"Test_{i+1}"
                            
                            # Discover what data files exist for this test (eager discovery!)
                            data_type, img_count, txt_count, icon, icon_color = self._discover_noise_data_for_test(
                                test.test_no, 
                                sap_code, 
                                test.date
                            )
                            
                            # Build data availability info
                            data_info_parts = []
                            if data_type == "txt_data":
                                data_info_parts.append(f"{txt_count} TXT file(s)")
                            elif data_type == "images":
                                data_info_parts.append(f"{img_count} image(s)")
                            elif data_type == "both":
                                data_info_parts.append(f"{img_count} image(s), {txt_count} TXT file(s)")
                            elif data_type == "none":
                                data_info_parts.append("No data files found")
                            else:
                                data_info_parts.append("Unknown")
                            
                            data_availability = " | ".join(data_info_parts)
                            
                            # Build detailed info string
                            details = []
                            if test.voltage and str(test.voltage).strip() and str(test.voltage) != "nan":
                                details.append(f"Voltage: {test.voltage}V")
                            if test.date and str(test.date).strip():
                                details.append(f"Date: {test.date}")
                            if test.client and str(test.client).strip():
                                details.append(f"Client: {test.client}")
                            if test.application and str(test.application).strip():
                                details.append(f"App: {test.application}")
                            
                            details_str = " | ".join(details) if details else "No additional details"
                            
                            # Create checkbox for this test
                            test_id = test.test_no if test.test_no and test.test_no.strip() else f"test_{i}"
                            is_test_selected = test_id in selected_noise_tests
                            
                            checkbox = ft.Checkbox(
                                value=is_test_selected,
                                on_change=lambda e, test_id=test_id, sap=sap_code: self._on_noise_test_checked(sap, test_id, e.control.value),
                                scale=0.9
                            )
                            
                            # Create detailed test info row with visual indicator
                            test_info_row = ft.Container(
                                content=ft.Row([
                                    checkbox,
                                    ft.Container(
                                        content=ft.Column([
                                            ft.Row([
                                                ft.Text(icon, size=16),  # Visual indicator icon
                                                ft.Text(
                                                    f"Test: {test_display}",
                                                    size=12,
                                                    weight=ft.FontWeight.W_500,
                                                    color=color('primary', 'blue')
                                                ),
                                            ], spacing=5),
                                            ft.Text(
                                                data_availability,
                                                size=11,
                                                color=icon_color,
                                                weight=ft.FontWeight.W_500
                                            ),  # Data availability
                                            ft.Text(
                                                details_str,
                                                size=11,
                                                color=color('text_muted', 'grey')
                                            ),
                                            ft.Text(
                                                f"Notes: {test.notes if test.notes and test.notes.strip() else 'No notes'}",
                                                size=10,
                                                color=color('text_muted', 'darkgrey'),
                                                italic=True
                                            )
                                        ], spacing=2),
                                        expand=True
                                    )
                                ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.START),
                                padding=ft.padding.symmetric(horizontal=5, vertical=5),
                                margin=ft.margin.only(bottom=5),
                                bgcolor=color('surface_variant', '#f8f9fa') if is_test_selected else None,
                                border_radius=5,
                                border=ft.border.all(1, color('primary', '#007bff')) if is_test_selected else ft.border.all(1, color('surface_variant', '#e9ecef'))
                            )
                            
                            noise_test_checkboxes.append(test_info_row)
                    else:
                        # No tests found for this SAP (should not happen if pre-check passed)
                        noise_test_checkboxes.append(
                            ft.Text(f"No noise tests found for {sap_code} in registry", 
                                   size=12, color=color('warning', 'orange'), italic=True)
                        )
                    
                    # Update container with test checkboxes
                    container.content = ft.Column(noise_test_checkboxes, spacing=3)
                    if self.parent_gui:
                        self.parent_gui._safe_page_update()
                
                update_with_tests()
                        
            except Exception as e:
                logger.error(f"Error loading noise tests for {sap_code}: {e}")
                def update_error():
                    container.content = ft.Column([
                        ft.Text(
                            f"Error loading noise tests: {str(e)}",
                            size=12,
                            color=color('error', 'red')
                        )
                    ])
                    if self.parent_gui:
                        self.parent_gui._safe_page_update()
                
                update_error()
        
        # Start loading in background thread
        run_in_background(load_noise_tests_for_sap)
    
    def _on_noise_test_checked(self, sap_code: str, test_no: str, checked: bool):
        """Handle noise test checkbox changes"""
        if not self.parent_gui:
            return
            
        state = self.parent_gui.state_manager.state
        
        # Initialize the set if it doesn't exist
        if sap_code not in state.selected_noise_test_labs:
            state.selected_noise_test_labs[sap_code] = set()
        
        # Add or remove the test
        if checked:
            state.selected_noise_test_labs[sap_code].add(test_no)
        else:
            state.selected_noise_test_labs[sap_code].discard(test_no)
            
        self.parent_gui.state_manager.notify_observers("noise_test_selection_changed", {"sap": sap_code, "test_no": test_no, "selected": checked})
    
    def _select_all_noise_tests(self, sap_code: str):
        """Select all noise tests for a SAP code"""
        if not self.parent_gui:
            return
            
        try:
            # Use cached noise tests for better performance
            noise_tests = self._get_cached_noise_tests()
            sap_noise_tests = [test for test in noise_tests if test.sap_code == sap_code]
            
            # Select all test numbers
            state = self.parent_gui.state_manager.state
            state.selected_noise_test_labs[sap_code] = set(test.test_no for test in sap_noise_tests if test.test_no)
            
            # Refresh the UI
            self._update_noise_test_checkboxes(sap_code)
            self.parent_gui.state_manager.notify_observers("noise_test_selection_changed", {"sap": sap_code, "action": "select_all"})
                
        except Exception as e:
            logger.error(f"Error selecting all noise tests for {sap_code}: {e}")
    
    def _select_none_noise_tests(self, sap_code: str):
        """Deselect all noise tests for a SAP code"""
        if not self.parent_gui:
            return
            
        # Clear selections
        state = self.parent_gui.state_manager.state
        state.selected_noise_test_labs[sap_code] = set()
        
        # Refresh the UI
        self._update_noise_test_checkboxes(sap_code)
        self.parent_gui.state_manager.notify_observers("noise_test_selection_changed", {"sap": sap_code, "action": "select_none"})

    def _disable_noise_feature(self):
        """Disable noise feature by clearing all noise SAP selections"""
        if self.parent_gui and hasattr(self.parent_gui, 'state_manager'):
            self.parent_gui.state_manager.state.selected_noise_saps.clear()
            self.parent_gui.state_manager.state.include_noise = False
            # Refresh the tab to update UI
            self.parent_gui._safe_page_update()

    def _disable_comparison_feature(self):
        """Disable comparison feature by clearing all comparison SAP selections"""
        if self.parent_gui and hasattr(self.parent_gui, 'state_manager'):
            self.parent_gui.state_manager.state.selected_comparison_saps.clear()
            self.parent_gui.state_manager.state.include_comparison = False
            # Refresh the tab to update UI
            self.parent_gui._safe_page_update()

    def _get_cached_noise_sap_codes(self):
        """
        Get SAP codes from noise registry cache using NoiseRegistryLoader service.
        
        This method delegates to the NoiseRegistryLoader service for centralized
        caching and loading logic. The service handles timeout protection, threading,
        and cache management automatically.
        
        Returns:
            List of SAP codes from noise registry, or empty list if unavailable
        """
        from ...config.directory_config import NOISE_REGISTRY_FILE, NOISE_TEST_DIR
        
        # Use the service to get cached SAP codes
        return self.noise_registry_loader.get_sap_codes(
            registry_path=str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
            noise_test_dir=str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None
        )
    
    def preload_noise_cache(self):
        """
        Preload noise SAP codes cache in the background for better performance.
        
        Delegates to NoiseRegistryLoader service for background preloading.
        """
        from ...config.directory_config import NOISE_REGISTRY_FILE, NOISE_TEST_DIR
        
        self.noise_registry_loader.preload_async(
            registry_path=str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
            noise_test_dir=str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None
        )
    
    def _on_noise_test_checked_with_feedback(self, sap_code: str, test_no: str, checked: bool):
        """Handle noise test checkbox changes with visual feedback"""
        if not self.parent_gui:
            return
        
        # Show brief visual feedback
        try:
            # Find the container for this SAP
            if sap_code in self.noise_test_lab_containers:
                container = self.noise_test_lab_containers[sap_code]
                # Briefly change background to indicate processing
                original_bgcolor = container.bgcolor
                container.bgcolor = "#f0f0f0"
                self.parent_gui._safe_page_update()
                
                # Call the original method
                self._on_noise_test_checked(sap_code, test_no, checked)
                
                # Restore original background without blocking - use thread pool
                import time
                def restore_bg():
                    time.sleep(0.1)  # 100ms delay
                    container.bgcolor = original_bgcolor
                    if self.parent_gui:
                        self.parent_gui._safe_page_update()
                
                # Use thread pool for delayed callback
                run_in_background(restore_bg)
            else:
                # Fallback to original method
                self._on_noise_test_checked(sap_code, test_no, checked)
                
        except Exception as e:
            logger.error(f"Error in noise test feedback: {e}")
            # Fallback to original method
            self._on_noise_test_checked(sap_code, test_no, checked)

    # ==================== Life Test (LF) Methods ====================
    
    def _on_lf_sap_checkbox_changed(self, e, sap_code: str):
        """Handle LF SAP checkbox changes"""
        logger.info(f"🔬 LF SAP checkbox changed: {sap_code} -> {e.control.value if hasattr(e.control, 'value') else 'NO VALUE'}")
        
        if not self.parent_gui or not hasattr(e.control, 'value'):
            logger.warning("⚠️ No parent GUI or control value")
            return
        
        checked = e.control.value
        state = self.parent_gui.state_manager.state
        
        if checked:
            state.selected_lf_saps.add(sap_code)
            logger.info(f"✅ Added SAP {sap_code} to selected LF SAPs")
            # Load LF tests for this SAP
            self._update_lf_test_checkboxes(sap_code)
        else:
            state.selected_lf_saps.discard(sap_code)
            logger.info(f"❌ Removed SAP {sap_code} from selected LF SAPs")
            # Clear selected tests for this SAP
            if sap_code in state.selected_lf_test_numbers:
                del state.selected_lf_test_numbers[sap_code]
        
        # Update container visibility
        if sap_code in self.lf_test_containers:
            self.lf_test_containers[sap_code].visible = checked
            self.parent_gui._safe_page_update()
        else:
            logger.warning(f"⚠️ LF container not found for {sap_code}")
    
    def _update_lf_test_checkboxes(self, sap_code: str):
        """Load and display LF tests for a SAP code"""
        logger.info(f"🔄 Loading LF tests for SAP: {sap_code}")
        
        if not self.parent_gui or sap_code not in self.lf_test_containers:
            logger.warning(f"⚠️ Cannot load LF tests - parent_gui={bool(self.parent_gui)}, container_exists={sap_code in self.lf_test_containers if hasattr(self, 'lf_test_containers') else False}")
            return
        
        color = self._color

        try:
            # Get LF registry reader
            from ...services.lf_registry_reader import LifeTestRegistryReader
            lf_reader = LifeTestRegistryReader()
            logger.info(f"✅ Created LF registry reader")
            
            # Get tests for this SAP
            lf_tests = lf_reader.get_tests_for_sap(sap_code)
            logger.info(f"📊 Found {len(lf_tests)} LF test(s) for SAP {sap_code}")
            
            if not lf_tests:
                # Show message if no tests found
                self.lf_test_containers[sap_code].content = ft.Column([
                    ft.Text(
                        "No LF tests found for this SAP code",
                        color=color('text_muted', 'grey'),
                        italic=True,
                        size=12
                    )
                ])
                self.parent_gui._safe_page_update()
                return
            
            # Get current selections
            state = self.parent_gui.state_manager.state
            selected_tests = state.selected_lf_test_numbers.get(sap_code, set())
            
            # Create checkboxes for each test
            test_checkboxes = []
            for lf_test in lf_tests:
                # Format label with test number, notes, and file existence
                file_status = "✅" if lf_test.file_exists else "❌"
                notes_preview = lf_test.notes[:50] + "..." if lf_test.notes and len(lf_test.notes) > 50 else (lf_test.notes or "No notes")
                label = f"{file_status} {lf_test.test_number} - {notes_preview}"
                
                checkbox = ft.Checkbox(
                    label=label,
                    value=(lf_test.test_number in selected_tests),
                    on_change=lambda e, test_num=lf_test.test_number, sap=sap_code: self._on_lf_test_checked(sap, test_num, e.control.value),
                    tooltip=f"Notes: {lf_test.notes or 'No notes'}\nFile: {'Found' if lf_test.file_exists else 'Not found'}",
                    disabled=(not lf_test.file_exists)  # Disable if file doesn't exist
                )
                test_checkboxes.append(checkbox)
            
            # Update container with checkboxes
            self.lf_test_containers[sap_code].content = ft.Column([
                ft.Row([
                    ft.Text(f"Found {len(lf_tests)} LF test(s)", size=12, weight=ft.FontWeight.BOLD),
                    ft.TextButton("Select All", on_click=lambda e, s=sap_code: self._select_all_lf_tests(s)),
                    ft.TextButton("Deselect All", on_click=lambda e, s=sap_code: self._deselect_all_lf_tests(s))
                ]),
                *test_checkboxes
            ], spacing=2)
            
            self.parent_gui._safe_page_update()
            logger.info(f"✅ Loaded {len(lf_tests)} LF tests for SAP {sap_code}")
            
        except Exception as e:
            logger.error(f"Error loading LF tests for {sap_code}: {e}")
            self.lf_test_containers[sap_code].content = ft.Column([
                ft.Text(
                    f"Error loading LF tests: {str(e)}",
                    color=color('error', 'red'),
                    size=12
                )
            ])
            self.parent_gui._safe_page_update()
    
    def _on_lf_test_checked(self, sap_code: str, test_number: str, checked: bool):
        """Handle LF test checkbox changes"""
        if not self.parent_gui:
            return
        
        state = self.parent_gui.state_manager.state
        
        # Initialize the set if it doesn't exist
        if sap_code not in state.selected_lf_test_numbers:
            state.selected_lf_test_numbers[sap_code] = set()
        
        # Add or remove the test
        if checked:
            state.selected_lf_test_numbers[sap_code].add(test_number)
            logger.debug(f"Selected LF test {test_number} for SAP {sap_code}")
        else:
            state.selected_lf_test_numbers[sap_code].discard(test_number)
            logger.debug(f"Deselected LF test {test_number} for SAP {sap_code}")
    
    def _select_all_lf_tests(self, sap_code: str):
        """Select all available LF tests for a SAP code"""
        if not self.parent_gui or sap_code not in self.lf_test_containers:
            return
        
        try:
            from ...services.lf_registry_reader import LifeTestRegistryReader
            lf_reader = LifeTestRegistryReader()
            lf_tests = lf_reader.get_tests_for_sap(sap_code)
            
            # Select only tests with existing files
            state = self.parent_gui.state_manager.state
            state.selected_lf_test_numbers[sap_code] = set(
                test.test_number for test in lf_tests if test.file_exists
            )
            
            # Refresh UI
            self._update_lf_test_checkboxes(sap_code)
            logger.info(f"Selected all LF tests for SAP {sap_code}")
            
        except Exception as e:
            logger.error(f"Error selecting all LF tests: {e}")
    
    def _deselect_all_lf_tests(self, sap_code: str):
        """Deselect all LF tests for a SAP code"""
        if not self.parent_gui:
            return
        
        state = self.parent_gui.state_manager.state
        if sap_code in state.selected_lf_test_numbers:
            state.selected_lf_test_numbers[sap_code].clear()
        
        # Refresh UI
        self._update_lf_test_checkboxes(sap_code)
        logger.info(f"Deselected all LF tests for SAP {sap_code}")

    def validate_config_selections(self):
        """Validate all configurations before proceeding to generate step"""
        validation_errors = []
        warnings = []
        
        if not self.parent_gui:
            return False, ["GUI not available"], []
        
        state = self.parent_gui.state_manager.state
        
        # 1. Check if any features are enabled
        has_performance = bool(state.selected_tests)
        has_comparison = bool(state.selected_comparison_saps)
        has_noise = bool(state.selected_noise_saps)
        
        if not has_performance:
            validation_errors.append("No performance tests available (this should not happen at this step)")
        
        # 2. Validate comparison selections
        if has_comparison:
            for sap in state.selected_comparison_saps:
                selected_test_labs = state.selected_comparison_test_labs.get(sap, set())
                if not selected_test_labs:
                    validation_errors.append(f"SAP {sap} is selected for comparison but no test labs are chosen")
                else:
                    # Verify the test labs exist in our selected tests
                    available_tests = [test for test in state.selected_tests.values() 
                                     if test.sap_code == sap]
                    available_test_labs = set(test.test_lab_number for test in available_tests)
                    invalid_labs = selected_test_labs - available_test_labs
                    if invalid_labs:
                        validation_errors.append(f"SAP {sap}: Test labs {invalid_labs} not found in selected tests")
        
        # 3. Validate noise selections
        if has_noise:
            try:
                all_noise_tests = self._get_cached_noise_tests()
                
                for sap in state.selected_noise_saps:
                    selected_noise_tests = state.selected_noise_test_labs.get(sap, set())
                    if not selected_noise_tests:
                        validation_errors.append(f"SAP {sap} is selected for noise but no tests are chosen")
                    else:
                        # Verify the noise tests exist in registry
                        available_noise_tests = [test for test in all_noise_tests if test.sap_code == sap]
                        available_test_nos = set(test.test_no for test in available_noise_tests if test.test_no)
                        invalid_tests = selected_noise_tests - available_test_nos
                        if invalid_tests:
                            validation_errors.append(f"SAP {sap}: Noise tests {invalid_tests} not found in registry")
            except Exception as e:
                validation_errors.append(f"Could not validate noise tests: {str(e)}")
        
        # 4. Generate warnings for empty features
        if not has_comparison and not has_noise:
            warnings.append("Only performance reports will be generated (no comparison or noise selected)")
        elif not has_comparison:
            warnings.append("No comparison reports will be generated")
        elif not has_noise:
            warnings.append("No noise reports will be generated")
        
        # 5. Summary information
        if not validation_errors:
            total_performance = len(state.selected_tests)
            total_comparison = sum(len(state.selected_comparison_test_labs.get(sap, set())) 
                                 for sap in state.selected_comparison_saps)
            total_noise = sum(len(state.selected_noise_test_labs.get(sap, set())) 
                            for sap in state.selected_noise_saps)
            
            logger.debug(f"Configuration validation passed:")
            logger.debug(f"  Performance tests: {total_performance}")
            logger.debug(f"  Comparison tests: {total_comparison} from {len(state.selected_comparison_saps)} SAPs")
            logger.debug(f"  Noise tests: {total_noise} from {len(state.selected_noise_saps)} SAPs")
        
        return len(validation_errors) == 0, validation_errors, warnings

    def prepare_final_test_data(self):
        """Prepare and validate all test data for the final report generation"""
        if not self.parent_gui:
            return None
        
        state = self.parent_gui.state_manager.state
        
        # Prepare performance test data (always included)
        performance_data = {
            'total_tests': len(state.selected_tests),
            'tests_by_sap': {},
            'test_details': list(state.selected_tests.values())
        }
        
        # Group performance tests by SAP
        for test in state.selected_tests.values():
            sap = test.sap_code
            if sap not in performance_data['tests_by_sap']:
                performance_data['tests_by_sap'][sap] = []
            performance_data['tests_by_sap'][sap].append(test)
        
        # Prepare comparison test data
        comparison_data = {
            'enabled': bool(state.selected_comparison_saps),
            'saps': list(state.selected_comparison_saps),
            'total_tests': 0,
            'tests_by_sap': {}
        }
        
        for sap in state.selected_comparison_saps:
            selected_test_labs = state.selected_comparison_test_labs.get(sap, set())
            if selected_test_labs:
                # Get the actual test objects for these test labs
                sap_tests = [test for test in state.selected_tests.values() 
                           if test.sap_code == sap and test.test_lab_number in selected_test_labs]
                comparison_data['tests_by_sap'][sap] = sap_tests
                comparison_data['total_tests'] += len(sap_tests)
        
        # Prepare noise test data
        noise_data = {
            'enabled': bool(state.selected_noise_saps),
            'saps': list(state.selected_noise_saps),
            'total_tests': 0,
            'tests_by_sap': {}
        }
        
        if state.selected_noise_saps:
            try:
                all_noise_tests = self._get_cached_noise_tests()
                
                for sap in state.selected_noise_saps:
                    selected_noise_tests = state.selected_noise_test_labs.get(sap, set())
                    if selected_noise_tests:
                        # Get the actual noise test objects
                        sap_noise_tests = [test for test in all_noise_tests 
                                         if test.sap_code == sap and test.test_no in selected_noise_tests]
                        noise_data['tests_by_sap'][sap] = sap_noise_tests
                        noise_data['total_tests'] += len(sap_noise_tests)
            except Exception as e:
                logger.error(f"Error preparing noise data: {e}")
                noise_data['error'] = str(e)
        
        # Create final summary
        final_data = {
            'performance': performance_data,
            'comparison': comparison_data,
            'noise': noise_data,
            'summary': {
                'total_performance_tests': performance_data['total_tests'],
                'total_comparison_tests': comparison_data['total_tests'],
                'total_noise_tests': noise_data['total_tests'],
                'total_saps': len(set(performance_data['tests_by_sap'].keys())),
                'comparison_saps': len(comparison_data['saps']),
                'noise_saps': len(noise_data['saps'])
            }
        }
        
        # Store this in the state for the generate tab
        state.final_test_data = final_data
        
        return final_data

    def _generate_comparison_sections_fast(self, found_sap_codes):
        """Generate comparison sections quickly without heavy computation"""
        comparison_sap_sections = []
        
        if not self.parent_gui:
            return comparison_sap_sections
            
        for sap in found_sap_codes:
            # Get selected tests from step 2 for this SAP to show in label  
            selected_tests_count = 0
            if self._has_selected_tests():
                selected_tests_count = len([test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                          if test.sap_code == sap])
            
            # Only show SAP codes that have selected tests from step 2
            if selected_tests_count == 0:
                continue
                
            # Create the main SAP checkbox with test count from step 2
            sap_label = f"{sap} ({selected_tests_count} test{'s' if selected_tests_count != 1 else ''} from previous selection)"
            sap_checkbox = ft.Checkbox(
                label=sap_label, 
                value=(sap in self.parent_gui.state_manager.state.selected_comparison_saps if hasattr(self.parent_gui.state_manager, 'state') else False), 
                on_change=self.parent_gui.event_handlers.on_sap_checked("comparison") if hasattr(self.parent_gui, 'event_handlers') else None,
                data=sap
            )
            
            # Container for test lab checkboxes (initially hidden)
            if sap not in self.comparison_test_lab_containers:
                test_lab_container = ft.Container(
                    content=ft.Column([]),
                    visible=(sap in self.parent_gui.state_manager.state.selected_comparison_saps if hasattr(self.parent_gui.state_manager, 'state') else False),
                    margin=ft.margin.only(left=30, top=5, bottom=10),
                    bgcolor=self.theme_color('surface_variant', '#f8f9fa'),
                    padding=ft.padding.all(10),
                    border_radius=5,
                    border=ft.border.all(1, self.theme_color('outline', '#e9ecef'))
                )
                self.comparison_test_lab_containers[sap] = test_lab_container
            else:
                test_lab_container = self.comparison_test_lab_containers[sap]
                # Update visibility based on current selection state
                test_lab_container.visible = (sap in self.parent_gui.state_manager.state.selected_comparison_saps if hasattr(self.parent_gui.state_manager, 'state') else False)
            
            # Create a section with the SAP checkbox and its test lab container
            sap_section = ft.Column([
                sap_checkbox,
                self.comparison_test_lab_containers[sap]
            ], spacing=2)
            
            comparison_sap_sections.append(sap_section)
            
            # Load test lab checkboxes lazily when SAP is selected
            if hasattr(self.parent_gui.state_manager, 'state') and sap in self.parent_gui.state_manager.state.selected_comparison_saps:
                # Load in background to avoid blocking UI
                run_in_background(lambda: self._update_test_lab_checkboxes(sap))
                
        return comparison_sap_sections
        
    def _generate_noise_sections_fast(self, found_sap_codes):
        """Generate noise sections with proper filtering"""
        noise_sap_sections = []
        
        if not self.parent_gui:
            return noise_sap_sections
        
        # Force load noise SAP codes for proper filtering
        logger.debug("ConfigTab: Loading noise SAP codes for filtering...")
        noise_sap_codes = self._get_cached_noise_sap_codes()
        
        # If cache is still not loaded, wait a bit for it to load
        if len(noise_sap_codes) == 0:
            import time
            logger.debug("ConfigTab: Cache not ready, waiting for background load...")
            # Wait up to 5 seconds for cache to load using polling with small intervals
            wait_count = 0
            while len(noise_sap_codes) == 0 and wait_count < 50:
                time.sleep(0.1)  # Non-blocking wait
                noise_sap_codes = self._get_cached_noise_sap_codes()
                wait_count += 1
            logger.debug(f"ConfigTab: After waiting, got {len(noise_sap_codes)} noise SAP codes")
        
        logger.debug(f"ConfigTab: Using {len(noise_sap_codes)} noise SAP codes for filtering")
        logger.debug(f"ConfigTab: First 10 noise SAP codes: {noise_sap_codes[:10]}")
            
        for sap in found_sap_codes:
            # Get selected tests count from step 2 (fast operation)
            selected_tests_count = 0
            if self._has_selected_tests():
                selected_tests_count = len([test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                          if test.sap_code == sap])
            
            # Only show SAP codes that have selected tests from step 2 AND are in noise registry
            if selected_tests_count == 0:
                logger.debug(f"ConfigTab: Skipping {sap} - no performance tests")
                continue
                
            if sap not in noise_sap_codes:
                logger.debug(f"ConfigTab: Skipping {sap} - not in noise registry")
                continue
                
            logger.debug(f"ConfigTab: Including {sap} - {selected_tests_count} performance tests, in noise registry")
                
            # Create the main SAP checkbox with proper label
            sap_label = f"{sap} ({selected_tests_count} performance, ✓ available in noise registry)"
            sap_checkbox = ft.Checkbox(
                label=sap_label,
                value=(sap in self.parent_gui.state_manager.state.selected_noise_saps if hasattr(self.parent_gui.state_manager, 'state') else False),
                on_change=self.parent_gui.event_handlers.on_sap_checked("noise") if hasattr(self.parent_gui, 'event_handlers') else None,
                data=sap
            )
            
            # Container for noise test selection (initially hidden)
            if sap not in self.noise_test_lab_containers:
                noise_test_container = ft.Container(
                    content=ft.Column([]),
                    visible=(sap in self.parent_gui.state_manager.state.selected_noise_saps if hasattr(self.parent_gui.state_manager, 'state') else False),
                    margin=ft.margin.only(left=30, top=5, bottom=10),
                    bgcolor=self.theme_color('success_container', '#e8f5e8'),  # Light green background for noise
                    padding=ft.padding.all(10),
                    border_radius=5,
                    border=ft.border.all(1, self.theme_color('success', '#2e7d32'))
                )
                self.noise_test_lab_containers[sap] = noise_test_container
            else:
                noise_test_container = self.noise_test_lab_containers[sap]
                # Update visibility based on current selection state
                noise_test_container.visible = (sap in self.parent_gui.state_manager.state.selected_noise_saps if hasattr(self.parent_gui.state_manager, 'state') else False)
            
            # Create a section with the SAP checkbox and its noise test container
            noise_sap_section = ft.Column([
                sap_checkbox,
                self.noise_test_lab_containers[sap]
            ], spacing=2)
            
            noise_sap_sections.append(noise_sap_section)
            
            # Load noise test checkboxes lazily if SAP is selected
            if hasattr(self.parent_gui.state_manager, 'state') and sap in self.parent_gui.state_manager.state.selected_noise_saps:
                run_in_background(lambda: self._update_noise_test_checkboxes(sap))
                
        logger.debug(f"ConfigTab: Generated {len(noise_sap_sections)} noise sections")
        return noise_sap_sections
    
    def _generate_lf_sections_fast(self, found_sap_codes):
        """Generate Life Test (LF) sections with proper filtering"""
        lf_sap_sections = []
        
        logger.info("=" * 60)
        logger.info("🔬 GENERATING LF SECTIONS")
        logger.info("=" * 60)
        logger.info(f"📋 Found SAP codes: {found_sap_codes}")
        
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            logger.warning("⚠️ No parent GUI or state manager")
            return lf_sap_sections
        
        # Initialize LF test containers dict if not exists
        if not hasattr(self, 'lf_test_containers'):
            self.lf_test_containers = {}
            logger.info("✅ Initialized lf_test_containers")
        
        logger.info(f"🔄 Processing {len(found_sap_codes)} SAP codes for LF sections...")
        
        for sap in found_sap_codes:
            logger.debug(f"  📌 Creating LF section for SAP: {sap}")
            # Create checkbox for this SAP
            sap_checkbox = ft.Checkbox(
                label=f"SAP: {sap}",
                value=False,
                on_change=lambda e, s=sap: self._on_lf_sap_checkbox_changed(e, s)
            )
            
            # Create container for LF tests (will be populated when checkbox is checked)
            lf_test_container = ft.Container(
                content=ft.Column([], spacing=2),
                padding=ft.padding.only(left=30),
                visible=False
            )
            self.lf_test_containers[sap] = lf_test_container
            
            # Restore state if SAP was previously selected
            if hasattr(self.parent_gui.state_manager, 'state'):
                if sap in self.parent_gui.state_manager.state.selected_lf_saps:
                    sap_checkbox.value = True
                    lf_test_container.visible = True
            
            # Create section with SAP checkbox and LF test container
            lf_sap_section = ft.Column([
                sap_checkbox,
                lf_test_container
            ], spacing=2)
            
            lf_sap_sections.append(lf_sap_section)
            
            # Load LF tests lazily if SAP is selected
            if hasattr(self.parent_gui.state_manager, 'state') and sap in self.parent_gui.state_manager.state.selected_lf_saps:
                run_in_background(lambda s=sap: self._update_lf_test_checkboxes(s))
        
        logger.info(f"✅ Generated {len(lf_sap_sections)} LF section(s)")
        logger.info("=" * 60)
        return lf_sap_sections
        
    def _preload_noise_registry_async(self):
        """
        Preload noise SAP codes for fast pre-check - Works even without parent_gui.
        
        Delegates to NoiseRegistryLoader service for background preloading.
        Only starts preload if not already loading.
        """
        from ...config.directory_config import NOISE_REGISTRY_FILE, NOISE_TEST_DIR
        
        # Only start background loading if not already in progress
        if not self.noise_registry_loader.is_loading():
            self.noise_registry_loader.preload_async(
                registry_path=str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
                noise_test_dir=str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None
            )
    
    def _update_noise_label_async(self, sap_code, checkbox):
        """Update noise section label and hide if not in noise registry"""
        try:
            # Safety check
            if not self.parent_gui:
                return
                
            # FAST PRE-CHECK: Only check if SAP code exists in noise registry
            noise_sap_codes = self._get_cached_noise_sap_codes()
            
            # Get performance test count (fast operation)
            performance_count = 0
            if (hasattr(self.parent_gui, 'state_manager') and 
                hasattr(self.parent_gui.state_manager, 'state') and 
                hasattr(self.parent_gui.state_manager.state, 'selected_tests')):
                performance_count = sum(1 for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                      if test.sap_code == sap_code)
            
            # Update label and visibility based on SAP code presence in noise registry
            if sap_code in noise_sap_codes:
                new_label = f"{sap_code} ({performance_count} performance, ✓ available in noise registry)"
                checkbox.label = new_label
                checkbox.visible = True
            else:
                # SAP not in noise registry - hide this checkbox entirely
                checkbox.visible = False
                # Also hide the parent container if it exists
                if hasattr(checkbox, 'parent') and checkbox.parent:
                    checkbox.parent.visible = False
            
            # Update the UI
            if self.parent_gui and hasattr(self.parent_gui, '_safe_page_update'):
                self.parent_gui._safe_page_update()
                
        except Exception as e:
            logger.error(f"Error updating noise label for {sap_code}: {e}")

    def _load_more_noise_tests(self, sap_code: str, current_limit: int):
        """Load more noise tests for a SAP code when user clicks 'Load More'"""
        if not self.parent_gui or sap_code not in self.noise_test_lab_containers:
            return
            
        container = self.noise_test_lab_containers[sap_code]
        color = self._color
        
        # Show loading briefly
        temp_content = container.content
        container.content = ft.Column([
            ft.Row([
                ft.ProgressRing(width=16, height=16, stroke_width=2),
                ft.Text(
                    "Loading more tests...",
                    size=12,
                    color=color('text_muted', 'grey')
                )
            ], spacing=10)
        ])
        if self.parent_gui and hasattr(self.parent_gui, '_safe_page_update'):
            self.parent_gui._safe_page_update()
        
        try:
            # Get all noise tests for this SAP
            noise_tests = self._get_cached_noise_tests()
            sap_noise_tests = [test for test in noise_tests if test.sap_code == sap_code]
            sorted_tests = sorted(sap_noise_tests, key=lambda t: t.test_no if t.test_no else "")
            
            # Get selected tests
            selected_noise_tests = set()
            if (hasattr(self.parent_gui, 'state_manager') and 
                hasattr(self.parent_gui.state_manager, 'state') and
                hasattr(self.parent_gui.state_manager.state, 'selected_noise_test_labs')):
                selected_noise_tests = self.parent_gui.state_manager.state.selected_noise_test_labs.get(sap_code, set())
            
            # Show all tests now (remove the limit)
            all_checkboxes = []
            
            # Add header
            all_checkboxes.append(
                ft.Text(
                    f"All {len(sap_noise_tests)} noise tests for {sap_code}:",
                    size=14,
                    color=color('success', 'darkgreen'),
                    weight=ft.FontWeight.BOLD
                )
            )
            
            # Add control buttons
            select_all_btn = ft.ElevatedButton(
                "Select All",
                icon=ft.Icons.SELECT_ALL,
                on_click=lambda e: self._select_all_noise_tests(sap_code),
                scale=0.8,
                bgcolor=color('success_container', '#e8f5e8')
            )
            select_none_btn = ft.ElevatedButton(
                "Select None", 
                icon=ft.Icons.DESELECT,
                on_click=lambda e: self._select_none_noise_tests(sap_code),
                scale=0.8,
                bgcolor=color('surface_variant', '#f5f5f5')
            )
            
            all_checkboxes.append(
                ft.Row([select_all_btn, select_none_btn], spacing=5)
            )
            
            # Add all test checkboxes
            for i, test in enumerate(sorted_tests):
                test_display = test.test_no if test.test_no and test.test_no.strip() else f"Test_{i+1}"
                
                # Essential info only
                essential_info = ""
                if test.voltage and str(test.voltage).strip() and str(test.voltage) != "nan":
                    essential_info += f" ({test.voltage}V)"
                if test.date and str(test.date).strip():
                    essential_info += f" [{test.date}]"
                
                test_id = test.test_no if test.test_no and test.test_no.strip() else f"test_{i}"
                is_test_selected = test_id in selected_noise_tests
                
                checkbox = ft.Checkbox(
                    label=f"{test_display}{essential_info}",
                    value=is_test_selected,
                    on_change=lambda e, test_id=test_id, sap=sap_code: self._on_noise_test_checked_with_feedback(sap, test_id, e.control.value),
                    scale=0.9
                )
                
                all_checkboxes.append(checkbox)
            
            # Update container with all tests
            container.content = ft.Column(all_checkboxes, spacing=3, scroll=ft.ScrollMode.AUTO, height=400)
            if self.parent_gui and hasattr(self.parent_gui, '_safe_page_update'):
                self.parent_gui._safe_page_update()
                
        except Exception as e:
            logger.error(f"Error loading more noise tests: {e}")
            # Restore previous content
            container.content = temp_content
            if self.parent_gui and hasattr(self.parent_gui, '_safe_page_update'):
                self.parent_gui._safe_page_update()
    
    def _get_cached_noise_tests(self):
        """
        Get full noise test validation info for detailed operations.
        This is heavier than SAP code pre-check and should only be used when needed.
        """
        # If we have cached tests, return them
        if self._cached_noise_tests is not None:
            return self._cached_noise_tests
            
        # Load full test data when needed
        import time
        
        try:
            from ...validators.noise_test_validator import NoiseTestValidator
            from ...config.directory_config import NOISE_REGISTRY_FILE, NOISE_TEST_DIR
            
            if NOISE_REGISTRY_FILE and NOISE_TEST_DIR:
                sheet_name = "Registro"
                validator = NoiseTestValidator(str(NOISE_TEST_DIR), sheet_name)
                
                logger.debug("ConfigTab: Loading full noise test data for detailed UI...")
                start_time = time.time()
                
                # Load full validation data (heavier operation)
                self._cached_noise_tests = validator.validate_from_registry(str(NOISE_REGISTRY_FILE), max_rows=2000)
                
                load_time = time.time() - start_time
                logger.debug(f"ConfigTab: Loaded {len(self._cached_noise_tests)} full noise tests in {load_time:.2f}s")
                
                return self._cached_noise_tests
            else:
                logger.debug("ConfigTab: No noise registry configured for full test loading")
                self._cached_noise_tests = []
                return []
                
        except Exception as e:
            logger.error(f"Error loading full noise tests: {e}")
            self._cached_noise_tests = []
            return []

    def _has_state_manager(self):
        """Helper to check if parent GUI has a valid state manager"""
        return (self.parent_gui and 
                hasattr(self.parent_gui, 'state_manager') and 
                self.parent_gui.state_manager and
                hasattr(self.parent_gui.state_manager, 'state'))

    def _has_selected_tests(self):
        """Helper to check if state has selected tests"""
        try:
            state_manager = getattr(self.parent_gui, 'state_manager', None)
            if not state_manager:
                return False
            state = getattr(state_manager, 'state', None)
            if not state:
                return False
            return hasattr(state, 'selected_tests')
        except (AttributeError, TypeError):
            return False
    
    def _has_selected_noise_labs(self):
        """Helper to check if state has selected noise test labs"""
        try:
            state_manager = getattr(self.parent_gui, 'state_manager', None)
            if not state_manager:
                return False
            state = getattr(state_manager, 'state', None)
            if not state:
                return False
            return hasattr(state, 'selected_noise_test_labs')
        except (AttributeError, TypeError):
            return False
    
    def _discover_noise_data_for_test(self, test_no: str, sap_code: str, date_str: str) -> tuple:
        """
        Discover what data files exist for a noise test.
        Returns: (data_type, image_count, txt_count, icon, color)
        """
        color = self._color
        try:
            # Extract year from date
            import pandas as pd
            year = None
            if date_str and str(date_str).strip():
                try:
                    parsed_date = pd.to_datetime(date_str, errors='coerce')
                    if pd.notna(parsed_date):
                        year = str(parsed_date.year)
                except Exception as e:
                    logger.debug(f"Error parsing date {date_str}: {e}")
                    pass
            
            if not year:
                logger.debug(f"No year extracted from date {date_str} for test {test_no}")
                return ("unknown", 0, 0, "❓", color('text_muted', 'grey'))
            
            # Get noise handler - prefer existing one from app
            noise_handler = None
            if hasattr(self.parent_gui, 'app') and self.parent_gui.app:
                if hasattr(self.parent_gui.app, 'noise_handler'):
                    noise_handler = self.parent_gui.app.noise_handler
                    logger.debug(f"Using existing noise_handler for test {test_no}")
            
            if not noise_handler:
                # Need to create temporary handler - check if noise folder is available
                if not hasattr(self.parent_gui, 'state_manager') or not self.parent_gui.state_manager:
                    logger.debug(f"No state_manager available for test {test_no}")
                    return ("unknown", 0, 0, "❓", color('text_muted', 'grey'))
                
                state = self.parent_gui.state_manager.state
                noise_folder = getattr(state, 'selected_noise_folder', None)
                
                if not noise_folder:
                    logger.debug(f"No noise folder selected for test {test_no}")
                    return ("unknown", 0, 0, "❓", color('text_muted', 'grey'))
                
                # Create temporary handler with registry if available
                from ...simplified_noise_handler import SimplifiedNoiseDataHandler
                from ...config.app_config import AppConfig
                config = AppConfig(noise_dir=noise_folder)
                
                # Try to get registry from state if available
                registry_df = getattr(state, 'noise_registry_df', None)
                if registry_df is not None:
                    noise_handler = SimplifiedNoiseDataHandler(config, registry_df=registry_df)
                    logger.debug(f"Created temporary noise_handler with registry ({len(registry_df)} records) for test {test_no}")
                else:
                    noise_handler = SimplifiedNoiseDataHandler(config)
                    logger.debug(f"Created temporary noise_handler without registry for test {test_no}")
            
            # Discover files
            noise_info = noise_handler.get_noise_test_info_by_test_year(
                test_number=test_no,
                year=year,
                sap_code=sap_code
            )
            logger.debug(f"Noise info for test {test_no}: {noise_info.data_type if noise_info else 'None'}")
            
            if noise_info:
                data_type = noise_info.data_type
                image_count = len(noise_info.image_paths) if noise_info.image_paths else 0
                txt_count = len(noise_info.txt_files) if noise_info.txt_files else 0
                
                # Determine icon and color based on data type
                if data_type == "txt_data":
                    return (data_type, image_count, txt_count, "📊", color('success', 'green'))
                elif data_type == "images":
                    return (data_type, image_count, txt_count, "🖼️", color('success', 'green'))
                elif txt_count > 0 and image_count > 0:
                    return ("both", image_count, txt_count, "📊🖼️", color('primary', 'blue'))
                else:
                    return (data_type, image_count, txt_count, "⚠️", color('warning', 'orange'))

            return ("none", 0, 0, "⚠️", color('warning', 'orange'))

        except Exception as e:
            logger.debug(f"Could not discover data for test {test_no}: {e}")
            return ("unknown", 0, 0, "❓", color('text_muted', 'grey'))
    
    def _add_new_comparison_group(self, e):
        """Add a new comparison group"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager') or not self.comparison_groups_container:
            return
        
        # Get available SAP codes with selected tests
        found_sap_codes = []
        if hasattr(self.parent_gui.state_manager, 'state') and hasattr(self.parent_gui.state_manager.state, 'found_sap_codes'):
            found_sap_codes = self.parent_gui.state_manager.state.found_sap_codes.copy()
        
        # Filter to only SAP codes that have selected tests
        available_sap_codes = []
        if self._has_selected_tests():
            for sap in found_sap_codes:
                selected_tests_count = len([test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                          if test.sap_code == sap])
                if selected_tests_count > 0:
                    available_sap_codes.append(sap)
        
        if not available_sap_codes:
            # Show error - no SAP codes available
            return
        
        # Increment counter and create new comparison group
        self.comparison_counter += 1
        group_id = f"comparison_{self.comparison_counter}"
        
        # Store the group info FIRST (with empty containers)
        group_info = {
            'id': group_id,
            'number': self.comparison_counter,
            'container': None,  # Will be set after creation
            'sap_containers': {}
        }
        self.comparison_groups.append(group_info)
        
        # Create the comparison group container (now the group exists in the list)
        comparison_group = self._create_comparison_group(group_id, self.comparison_counter, available_sap_codes)
        
        # Update the container reference
        group_info['container'] = comparison_group
        
        # Add to the comparison groups container
        if (hasattr(self.comparison_groups_container, 'current') and 
            self.comparison_groups_container.current and
            hasattr(self.comparison_groups_container.current, 'content')):
            
            # Rebuild the UI with all groups
            self._rebuild_comparison_groups_ui()
            
            # Update the UI
            if self.parent_gui:
                self.parent_gui._safe_page_update()
    
    def _create_comparison_group(self, group_id: str, group_number: int, available_sap_codes: list):
        """Create a comparison group with SAP checkboxes and test lab selection"""
        # Create SAP sections for this comparison group
        sap_sections = []
        group_sap_containers = {}
        color = self._color
        
        for sap in available_sap_codes:
            # Get selected tests count for this SAP
            selected_tests_count = 0
            if (self._has_selected_tests() and 
                self.parent_gui and 
                hasattr(self.parent_gui, 'state_manager') and 
                hasattr(self.parent_gui.state_manager, 'state')):
                selected_tests_count = len([test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                          if test.sap_code == sap])
            
            if selected_tests_count == 0:
                continue
                
            # Create SAP checkbox
            sap_label = f"{sap} ({selected_tests_count} test{'s' if selected_tests_count != 1 else ''} available)"
            sap_checkbox = ft.Checkbox(
                label=sap_label,
                value=False,  # Default unchecked for new comparison groups
                on_change=lambda e, sap_code=sap, grp_id=group_id: self._on_comparison_group_sap_checked(grp_id, sap_code, e.control.value),
                data=sap
            )
            logger.debug(f"Created SAP checkbox for {sap} in group {group_id}")
            
            # Container for test lab checkboxes (initially hidden)
            test_lab_container = ft.Container(
                content=ft.Column([]),
                visible=False,
                margin=ft.margin.only(left=30, top=5, bottom=10),
                bgcolor=color('surface_variant', '#f8f9fa'),
                padding=ft.padding.all(10),
                border_radius=5,
                border=ft.border.all(1, color('outline', '#e9ecef'))
            )
            
            group_sap_containers[sap] = test_lab_container
            
            # Create SAP section
            sap_section = ft.Column([
                sap_checkbox,
                test_lab_container
            ], spacing=2)
            
            sap_sections.append(sap_section)
        
        # Store the SAP containers for this group
        logger.debug(f"Storing SAP containers for group {group_id}: {list(group_sap_containers.keys())}")
        for group_info in self.comparison_groups:
            if group_info['id'] == group_id:
                group_info['sap_containers'] = group_sap_containers
                logger.debug(f"Successfully updated group {group_id} with SAP containers")
                break
        else:
            logger.warning(f"Group {group_id} not found in comparison_groups list! Available: {[g['id'] for g in self.comparison_groups]}")
        
        # Create the comparison group container
        comparison_group_container = ft.Container(
            content=ft.Column([
                # Header with title and delete button
                ft.Row([
                    ft.Text(
                        f"Comparison {group_number}",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color=color('warning', '#fbc02d')
                    ),
                    ft.Container(expand=True),
                    ft.IconButton(
                        icon=ft.Icons.DELETE,
                        icon_color=color('error', 'red'),
                        tooltip=f"Delete Comparison {group_number}",
                        on_click=lambda e, grp_id=group_id: self._delete_comparison_group(grp_id)
                    )
                ], spacing=10),
                
                ft.Text("Select SAP codes and test labs for this comparison. At least 2 test labs must be selected.", 
                       color=color('text_muted', 'grey'), size=12),
                
                # SAP sections
                ft.Column(sap_sections, spacing=5),
                
                # Validation message container
                ft.Container(
                    content=ft.Text("", size=12),
                    ref=ft.Ref[ft.Container](),
                    data=f"validation_{group_id}"
                )
            ], spacing=10),
            padding=ft.padding.all(15),
            margin=ft.margin.only(bottom=15),
            bgcolor=color('warning_container', '#fff8e1'),
            border_radius=8,
            border=ft.border.all(1, color('warning', '#ffcc02')),
            data=group_id
        )
        
        return comparison_group_container
    
    def _on_comparison_group_sap_checked(self, group_id: str, sap_code: str, checked: bool):
        """Handle SAP checkbox changes for a specific comparison group"""
        if not self.parent_gui:
            logger.debug(f"No parent GUI available for group {group_id}, SAP {sap_code}")
            return
        
        logger.debug(f"Comparison group SAP {sap_code} checked: {checked} for group {group_id}")
        
        # Find the group
        group_info = None
        for group in self.comparison_groups:
            if group['id'] == group_id:
                group_info = group
                break
        
        if not group_info:
            logger.debug(f"Group {group_id} not found in comparison_groups list")
            return
            
        if sap_code not in group_info['sap_containers']:
            logger.debug(f"SAP {sap_code} not found in group {group_id} containers. Available: {list(group_info['sap_containers'].keys())}")
            return
        
        container = group_info['sap_containers'][sap_code]
        logger.debug(f"Found container for SAP {sap_code}, setting visible={checked}")
        
        if checked:
            # Show test lab selection for this SAP in this group
            container.visible = True
            self._update_comparison_group_test_labs(group_id, sap_code)
        else:
            # Hide test lab selection
            container.visible = False
            container.content = ft.Column([])
            # Clear any selections for this SAP in this group
            self._clear_comparison_group_selections(group_id, sap_code)
        
        # Validate the group
        self._validate_comparison_group(group_id)
        
        logger.debug(f"Updating page for group {group_id}")
        if self.parent_gui:
            self.parent_gui._safe_page_update()
    
    def _update_comparison_group_test_labs(self, group_id: str, sap_code: str):
        """Update test lab checkboxes for a specific SAP in a specific comparison group"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            logger.debug(f"Cannot update test labs - no parent GUI or state manager")
            return
        
        logger.debug(f"Updating test labs for group {group_id}, SAP {sap_code}")
        color = self._color
        
        # Find the group
        group_info = None
        for group in self.comparison_groups:
            if group['id'] == group_id:
                group_info = group
                break
        
        if not group_info or sap_code not in group_info['sap_containers']:
            logger.debug(f"Group info not found or SAP container missing")
            return
        
        container = group_info['sap_containers'][sap_code]
        
        # Get available tests for this SAP (from step 2 selection)
        selected_tests_from_step2 = []
        if hasattr(self.parent_gui.state_manager.state, 'selected_tests'):
            selected_tests_from_step2 = [test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                        if test.sap_code == sap_code]
        
        logger.debug(f"Found {len(selected_tests_from_step2)} tests from step 2 for SAP {sap_code}")
        
        # Get currently selected test labs for this group and SAP
        selected_test_labs = set()
        if (hasattr(self.parent_gui.state_manager.state, 'comparison_groups') and
            group_id in self.parent_gui.state_manager.state.comparison_groups and
            sap_code in self.parent_gui.state_manager.state.comparison_groups[group_id]):
            selected_test_labs = self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code]
        
        test_lab_checkboxes = []
        if selected_tests_from_step2:
            test_lab_checkboxes.append(
                ft.Text(
                    f"Select test labs for {sap_code} in this comparison:",
                    size=12,
                    color=color('text_muted', 'grey'),
                    weight=ft.FontWeight.W_500
                )
            )
            
            # Add Select All / Select None buttons for this group
            select_all_btn = ft.ElevatedButton(
                "Select All",
                icon=ft.Icons.SELECT_ALL,
                on_click=lambda e, grp_id=group_id, sap=sap_code: self._select_all_group_test_labs(grp_id, sap),
                scale=0.8,
                bgcolor=color('primary_container', '#e3f2fd')
            )
            select_none_btn = ft.ElevatedButton(
                "Select None", 
                icon=ft.Icons.DESELECT,
                on_click=lambda e, grp_id=group_id, sap=sap_code: self._select_none_group_test_labs(grp_id, sap),
                scale=0.8,
                bgcolor=color('surface_variant', '#f5f5f5')
            )
            
            test_lab_checkboxes.append(
                ft.Row([select_all_btn, select_none_btn], spacing=5)
            )
            
            # Sort tests by test lab number for consistent display
            sorted_tests = sorted(selected_tests_from_step2, key=lambda t: t.test_lab_number)
            
            for test in sorted_tests:
                # Format voltage display
                voltage_display = f"{test.voltage}V" if test.voltage and test.voltage.strip() else "N/A"
                
                # Format notes display (truncate if too long)
                notes_display = test.notes if test.notes and test.notes.strip() else "No notes"
                if len(notes_display) > 40:
                    notes_display = notes_display[:37] + "..."
                
                # Check if this test lab is selected in this group
                is_test_lab_selected = test.test_lab_number in selected_test_labs
                
                checkbox = ft.Checkbox(
                    value=is_test_lab_selected,
                    on_change=lambda e, tl=test.test_lab_number, sap=sap_code, grp_id=group_id: self._on_comparison_group_test_lab_checked(grp_id, sap, tl, e.control.value),
                    scale=0.9
                )
                
                # Create detailed test info row
                test_info_row = ft.Container(
                    content=ft.Row([
                        checkbox,
                        ft.Container(
                            content=ft.Column([
                                ft.Text(
                                    f"Test: {test.test_lab_number}",
                                    size=12,
                                    weight=ft.FontWeight.W_500,
                                    color=color('primary', 'blue')
                                ),
                                ft.Row([
                                    ft.Text(
                                        f"Voltage: {voltage_display}",
                                        size=11,
                                        color=color('success', 'darkgreen')
                                    ),
                                    ft.Text("•", size=11, color=color('text_muted', 'grey')),
                                    ft.Text(
                                        f"Notes: {notes_display}",
                                        size=11,
                                        color=color('text_muted', 'grey'),
                                        tooltip=test.notes if test.notes else "No notes available"
                                    )
                                ], spacing=5)
                            ], spacing=2),
                            expand=True
                        )
                    ], alignment=ft.MainAxisAlignment.START, vertical_alignment=ft.CrossAxisAlignment.START),
                    padding=ft.padding.symmetric(horizontal=5, vertical=3),
                    margin=ft.margin.only(bottom=3),
                    bgcolor=color('surface_variant', '#f8f9fa') if is_test_lab_selected else None,
                    border_radius=3,
                    border=ft.border.all(1, color('primary', '#007bff')) if is_test_lab_selected else ft.border.all(1, color('outline', '#e9ecef'))
                )
                
                test_lab_checkboxes.append(test_info_row)
        else:
            test_lab_checkboxes.append(
                ft.Text(
                    "No test labs available for this SAP",
                    size=12,
                    color=color('warning', 'orange')
                )
            )
        
        container.content = ft.Column(test_lab_checkboxes, spacing=2)
        
        # Update UI
        if self.parent_gui:
            self.parent_gui._safe_page_update()
    
    def _select_all_group_test_labs(self, group_id: str, sap_code: str):
        """Select all test labs for a SAP in a specific comparison group"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
            
        # Get all available tests for this SAP from step 2
        selected_tests_from_step2 = []
        if hasattr(self.parent_gui.state_manager.state, 'selected_tests'):
            selected_tests_from_step2 = [test for test in self.parent_gui.state_manager.state.selected_tests.values() 
                                        if test.sap_code == sap_code]
        
        # Initialize comparison group data if needed
        if not hasattr(self.parent_gui.state_manager.state, 'comparison_groups'):
            self.parent_gui.state_manager.state.comparison_groups = {}
        
        if group_id not in self.parent_gui.state_manager.state.comparison_groups:
            self.parent_gui.state_manager.state.comparison_groups[group_id] = {}
            
        # Select all test lab numbers from step 2
        self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code] = set(
            test.test_lab_number for test in selected_tests_from_step2
        )
        
        # Refresh the UI for this group/SAP
        self._update_comparison_group_test_labs(group_id, sap_code)
        
        # Validate the group
        self._validate_comparison_group(group_id)
    
    def _select_none_group_test_labs(self, group_id: str, sap_code: str):
        """Deselect all test labs for a SAP in a specific comparison group"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
            
        # Initialize comparison group data if needed
        if not hasattr(self.parent_gui.state_manager.state, 'comparison_groups'):
            self.parent_gui.state_manager.state.comparison_groups = {}
        
        if group_id not in self.parent_gui.state_manager.state.comparison_groups:
            self.parent_gui.state_manager.state.comparison_groups[group_id] = {}
            
        # Clear selections
        self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code] = set()
        
        # Refresh the UI for this group/SAP
        self._update_comparison_group_test_labs(group_id, sap_code)
        
        # Validate the group
        self._validate_comparison_group(group_id)
    
    def _on_comparison_group_test_lab_checked(self, group_id: str, sap_code: str, test_lab: str, checked: bool):
        """Handle test lab checkbox changes for a specific comparison group"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager') or not hasattr(self.parent_gui.state_manager, 'state'):
            return
            
        # Initialize comparison group data in state if needed
        if not hasattr(self.parent_gui.state_manager.state, 'comparison_groups'):
            self.parent_gui.state_manager.state.comparison_groups = {}
        
        if group_id not in self.parent_gui.state_manager.state.comparison_groups:
            self.parent_gui.state_manager.state.comparison_groups[group_id] = {}
        
        if sap_code not in self.parent_gui.state_manager.state.comparison_groups[group_id]:
            self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code] = set()
        
        # Add or remove the test lab
        if checked:
            self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code].add(test_lab)
        else:
            self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code].discard(test_lab)
        
        # Validate the group
        self._validate_comparison_group(group_id)
    
    def _validate_comparison_group(self, group_id: str):
        """Validate that a comparison group has at least 2 test labs selected"""
        if (not self.parent_gui or 
            not hasattr(self.parent_gui, 'state_manager') or 
            not hasattr(self.parent_gui.state_manager, 'state') or
            not hasattr(self.parent_gui.state_manager.state, 'comparison_groups')):
            return
        color = self._color
        
        # Count total selected test labs in this group
        total_test_labs = 0
        group_data = self.parent_gui.state_manager.state.comparison_groups.get(group_id, {})
        
        for sap_code, test_labs in group_data.items():
            total_test_labs += len(test_labs)
        
        # Find the validation message container for this group
        group_info = None
        for group in self.comparison_groups:
            if group['id'] == group_id:
                group_info = group
                break
        
        if group_info and hasattr(group_info['container'], 'content'):
            # Find the validation container
            validation_container = None
            for control in group_info['container'].content.controls:
                if hasattr(control, 'data') and control.data == f"validation_{group_id}":
                    validation_container = control
                    break
            
            if validation_container:
                if total_test_labs < 2:
                    validation_container.content = ft.Text(
                        f"⚠️ At least 2 test labs required for comparison (currently {total_test_labs} selected)",
                        size=12,
                        color=color('warning', 'orange'),
                        weight=ft.FontWeight.W_500
                    )
                    validation_container.bgcolor = color('warning_container', '#fff3e0')
                    validation_container.padding = ft.padding.all(8)
                    validation_container.border_radius = 4
                else:
                    validation_container.content = ft.Text(
                        f"✅ Valid comparison with {total_test_labs} test labs selected",
                        size=12,
                        color=color('success', 'green'),
                        weight=ft.FontWeight.W_500
                    )
                    validation_container.bgcolor = color('success_container', '#e8f5e8')
                    validation_container.padding = ft.padding.all(8)
                    validation_container.border_radius = 4
                
                if self.parent_gui:
                    self.parent_gui._safe_page_update()
    
    def _clear_comparison_group_selections(self, group_id: str, sap_code: str):
        """Clear selections for a specific SAP in a comparison group"""
        if (self.parent_gui and 
            hasattr(self.parent_gui, 'state_manager') and 
            hasattr(self.parent_gui.state_manager, 'state') and
            hasattr(self.parent_gui.state_manager.state, 'comparison_groups') and
            group_id in self.parent_gui.state_manager.state.comparison_groups and
            sap_code in self.parent_gui.state_manager.state.comparison_groups[group_id]):
            
            self.parent_gui.state_manager.state.comparison_groups[group_id][sap_code].clear()
    
    def _delete_comparison_group(self, group_id: str):
        """Delete a specific comparison group"""
        if (not self.comparison_groups_container or 
            not hasattr(self.comparison_groups_container, 'current') or
            not self.comparison_groups_container.current):
            return
        
        # Find and remove from comparison_groups list
        group_to_remove = None
        for i, group in enumerate(self.comparison_groups):
            if group['id'] == group_id:
                group_to_remove = group
                self.comparison_groups.pop(i)
                break
        
        if group_to_remove:
            # Clear from state first
            if (self.parent_gui and
                hasattr(self.parent_gui, 'state_manager') and
                hasattr(self.parent_gui.state_manager, 'state') and
                hasattr(self.parent_gui.state_manager.state, 'comparison_groups') and
                group_id in self.parent_gui.state_manager.state.comparison_groups):
                del self.parent_gui.state_manager.state.comparison_groups[group_id]
            
            # For UI update, we'll rebuild the comparison groups container
            # This is safer than trying to manipulate controls directly
            self._rebuild_comparison_groups_ui()
            
            # Update UI
            if self.parent_gui:
                self.parent_gui._safe_page_update()
    
    def _rebuild_comparison_groups_ui(self):
        """Rebuild the comparison groups UI after deletion"""
        if (not self.comparison_groups_container or 
            not hasattr(self.comparison_groups_container, 'current') or
            not self.comparison_groups_container.current):
            return
        
        try:
            # Clear the container and rebuild with remaining groups
            if hasattr(self.comparison_groups_container.current, 'content'):
                # Create new content with remaining groups
                remaining_containers = []
                for group_info in self.comparison_groups:
                    if 'container' in group_info:
                        remaining_containers.append(group_info['container'])
                
                # Update the container content
                self.comparison_groups_container.current.content = ft.Column(remaining_containers, spacing=10)
        except Exception as e:
            logger.error(f"Error rebuilding comparison groups UI: {e}")
    
    def _get_timestamp(self):
        """Get current timestamp"""
        import datetime
        return datetime.datetime.now().isoformat()
    
    # Keep the old dialog methods for compatibility but they won't be used

