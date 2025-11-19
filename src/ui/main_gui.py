"""
Modularized Main GUI for the Motor Report Generator Application.

This module acts as the central coordinator, initializing and wiring up the various
managers and UI components that make up the application. It delegates all logic
for state, events, and workflows to the specialized managers in the `core` directory.

Import Pattern:
--------------
This module uses relative imports from the gui subpackage (e.g., gui.core, gui.tabs)
while ensuring the project root is in sys.path. This pattern works for:
- Running as script: python src/gui/main_gui.py
- Running as module: python -m src.gui.main_gui
- Packaged executable: PyInstaller bundle
"""
import flet as ft
import logging
import sys
from pathlib import Path
import threading
import time
from typing import Optional

# Add project root to path for clean imports
try:
    project_root = Path(__file__).resolve().parent.parent.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))
except (NameError, IndexError):
    # Fallback for environments where __file__ is not defined
    project_root = Path.cwd()
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

# Relative imports from ui subpackage (correct pattern for submodules)
from .core.state_manager import StateManager
from .core.event_handlers import EventHandlers
from .core.workflow_manager import WorkflowManager
from .core.search_manager import SearchManager
from .core.report_manager import ReportManager

# UI Tabs
from .tabs.setup_tab import SetupTab
from .tabs.search_select_tab import SearchSelectTab
from .tabs.config_tab import ConfigTab
from .tabs.generate_tab import GenerateTab

# UI Components & Utils
from .core.status_manager import StatusManager
from .components.base import ProgressIndicators
from .utils.thread_pool import run_in_background, shutdown_thread_pool
from .theme import resolve_token, set_user_theme

# Project specific imports
from ..config.directory_config import PROJECT_ROOT
from ..config.directory_config import ensure_directories_initialized
from ..core.motor_report_engine import MotorReportApp
from ..config.app_config import AppConfig
from ..core.telemetry import log_duration
try:
    from src._version import VERSION
except Exception:
    from .._version import VERSION

logger = logging.getLogger(__name__)

class MotorReportAppGUI:
    """
    The main GUI class, responsible for building the UI and coordinating managers.
    """
    def __init__(self, page: ft.Page):
        self.page = page
        self._current_theme_mode: str = "system"
        self.app: Optional[MotorReportApp] = None

        # Initialize core managers first with profiling
        logger.info("Initializing Motor Report GUI...")
        
        with log_duration(logger, "StateManager initialization", level=logging.DEBUG):
            self.state_manager = StateManager()
        
        with log_duration(logger, "WorkflowManager initialization", level=logging.DEBUG):
            self.workflow_manager = WorkflowManager(self)
        
        # Lazy initialization for SearchManager and ReportManager (initialized on first access)
        self._search_manager: Optional[SearchManager] = None
        self._report_manager: Optional[ReportManager] = None
        
        with log_duration(logger, "EventHandlers initialization", level=logging.DEBUG):
            self.event_handlers = EventHandlers(self)
        
        # Now setup page settings (which needs event_handlers)
        self._setup_page_settings()
        
        # Initialize UI components (search_manager will be lazily initialized when accessed)
        self._initialize_ui_components()
        self._setup_file_pickers()
        
        # Build the main layout
        self.build_layout()

        # Post-build initializations
        self.event_handlers.handle_post_build_setup()
        self.workflow_manager.update_workflow_state()
        
        # Final robust refresh of all components after a short delay
        run_in_background(self._delayed_initial_refresh)
    
    @property
    def search_manager(self) -> SearchManager:
        """
        Lazy initialization of SearchManager.
        
        SearchManager is created on first access to reduce startup time.
        Only initialized when search functionality is actually needed.
        
        Returns:
            SearchManager instance
        """
        if self._search_manager is None:
            with log_duration(logger, "SearchManager lazy initialization", level=logging.DEBUG):
                self._search_manager = SearchManager(self)
        return self._search_manager
    
    @property
    def report_manager(self) -> ReportManager:
        """
        Lazy initialization of ReportManager.
        
        ReportManager is created on first access to reduce startup time.
        Only initialized when report generation functionality is actually needed.
        
        Returns:
            ReportManager instance
        """
        if self._report_manager is None:
            with log_duration(logger, "ReportManager lazy initialization", level=logging.DEBUG):
                self._report_manager = ReportManager(self)
        return self._report_manager

    def _setup_page_settings(self):
        """Configures the main Flet page settings."""
        self.page.title = "Motor Report Generator"
        self.page.scroll = ft.ScrollMode.ADAPTIVE
        self.page.vertical_alignment = ft.MainAxisAlignment.START
        self.page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH
        self.page.padding = ft.padding.all(20)
        self.page.on_disconnect = self._handle_disconnect
        self.page.on_platform_brightness_change = self._handle_platform_brightness_change

    def _initialize_ui_components(self):
        """Initializes all UI components and controls."""
        # Helper to resolve themed colors (falls back to plain strings)
        def themed_color(token: str, fallback: str):
            return resolve_token(self.page, token, fallback) or fallback

        self._themed_color = themed_color
        # Setup Tab Components
        self.tests_folder_path_text = ft.Text("No test folder selected.", italic=True, max_lines=2)
        self.registry_file_path_text = ft.Text("No registry file selected.", italic=True, max_lines=2)
        self.noise_folder_path_text = ft.Text("No noise folder selected.", italic=True, max_lines=2)
        self.noise_registry_path_text = ft.Text("No noise registry file selected.", italic=True, max_lines=2)
        
        self.select_tests_folder_button = ft.ElevatedButton(
            "Select Tests Folder", icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda e: self.folder_picker.get_directory_path(dialog_title="Select Performance Tests Folder")
        )
        self.select_registry_file_button = ft.ElevatedButton(
            "Select Registry File", icon=ft.Icons.DESCRIPTION,
            on_click=lambda e: self.file_picker.pick_files(dialog_title="Select Lab Registry File", allowed_extensions=["xlsx", "xls"])
        )
        self.select_noise_folder_button = ft.ElevatedButton(
            "Select Noise Folder", icon=ft.Icons.FOLDER_OPEN,
            on_click=lambda e: self.folder_picker.get_directory_path(dialog_title="Select Noise Tests Folder")
        )
        self.select_noise_registry_button = ft.ElevatedButton(
            "Select Noise Registry", icon=ft.Icons.DESCRIPTION,
            on_click=lambda e: self.file_picker.pick_files(dialog_title="Select Noise Registry File", allowed_extensions=["xlsx", "xls"])
        )

        # Search Tab Components
        self.search_input_field = ft.TextField(
            label="Enter SAP Code or Test Number", 
            hint_text="Try: 1,2,3 or 6120890812,6120890848",
            autofocus=True, width=400, on_submit=self.event_handlers.on_search_clicked
        )
        self.search_button = ft.ElevatedButton("Search", icon=ft.Icons.SEARCH, on_click=self.event_handlers.on_search_clicked)
        self.results_filter_inputs = {
            "test_lab": ft.TextField(
                label="Test Lab",
                hint_text="Filter",
                dense=True,
                width=140,
                on_change=self.search_manager.generate_filter_handler("test_lab")
            ),
            "date": ft.TextField(
                label="Date",
                hint_text="Filter",
                dense=True,
                width=140,
                on_change=self.search_manager.generate_filter_handler("date")
            ),
            "voltage": ft.TextField(
                label="Voltage",
                hint_text="Filter",
                dense=True,
                width=120,
                on_change=self.search_manager.generate_filter_handler("voltage")
            ),
            "notes": ft.TextField(
                label="Notes",
                hint_text="Filter",
                dense=True,
                width=240,
                expand=True,
                on_change=self.search_manager.generate_filter_handler("notes")
            )
        }
        self._style_primary_text_inputs()
        self.clear_filters_button = ft.TextButton(
            "Clear Filters",
            icon=ft.Icons.CLEAR_ALL,
            on_click=self.search_manager.clear_filters
        )
        self.results_filter_icon = ft.Icon(
            ft.Icons.FILTER_ALT,
            color=self._themed_color('primary', 'blue')
        )
        self.results_filters_row = ft.Container(
            content=ft.Row([
                self.results_filter_icon,
                self.results_filter_inputs["test_lab"],
                self.results_filter_inputs["date"],
                self.results_filter_inputs["voltage"],
                self.results_filter_inputs["notes"],
                self.clear_filters_button
            ], spacing=10, vertical_alignment=ft.CrossAxisAlignment.END),
            padding=ft.padding.only(top=10, bottom=5)
        )
        self.search_manager.register_filter_inputs(self.results_filter_inputs)
        # Theme-aware defaults for search placeholder
        placeholder_text_color = self._themed_color('on_surface', 'grey')
        placeholder_accent = self._themed_color('primary', ft.Colors.BLUE_700)

        self.results_area = ft.Column(
            controls=[
                ft.Container(
                    content=ft.Column([
                        ft.Icon(ft.Icons.SEARCH, size=48, color=placeholder_text_color),
                        ft.Text("Enter a search term to find tests", size=16, color=placeholder_text_color),
                        ft.Text("Valid examples:", size=14, weight=ft.FontWeight.W_500, color=placeholder_accent),
                        ft.Text("• Test numbers: 1, 2, 3 or 1,2,3", size=12, color=placeholder_text_color),
                        ft.Text("• SAP codes: 6120890812, 6120890848", size=12, color=placeholder_text_color),
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=ft.padding.all(20),
                    alignment=ft.alignment.center
                )
            ],
            spacing=5,
            scroll=ft.ScrollMode.ADAPTIVE,
            expand=True,
        )
        
        # SAP Navigation Container (appears above results when multiple SAP codes found)
        self.sap_navigation_container = ft.Container(
            content=ft.Text(""),  # Placeholder
            visible=False
        )
        self.selected_count_text = ft.Text("Selected: 0 tests", visible=False)
        
        # Workflow buttons - removed Apply/Clear Selection buttons for seamless workflow
        self.confirm_selection_button = ft.ElevatedButton("Confirm Selection", visible=False, on_click=self.event_handlers.on_confirm_test_selection)
        self.modify_selection_button = ft.ElevatedButton("Modify Selection", visible=False, on_click=self.event_handlers.on_modify_test_selection)
        self.new_search_button = ft.ElevatedButton("New Search", visible=False, on_click=self.event_handlers.on_start_new_search)
        
        # Config Tab Navigation Controls (for Step 3)
        # Config tab apply button (separate reference for state management)
        self.config_apply_button = ft.ElevatedButton(
            "Apply Selection",
            icon=ft.Icons.CHECK_CIRCLE,
            bgcolor=self._themed_color('success', 'green'),
            color=self._themed_color('on_success', 'white'),
            on_click=self.event_handlers.on_apply_config_selection if self._has_event_handlers() else None
        )
        self.config_clear_button = ft.ElevatedButton(
            "Clear Selection",
            icon=ft.Icons.CLEAR_ALL,
            bgcolor=self._themed_color('surface_variant', 'grey'),
            color=self._themed_color('on_surface', 'white'),
            on_click=self.event_handlers.on_clear_config_selection if self._has_event_handlers() else None
        )
        
        self.config_navigation_container = ft.Container(
            content=ft.Row([
                self.config_apply_button,
                self.config_clear_button,
            ], alignment=ft.MainAxisAlignment.END, spacing=10),
            visible=True
        )

        # Config Tab Components
        self.include_noise_switch = ft.Switch(label="Include Noise Data", value=True, on_change=lambda e: self._handle_switch_change("include_noise", e.control.value))
        self.include_comparison_switch = ft.Switch(label="Generate Comparison Sheet", value=True, on_change=lambda e: self._handle_switch_change("include_comparison", e.control.value))
        self.registry_sheet_name_input = ft.TextField(label="Registry Sheet Name", value="REGISTRO", width=400)
        self.noise_registry_sheet_name_input = ft.TextField(label="Noise Registry Sheet Name", value="Registro", width=400)
        self.config_summary_area = ft.Column()

        # Generate Tab Components
        self.generation_summary = ft.Markdown("Report is not yet configured.")
        self.generate_button = ft.ElevatedButton(
            "Generate Report", icon=ft.Icons.CREATE, on_click=self.event_handlers.on_generate_report_clicked,
            bgcolor=self._themed_color('primary', ft.Colors.BLUE_700),
            color=self._themed_color('on_primary', ft.Colors.WHITE),
        )

        # Global components
        self.status_text = ft.Text("Welcome!", size=12, expand=True)
        self.progress_bar = ft.ProgressBar(width=200, visible=False)
        self.status_bar = ft.Row([self.status_text, self.progress_bar], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        
        self.status_manager = StatusManager(
            self.status_text,
            self.progress_bar,
            self._safe_page_update,
            color_resolver=self._resolve_status_color,
        )
        self.progress_indicators = ProgressIndicators()
        self.progress_indicators.create_indicators(4)
        
        # Initialize SAP navigator (updated import path after refactor)
        from .utils.helpers import SAPNavigationManager
        self.sap_navigator = SAPNavigationManager(self)

        # Theme toggle in header (Dropdown persisted in client_storage)
        try:
            stored_theme = None
            if hasattr(self.page, 'client_storage') and self.page.client_storage is not None:
                try:
                    stored_theme = self.page.client_storage.get('theme')
                except Exception:
                    stored_theme = None

            theme_value = stored_theme or 'system'
            self._current_theme_mode = theme_value
            self.theme_dropdown = ft.Dropdown(
                label='Theme',
                value=theme_value,
                options=[ft.dropdown.Option(o) for o in ("system", "light", "dark")],
                on_change=lambda e: self._on_theme_changed(e.control.value),
                width=160,
            )
            self._style_theme_dropdown()
        except Exception:
            self.theme_dropdown = ft.Container()

        # Ensure global controls pick up the active palette immediately
        self._apply_theme_colors_to_global_controls()

    def _setup_file_pickers(self):
        """Sets up the file and folder pickers."""
        self.folder_picker = ft.FilePicker(on_result=self.event_handlers.on_folder_picked)
        self.file_picker = ft.FilePicker(on_result=self.event_handlers.on_registry_file_picked)
        self.save_file_picker = ft.FilePicker(on_result=self.event_handlers.on_save_file_picked)
        self.page.overlay.extend([self.folder_picker, self.file_picker, self.save_file_picker])

    def build_layout(self):
        """Constructs the main UI layout with a tabbed interface."""
        logger.debug("Building main UI layout...")
        
        # Create tab instances with profiling
        with log_duration(logger, "SetupTab creation", level=logging.DEBUG):
            self.setup_tab = SetupTab(self)
        
        with log_duration(logger, "SearchSelectTab creation", level=logging.DEBUG):
            self.search_select_tab = SearchSelectTab(self)
        
        with log_duration(logger, "ConfigTab creation", level=logging.DEBUG):
            self.config_tab = ConfigTab(self)
        
        with log_duration(logger, "GenerateTab creation", level=logging.DEBUG):
            self.generate_tab = GenerateTab(self)
        
        self._tab_icon_defs = [
            ft.Icons.SETTINGS,
            ft.Icons.SEARCH,
            ft.Icons.TUNE,
            ft.Icons.CREATE,
        ]

        self.tabs = ft.Tabs(
            selected_index=0,
            on_change=self.event_handlers.on_tab_change,
            tabs=[
                ft.Tab(text="1. Setup", icon=self._tab_icon(self._tab_icon_defs[0]), content=self.setup_tab.get_tab_content()),
                ft.Tab(text="2. Search & Select", icon=self._tab_icon(self._tab_icon_defs[1]), content=self.search_select_tab.get_tab_content()),
                ft.Tab(text="3. Configure", icon=self._tab_icon(self._tab_icon_defs[2]), content=self.config_tab.get_tab_content()),
                ft.Tab(text="4. Generate", icon=self._tab_icon(self._tab_icon_defs[3]), content=self.generate_tab.get_tab_content()),
            ],
            expand=True,
        )

        self._refresh_tab_icons()

        # Header row with app title and theme toggle
        # Wrap theme control in a small padded container so its popup isn't clipped
        theme_control = getattr(self, 'theme_dropdown', ft.Container())
        theme_control_wrapper = ft.Container(
            content=theme_control,
            padding=ft.padding.symmetric(vertical=4, horizontal=6),
            margin=ft.margin.only(top=6),
            alignment=ft.alignment.center_right
        )

        header_row = ft.Row([
            ft.Text(f"Motor Report Generator v{VERSION}", size=32, weight=ft.FontWeight.BOLD),
            ft.Container(expand=True),
            # place themed control in wrapper to avoid clipping at the top edge
            theme_control_wrapper
        ], vertical_alignment=ft.CrossAxisAlignment.CENTER)

        self.page.add(
            ft.Column(
                [
                    header_row,
                    ft.Divider(),
                    self.tabs,
                    ft.Divider(),
                    self.status_bar,
                ],
                expand=True, spacing=10
            )
        )

    def _delayed_initial_refresh(self):
        """Waits a moment then refreshes the UI to ensure all components are rendered correctly."""
        time.sleep(0.2)
        self.refresh_components(['paths', 'setup'])
        logger.info("Initial UI refresh completed")

    def refresh_components(self, components=None):
        """
        Unified component refresh method.
        Args:
            components: List of component types to refresh ['paths', 'setup', 'search', 'all']
        """
        if components is None:
            components = ['all']
        
        try:
            if 'paths' in components or 'all' in components:
                self._refresh_all_path_displays()
                self._update_path_displays_with_auto_detected_values()
            
            if 'setup' in components or 'all' in components:
                self._safe_page_update()
            
            if 'search' in components or 'all' in components:
                self._display_search_results()
                self._safe_page_update()
                
        except Exception as e:
            logger.error(f"Error refreshing components {components}: {e}")

    def _safe_page_update(self):
        """Optimized page update method for better performance"""
        try:
            # Check if page and session are still valid
            if not self._has_valid_page_session():
                return False
            
            # PERFORMANCE OPTIMIZATION: Batch updates
            update_success = False
            
            try:
                # Single page update - most efficient approach
                self.page.update()
                update_success = True
                logger.debug("Optimized page update successful")
            except Exception as e:
                error_msg = str(e).lower()
                # Ignore "control must be added" errors during initialization
                if "must be added to the page" in error_msg or "control must be" in error_msg:
                    logger.debug(f"Page update skipped - controls not yet added: {e}")
                    return False
                
                logger.debug(f"Page update failed: {e}")
                # Fallback: try component-level updates
                try:
                    if hasattr(self, 'tabs') and self.tabs:
                        self.tabs.update()
                    if hasattr(self, 'status_bar') and self.status_bar:
                        self.status_bar.update()
                    update_success = True
                    logger.debug("Component-level updates successful")
                except Exception as fallback_e:
                    logger.debug(f"Fallback updates also failed: {fallback_e}")
            
            return update_success
                
        except (RuntimeError, AttributeError, Exception) as e:
            # Handle shutdown and session-related errors quietly
            if "shutdown" in str(e).lower() or "session" in str(e).lower():
                logger.debug(f"Page update skipped due to shutdown: {e}")
            else:
                logger.debug(f"Page update error: {e}")
        return False

    def _resolve_status_color(self, color: Optional[str]) -> Optional[str]:
        """Map legacy color keywords to themed tokens for better contrast."""
        if not hasattr(self, '_themed_color'):
            return color

        if not color:
            return self._themed_color('on_surface', '#1f2933')

        normalized = str(color).lower()
        if normalized.startswith('#'):
            return color

        token_map = {
            'green': ('success', '#2e7d32'),
            'red': ('error', '#c62828'),
            'blue': ('info', '#0288d1'),
            'orange': ('warning', '#f57c00'),
            'black': ('on_surface', '#1f2933'),
            'white': ('on_surface', '#f1f5f9'),
        }

        token, fallback = token_map.get(normalized, ('on_surface', color))
        return self._themed_color(token, fallback if isinstance(fallback, str) else color)

    def _apply_theme_mode(self, mode: str, *, reason: str = "user"):
        """Central handler for applying a requested theme mode."""
        if mode not in ("light", "dark", "system"):
            return

        self._current_theme_mode = mode

        if reason != "platform":
            self._persist_theme_preference(mode)

        try:
            set_user_theme(self.page, mode)
        except Exception as exc:
            logger.debug(f"set_user_theme failed: {exc}")

        self._sync_theme_dropdown_value(mode)
        self._refresh_theme_after_change(reason)

    def _persist_theme_preference(self, mode: str):
        """Persist the chosen theme mode to client storage for future sessions."""
        storage = getattr(self.page, 'client_storage', None)
        if not storage:
            return
        try:
            storage.set('theme', mode)
        except Exception as exc:
            logger.debug(f"Could not persist theme preference: {exc}")

    def _sync_theme_dropdown_value(self, mode: str):
        """Keep the theme dropdown selection aligned with the applied mode."""
        dropdown = getattr(self, 'theme_dropdown', None)
        if not dropdown or not hasattr(dropdown, 'value'):
            return
        if dropdown.value == mode:
            return
        try:
            dropdown.value = mode
            dropdown.update()
        except Exception:
            pass

    def _refresh_theme_after_change(self, reason: str):
        """Refresh UI controls after a theme change to pick up new tokens."""
        self._apply_theme_colors_to_global_controls()
        self._rebuild_tabs_for_theme()

        if hasattr(self, 'search_manager') and self._search_manager:
            try:
                self.search_manager.display_search_results()
            except Exception as exc:
                logger.debug(f"Search results refresh failed after theme change: {exc}")

        try:
            if hasattr(self, 'status_manager') and self.status_manager:
                source = "system" if reason == "platform" else reason
                self.status_manager.update_status(
                    f"Theme set to {self._current_theme_mode} ({source})",
                    "blue"
                )
        except Exception as exc:
            logger.debug(f"Status update failed after theme change: {exc}")

        self._safe_page_update()

    def _apply_theme_colors_to_global_controls(self):
        """Reapply semantic colors to shared controls that are not rebuilt automatically."""
        if not hasattr(self, '_themed_color'):
            return
        try:
            if hasattr(self, 'config_apply_button') and self.config_apply_button:
                self.config_apply_button.bgcolor = self._themed_color('success', 'green')
                self.config_apply_button.color = self._themed_color('on_success', 'white')
            if hasattr(self, 'config_clear_button') and self.config_clear_button:
                self.config_clear_button.bgcolor = self._themed_color('surface_variant', 'grey')
                self.config_clear_button.color = self._themed_color('on_surface', 'white')
            if hasattr(self, 'generate_button') and self.generate_button:
                self.generate_button.bgcolor = self._themed_color('primary', ft.Colors.BLUE_700)
                self.generate_button.color = self._themed_color('on_primary', ft.Colors.WHITE)
            if hasattr(self, 'results_filter_icon') and self.results_filter_icon:
                self.results_filter_icon.color = self._themed_color('primary', 'blue')
            self._style_primary_text_inputs()
            self._style_theme_dropdown()
            if hasattr(self, 'setup_tab') and hasattr(self.setup_tab, 'apply_textfield_theme'):
                self.setup_tab.apply_textfield_theme()
            self._refresh_tab_icons()
        except Exception as exc:
            logger.debug(f"Global control theming failed: {exc}")

    def _style_theme_dropdown(self):
        dropdown = getattr(self, 'theme_dropdown', None)
        if not dropdown or not hasattr(self, '_themed_color'):
            return
        try:
            dropdown.color = self._themed_color('on_surface', '#1f2933')
            dropdown.bgcolor = self._themed_color('surface', '#f7f9fc')
            dropdown.fill_color = self._themed_color('surface_variant', '#e0e7ef')
            dropdown.border_color = self._themed_color('outline', '#c5cbd3')
            dropdown.focused_border_color = self._themed_color('primary', '#1565c0')
            dropdown.icon_color = self._themed_color('on_surface', '#ffffff')
            dropdown.label_style = ft.TextStyle(color=self._themed_color('on_surface', '#ffffff'))
            dropdown.hint_style = ft.TextStyle(color=self._themed_color('on_surface', '#ffffff'))
            dropdown.text_style = ft.TextStyle(color=self._themed_color('on_surface', '#ffffff'))
        except Exception as exc:
            logger.debug(f"Theme dropdown styling failed: {exc}")

    def _style_primary_text_inputs(self):
        if not hasattr(self, '_themed_color'):
            return
        inputs = []
        if hasattr(self, 'search_input_field') and self.search_input_field:
            inputs.append(self.search_input_field)
        if hasattr(self, 'results_filter_inputs'):
            inputs.extend(self.results_filter_inputs.values())
        for text_field in inputs:
            self._style_text_field(text_field)

    def _style_text_field(self, control: Optional[ft.TextField]):
        if not control or not hasattr(self, '_themed_color'):
            return
        try:
            readable = self._themed_color('on_surface', '#ffffff')
            muted = self._themed_color('text_muted', '#d1d9e6')
            field_bg = self._themed_color('surface_variant', '#2a3340')
            outline = self._themed_color('outline', '#8895a8')
            control.color = readable
            control.bgcolor = field_bg
            control.fill_color = field_bg
            control.border_color = outline
            control.focused_border_color = self._themed_color('primary', '#90caf9')
            control.text_style = ft.TextStyle(color=readable)
            control.label_style = ft.TextStyle(color=readable)
            control.hint_style = ft.TextStyle(color=muted)
        except Exception as exc:
            logger.debug(f"Text field styling failed: {exc}")

    def _tab_icon(self, icon_value):
        color = self._themed_color('on_surface', '#fefefe') if hasattr(self, '_themed_color') else '#ffffff'
        if isinstance(icon_value, ft.Icon):
            icon_value.color = color
            return icon_value
        return ft.Icon(icon_value, color=color)

    def _refresh_tab_icons(self):
        if not hasattr(self, 'tabs') or not getattr(self.tabs, 'tabs', None):
            return
        icon_defs = getattr(self, '_tab_icon_defs', None)
        if not icon_defs:
            return
        for idx, icon_value in enumerate(icon_defs):
            if idx >= len(self.tabs.tabs):
                continue
            try:
                self.tabs.tabs[idx].icon = self._tab_icon(icon_value)
            except Exception:
                continue
        try:
            self.tabs.update()
        except Exception:
            pass

    def _rebuild_tabs_for_theme(self):
        """Rebuild tab contents so that all themed controls re-evaluate their tokens."""
        if not hasattr(self, 'tabs') or not getattr(self.tabs, 'tabs', None):
            return

        tab_map = [
            (0, 'setup_tab'),
            (1, 'search_select_tab'),
            (2, 'config_tab'),
            (3, 'generate_tab'),
        ]

        for index, attr in tab_map:
            if index >= len(self.tabs.tabs):
                continue
            tab_instance = getattr(self, attr, None)
            if not tab_instance:
                continue
            try:
                self.tabs.tabs[index].content = tab_instance.get_tab_content()
            except Exception as exc:
                logger.debug(f"Unable to rebuild tab {attr}: {exc}")

    def _handle_platform_brightness_change(self, e):
        """Auto-apply theme changes when running in system mode and the OS updates its theme."""
        if self._current_theme_mode == "system":
            self._apply_theme_mode("system", reason="platform")

    def _refresh_all_path_displays(self):
        """Force refresh all path display components."""
        self._safe_component_update(self.tests_folder_path_text)
        self._safe_component_update(self.registry_file_path_text)
        self._safe_component_update(self.noise_folder_path_text)
        self._safe_component_update(self.noise_registry_path_text)

    def _update_path_displays_with_auto_detected_values(self):
        """Update path displays with auto-detected values from directory configuration."""
        try:
            from ..config.directory_config import (
                PERFORMANCE_TEST_DIR,
                LAB_REGISTRY_FILE,
                NOISE_TEST_DIR,
                NOISE_REGISTRY_FILE,
                TEST_LAB_CARICHI_DIR,
            )
            
            # Performance test folder
            success_color = self._themed_color('success', '#2e7d32')
            warning_color = self._themed_color('warning', '#f57c00')

            if PERFORMANCE_TEST_DIR and PERFORMANCE_TEST_DIR.exists():
                self.tests_folder_path_text.value = f"🎯 Auto-detected: {PERFORMANCE_TEST_DIR}"
                self.tests_folder_path_text.color = success_color
                self.tests_folder_path_text.italic = False
                # Update state manager with auto-detected path
                if self.state_manager:
                    self.state_manager.update_paths(tests_folder=str(PERFORMANCE_TEST_DIR))
            else:
                self.tests_folder_path_text.value = "⚠️ No test folder auto-detected. Please select manually."
                self.tests_folder_path_text.color = warning_color
            
            # Performance registry file
            if LAB_REGISTRY_FILE and LAB_REGISTRY_FILE.exists():
                self.registry_file_path_text.value = f"🎯 Auto-detected: {LAB_REGISTRY_FILE}"
                self.registry_file_path_text.color = success_color
                self.registry_file_path_text.italic = False
                # Update state manager with auto-detected path
                if self.state_manager:
                    self.state_manager.update_paths(registry_file=str(LAB_REGISTRY_FILE))
            else:
                self.registry_file_path_text.value = "⚠️ No registry file auto-detected. Please select manually."
                self.registry_file_path_text.color = warning_color
            
            # Noise test folder
            if NOISE_TEST_DIR and NOISE_TEST_DIR.exists():
                self.noise_folder_path_text.value = f"🎯 Auto-detected: {NOISE_TEST_DIR}"
                self.noise_folder_path_text.color = success_color
                self.noise_folder_path_text.italic = False
                # Update state manager with auto-detected path
                if self.state_manager:
                    self.state_manager.update_paths(noise_folder=str(NOISE_TEST_DIR))
            else:
                self.noise_folder_path_text.value = "⚠️ No noise folder auto-detected. Please select manually."
                self.noise_folder_path_text.color = warning_color
            
            # Noise registry file
            if NOISE_REGISTRY_FILE and NOISE_REGISTRY_FILE.exists():
                self.noise_registry_path_text.value = f"🎯 Auto-detected: {NOISE_REGISTRY_FILE}"
                self.noise_registry_path_text.color = success_color
                self.noise_registry_path_text.italic = False
                # Update state manager with auto-detected path
                if self.state_manager:
                    self.state_manager.update_paths(noise_registry=str(NOISE_REGISTRY_FILE))
            else:
                self.noise_registry_path_text.value = "⚠️ No noise registry auto-detected. Please select manually."
                self.noise_registry_path_text.color = warning_color

            # Test Lab (Carichi Nominali) directory
            if TEST_LAB_CARICHI_DIR and TEST_LAB_CARICHI_DIR.exists() and self.state_manager:
                self.state_manager.update_paths(test_lab_dir=str(TEST_LAB_CARICHI_DIR))
            
            # Initialize backend with auto-detected paths
            if PERFORMANCE_TEST_DIR and LAB_REGISTRY_FILE:
                run_in_background(self._initialize_backend)
                self.status_manager.update_status("🎯 Auto-detected paths loaded successfully. Backend initialized.", "green")
            
            self._safe_page_update()
            
        except Exception as e:
            logger.error(f"Error updating path displays: {e}")
            self.status_manager.update_status("⚠️ Error loading auto-detected paths", "orange")



    def _safe_component_update(self, component: ft.Control, force_page_update=False):
        """Safely update a specific UI component with retries."""
        if not component: return False
        try:
            component.update()
            if force_page_update:
                self._safe_page_update()
            return True
        except Exception as e:
            logger.warning(f"Could not update component {type(component).__name__}: {e}")
            return False

    def _handle_switch_change(self, config_key: str, value: bool):
        """Handle switch change events with visual feedback"""
        self.status_manager.update_status(f"⚙️ Updated {config_key}: {'enabled' if value else 'disabled'}", "blue")
        self.event_handlers.on_configuration_changed(config_key, value)
        self._safe_page_update()

    def _on_theme_changed(self, mode: str):
        """Apply and persist the user theme selection."""
        try:
            self._apply_theme_mode(mode, reason="user")
        except Exception as e:
            logger.debug(f"Error changing theme: {e}")

    def _go_to_tab(self, tab_index: int):
        """Navigate to a specific tab index."""
        self.tabs.selected_index = tab_index
        self.workflow_manager.handle_tab_change(tab_index)
        self._safe_page_update()

    def _go_to_configure_tab(self):
        self._go_to_tab(2)

    def _go_to_search_select_tab(self):
        self._go_to_tab(1)

    def _display_search_results(self):
        """Displays the current search results in the results_area."""
        self.search_manager.display_search_results()

    def _enhanced_search_results_display(self):
        """Enhanced results delegate to SearchManager for consistent UI."""
        self.search_manager.display_search_results()

    def _show_test_details(self, test):
        """Show detailed information about a test in a dialog."""
        try:
            details_content = ft.Column([
                ft.Text(f"Test Details: {test.test_lab_number}", size=18, weight=ft.FontWeight.BOLD),
                ft.Divider(),
                ft.Text(f"SAP Code: {test.sap_code}"),
                ft.Text(f"Voltage: {test.voltage}"),
                ft.Text(f"Notes: {test.notes}"),
                ft.Divider(),
                ft.Text("Backend Status:", weight=ft.FontWeight.W_500),
            ], spacing=10)
            
            # Try to get backend information
            if hasattr(self, 'app') and self.app:
                try:
                    test_data = self.app._process_single_test(test.test_lab_number)
                    if test_data:
                        details_content.controls.extend([
                            ft.Text(
                                f"✓ INF file: {'Found' if test_data.inf_data else 'Not found'}",
                                color=self._themed_color('success', 'green') if test_data.inf_data else self._themed_color('error', 'red')
                            ),
                            ft.Text(
                                f"✓ CSV file: {'Found' if test_data.csv_data is not None else 'Not found'}",
                                color=self._themed_color('success', 'green') if test_data.csv_data is not None else self._themed_color('error', 'red')
                            ),
                            ft.Text(
                                f"✓ Noise data: {'Found' if test_data.noise_info else 'Not found'}",
                                color=self._themed_color('success', 'green') if test_data.noise_info else self._themed_color('warning', 'orange')
                            ),
                            ft.Text(f"Status: {test_data.status_message}")
                        ])
                except Exception as e:
                    details_content.controls.append(ft.Text(
                        f"Error checking backend status: {str(e)}",
                        color=self._themed_color('error', 'red')
                    ))
            
            dialog = ft.AlertDialog(
                title=ft.Text("Test Information"),
                content=ft.Container(
                    content=details_content,
                    width=400,
                    height=300
                ),
                actions=[
                    ft.TextButton("Close", on_click=lambda e: self._close_dialog())
                ]
            )
            
            # Use Flet's page.open method for dialogs
            self.page.open(dialog)
            
        except Exception as e:
            logger.error(f"Error showing test details: {e}")
            self.status_manager.update_status(
                f"Error showing test details: {str(e)}",
                self._themed_color('error', 'red')
            )

    def _close_dialog(self):
        """Close the current dialog using Flet's page.close method."""
        try:
            # Close any open dialogs using Flet's built-in method
            if hasattr(self.page, 'overlay') and self.page.overlay:
                # Find and close any open dialogs in overlay
                for control in self.page.overlay:
                    if hasattr(control, 'open') and getattr(control, 'open', False):
                        self.page.close(control)
        except Exception as e:
            logger.debug(f"Error closing dialog: {e}")

    def _update_config_summary(self):
        """Updates the summary on the configuration tab."""
        # Configuration tab updates are handled through workflow manager
        if self._has_workflow_manager():
            self.workflow_manager.refresh_tab('config')
        else:
            logger.debug("Config tab refresh handled through standard page refresh")

    def _initialize_backend(self):
        """Initializes the backend MotorReportApp."""
        self.event_handlers._initialize_backend()

    def _update_backend_config(self):
        """Updates the backend configuration with current paths."""
        if not self.app: return
        
        self.app.config.tests_folder = self.state_manager.state.selected_tests_folder
        self.app.config.registry_path = self.state_manager.state.selected_registry_file
        self.app.config.noise_dir = self.state_manager.state.selected_noise_folder
        self.app.config.noise_registry_path = self.state_manager.state.selected_noise_registry
        
        # Ensure logo path is set for Excel writer
        if not self.app.config.logo_path:
            from ..config.directory_config import LOGO_PATH
            if LOGO_PATH and LOGO_PATH.exists():
                self.app.config.logo_path = str(LOGO_PATH)
        
        logger.info("Backend config updated with new paths.")

    def _apply_search_selection(self, e=None):
        """Apply the current search selection - wrapper for event handler."""
        if self.event_handlers:
            self.event_handlers.on_apply_search_selection(e)

    def _clear_search_selection(self, e=None):
        """Clear the current search selection - wrapper for event handler."""
        if self.event_handlers:
            self.event_handlers.on_clear_search_selection(e)

    @property
    def search_selection_applied(self) -> bool:
        """Check if search selection has been applied."""
        return getattr(self.state_manager.state, 'search_selection_applied', False) if self.state_manager else False

    @property
    def selected_tests(self) -> dict:
        """Get selected tests from state manager."""
        return getattr(self.state_manager.state, 'selected_tests', {}) if self.state_manager else {}

    @property
    def selected_performance_saps(self) -> list:
        """Get selected performance SAP codes."""
        return getattr(self.state_manager.state, 'selected_performance_saps', []) if self.state_manager else []

    @property
    def selected_noise_saps(self) -> list:
        """Get selected noise SAP codes."""
        return getattr(self.state_manager.state, 'selected_noise_saps', []) if self.state_manager else []

    @property
    def selected_comparison_saps(self) -> list:
        """Get selected comparison SAP codes."""
        return getattr(self.state_manager.state, 'selected_comparison_saps', []) if self.state_manager else []

    def _on_performance_sap_checked(self, e):
        """Handle performance SAP checkbox changes."""
        sap_code = e.control.data
        selected = e.control.value
        if self.state_manager:
            self.state_manager.update_sap_selection('performance', sap_code, selected)

    def _on_noise_sap_checked(self, e):
        """Handle noise SAP checkbox changes."""
        sap_code = e.control.data
        selected = e.control.value
        if self.state_manager:
            self.state_manager.update_sap_selection('noise', sap_code, selected)

    def _on_comparison_sap_checked(self, e):
        """Handle comparison SAP checkbox changes."""
        sap_code = e.control.data
        selected = e.control.value
        if self.state_manager:
            self.state_manager.update_sap_selection('comparison', sap_code, selected)

    def _noise_data_exists(self, sap_code: str) -> bool:
        """Check if noise data exists for a SAP code."""
        # Simple implementation - can be enhanced
        return True  # For now, assume noise data exists

    def _disable_noise_feature(self):
        """Disable noise feature for all SAPs."""
        if self.state_manager:
            self.state_manager.state.selected_noise_saps.clear()
            self.status_manager.update_status("Noise feature disabled", "orange")

    def _disable_comparison_feature(self):
        """Disable comparison feature for all SAPs."""
        if self.state_manager:
            self.state_manager.state.selected_comparison_saps.clear()
            self.status_manager.update_status("Comparison feature disabled", "orange")

    def _apply_config_selection(self, e=None):
        """Apply configuration selection."""
        self.status_manager.update_status("Configuration applied", "green")
        self._safe_page_update()

    def _clear_config_selection(self, e=None):
        """Clear configuration selection."""
        if self.state_manager:
            self.state_manager.state.selected_performance_saps.clear()
            self.state_manager.state.selected_noise_saps.clear()
            self.state_manager.state.selected_comparison_saps.clear()
        self.status_manager.update_status("Configuration cleared", "orange")
        self._safe_page_update()

    def _go_to_generate_tab(self):
        """Navigate to generate tab."""
        self._go_to_tab(3)

    def _handle_disconnect(self, e=None):
        """Clean up resources when page disconnects."""
        logger.info("Page disconnecting, cleaning up resources...")
        try:
            # Shutdown thread pool gracefully
            shutdown_thread_pool(wait=False)  # Don't wait, exit quickly
            logger.info("Thread pool shutdown complete")
        except Exception as ex:
            logger.error(f"Error during thread pool cleanup: {ex}")
        
        # Call existing cleanup if any
        if hasattr(self, 'event_handlers') and hasattr(self.event_handlers, 'on_disconnect'):
            try:
                self.event_handlers.on_disconnect(e)
            except Exception as ex:
                logger.debug(f"Event handler cleanup error (expected during shutdown): {ex}")

    def _cleanup_and_exit(self):
        """Cleanup resources before exiting."""
        logger.info("GUI is disconnecting. Cleaning up...")
        # Add any cleanup logic here, e.g., saving state

    def refresh_page(self, force_rebuild=False):
        """
        Unified page refresh method with optional full rebuild.
        Replaces multiple redundant refresh methods.
        """
        try:
            if force_rebuild:
                logger.info("Performing full page rebuild...")
                current_tab_index = getattr(self.tabs, 'selected_index', 0) if hasattr(self, 'tabs') else 0
                
                # Reinitialize tab objects
                self.setup_tab = SetupTab(self)
                self.search_select_tab = SearchSelectTab(self)
                self.config_tab = ConfigTab(self)
                self.generate_tab = GenerateTab(self)
                
                # Rebuild layout
                self.build_layout()
                
                # Restore tab selection
                if self._has_tabs_with_selection():
                    self.tabs.selected_index = current_tab_index
                
                # Update workflow state
                if self._has_workflow_manager():
                    self.workflow_manager.update_workflow_state()
            else:
                logger.debug("Performing lightweight page refresh...")
            
            # Single page update
            self._safe_page_update()
            return True
            
        except Exception as e:
            logger.error(f"Error in page refresh: {e}")
            return False

    def _has_event_handlers(self):
        """Helper to check if event handlers are available"""
        return hasattr(self, 'event_handlers') and self.event_handlers

    def _has_workflow_manager(self):
        """Helper to check if workflow manager is available"""
        return hasattr(self, 'workflow_manager') and self.workflow_manager

    def _has_tabs_with_selection(self):
        """Helper to check if tabs exist and have selection capability"""
        return (hasattr(self, 'tabs') and self.tabs and 
                hasattr(self.tabs, 'selected_index'))

    def _has_valid_page_session(self):
        """Helper to check if page and session are valid"""
        return (self.page and 
                hasattr(self.page, 'session_id') and 
                self.page.session_id)

def main(page: ft.Page):
    """Main function to run the Flet application."""
    logger.info("Starting Motor Report Generator GUI")
    
    # Set project root for the application
    page.client_storage.set("project_root", str(PROJECT_ROOT))
    # Ensure data directories are initialized (per-user auto-discovery)
    try:
        ensure_directories_initialized()
    except Exception:
        logger.debug("ensure_directories_initialized() failed or deferred")

    # Ensure per-user directory cache file exists (create empty cache on first startup)
    try:
        from ..config.directory_cache import get_directory_cache
        cache = get_directory_cache()
        # Log effective cache path for troubleshooting and visibility
        try:
            logger.info(f"Directory cache path: {cache.cache_file}")
        except Exception:
            logger.debug("Could not read cache.cache_file for logging")

        # Ensure cache file exists on disk (uses atomic write)
        try:
            cache.ensure_exists()
        except Exception:
            logger.debug("Could not ensure per-user directory cache file exists")
    except Exception:
        logger.debug("Directory cache initialization skipped")

    gui = MotorReportAppGUI(page)
    page.update()

if __name__ == "__main__":
    ft.app(target=main)

