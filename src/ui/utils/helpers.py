"""
Progress and status management utilities
"""
import flet as ft
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)

ColorResolver = Optional[Callable[[str, str], Optional[str]]]


def _resolve_color(resolver: ColorResolver, token: str, fallback: str) -> str:
    """Best-effort wrapper that defers to a theme-aware color resolver."""

    if callable(resolver):
        try:
            resolved = resolver(token, fallback)
            if resolved:
                return resolved
        except Exception as exc:
            logger.debug("Color resolver failed for token '%s': %s", token, exc)
    return fallback


# NOTE: StatusManager has been moved to src/gui/core/status_manager.py
# for architectural consistency with other managers (StateManager, WorkflowManager, etc.)


# NOTE: WorkflowManager has been moved to src/gui/core/workflow_manager.py
# This duplicate class was removed to avoid confusion and maintain single source of truth


class SearchResultsFormatter:
    """Formats and displays search results with SAP code pagination"""
    
    @staticmethod
    def format_results(
        tests: list,
        on_test_selected: Callable,
        on_row_clicked: Callable,
        selected_tests: Optional[dict] = None,
        sap_navigator=None,
        color_resolver: ColorResolver = None,
    ) -> list:
        """Format tests with SAP code pagination - show one SAP at a time."""

        color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
        muted = color('text_muted', 'grey')
        if not tests:
            return [ft.Container(
                content=ft.Text("No results to display.", color=muted),
                alignment=ft.alignment.center,
                padding=20
            )]
        
        selected_tests = selected_tests or {}
        
        # Group tests by SAP code
        sap_groups = {}
        for test in tests:
            sap_code = test.sap_code or "Unknown SAP"
            if sap_code not in sap_groups:
                sap_groups[sap_code] = []
            sap_groups[sap_code].append(test)
        
        logger.info(f"SearchResultsFormatter: Grouped {len(tests)} tests into {len(sap_groups)} SAP groups")
        
        # Update SAP navigator with current groups
        if sap_navigator:
            sap_navigator.update_sap_codes(tests)
            current_sap_tests = sap_navigator.get_current_sap_tests(tests)
        else:
            # Fallback: show first SAP group if no navigator
            current_sap_tests = list(sap_groups.values())[0] if sap_groups else []
        
        if not current_sap_tests:
            return [ft.Container(
                content=ft.Text("No tests for current SAP selection.", color=muted),
                alignment=ft.alignment.center,
                padding=20
            )]
        
        controls = []
        
        # Get current SAP code
        current_sap = current_sap_tests[0].sap_code if current_sap_tests else "Unknown SAP"
        
        # SAP Code Header with summary and navigation info
        primary = color('primary', 'blue')
        sap_header = ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.INVENTORY, size=16, color=primary),
                ft.Text(f"SAP Code: {current_sap}", weight=ft.FontWeight.BOLD, size=14, color=primary),
                ft.Text(
                    f"({len(current_sap_tests)} test{'s' if len(current_sap_tests) != 1 else ''})",
                    size=12,
                    color=muted
                ),
                ft.Container(expand=True),  # Spacer
                ft.IconButton(
                    icon=ft.Icons.SELECT_ALL,
                    tooltip=f"Select all {len(current_sap_tests)} tests",
                    icon_size=16,
                    on_click=lambda e, tests=current_sap_tests: SearchResultsFormatter._select_all_tests(e, tests, on_test_selected)
                ),
            ], spacing=5),
            bgcolor=color('primary_container', '#e8f4f8'),
            padding=ft.padding.all(12),
            border_radius=8,
            margin=ft.margin.only(top=10, bottom=5),
            border=ft.border.all(1, color('outline', '#d0d0d0'))
        )
        controls.append(sap_header)
        
        # Add column headers with Date prominently displayed
        header_text_color = color('on_surface', 'darkblue')
        column_header = ft.Container(
            content=ft.Row([
                ft.Container(width=40),  # Space for checkbox
                ft.Container(
                    content=ft.Text("Test Lab", size=11, weight=ft.FontWeight.W_500, color=header_text_color),
                    width=100
                ),
                ft.Container(
                    content=ft.Text("Date", size=11, weight=ft.FontWeight.W_500, color=header_text_color),
                    width=120  # Increased width for date
                ),
                ft.Container(
                    content=ft.Text("Voltage", size=11, weight=ft.FontWeight.W_500, color=header_text_color),
                    width=80
                ),
                ft.Container(
                    content=ft.Text("Notes", size=11, weight=ft.FontWeight.W_500, color=header_text_color),
                    expand=True,
                    padding=ft.padding.only(right=10)
                ),
            ], alignment=ft.MainAxisAlignment.START),
            bgcolor=color('surface_variant', '#f0f7ff'),
            padding=ft.padding.symmetric(horizontal=15, vertical=5),
            border_radius=3,
            border=ft.border.all(1, color('outline', '#d0d0d0'))
        )
        controls.append(column_header)
        
        # Show tests for current SAP only
        test_rows = SearchResultsFormatter._create_test_rows(
            current_sap_tests,
            on_test_selected,
            on_row_clicked,
            selected_tests,
            color_resolver=color_resolver,
        )
        
        # Create container for current SAP group
        test_container = ft.Container(
            content=ft.Column(test_rows, spacing=2),
            margin=ft.margin.only(left=10, right=10, bottom=10),
            padding=ft.padding.all(8),
            bgcolor=color('surface_variant', '#fafafa'),
            border_radius=5,
            border=ft.border.all(1, color('outline', '#e0e0e0'))
        )
        controls.append(test_container)
        
        return controls
    
    @staticmethod
    def _select_all_tests(event, tests: list, on_test_selected: Callable):
        """Select all tests in a SAP code group."""
        logger.info(f"Selecting all {len(tests)} tests in group")
        for test in tests:
            # Create a mock event object for the checkbox
            class MockControl:
                def __init__(self):
                    self.value = True
                    self.data = test
            
            class MockEvent:
                def __init__(self):
                    self.control = MockControl()
            
            # Simulate checkbox change event for each test
            mock_event = MockEvent()
            on_test_selected(mock_event)

    @staticmethod
    def _create_test_rows(
        tests: list,
        on_test_selected: Callable,
        on_row_clicked: Callable,
        selected_tests: dict,
        *,
        color_resolver: ColorResolver = None,
    ) -> list:
        """Create the actual test row controls with enhanced date display."""
        color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
        base_bg = color('surface', '#ffffff')
        alt_bg = color('surface_variant', '#f8f8f8')
        border_color = color('outline', '#eeeeee')
        date_success = color('success', 'darkgreen')
        date_muted = color('text_muted', 'grey')
        rows = []
        for i, test in enumerate(tests):
            bg_color = base_bg if i % 2 == 0 else alt_bg
            checked = test.test_lab_number in selected_tests
            row_click_handler = on_row_clicked(test)
            if row_click_handler is None:
                row_click_handler = lambda e: None
            
            # Format date with better display - prioritize from test data
            test_date = "N/A"
            if hasattr(test, 'date') and test.date:
                test_date = str(test.date)
            elif hasattr(test, 'test_date') and test.test_date:
                test_date = str(test.test_date)
            
            # Ensure date formatting looks good
            if test_date != "N/A" and len(test_date) > 10:
                # Truncate very long dates and add ellipsis
                test_date = test_date[:10] + "..."
            
            test_row = ft.Container(
                content=ft.Row([
                    ft.Checkbox(
                        value=checked,
                        data=test,
                        on_change=on_test_selected,
                        width=40
                    ),
                    ft.Container(
                        content=ft.Text(test.test_lab_number, size=12, weight=ft.FontWeight.W_500),
                        width=100
                    ),
                    ft.Container(
                        content=ft.Text(
                            test_date,
                            size=12,
                            color=date_success if test_date != "N/A" else date_muted,
                            weight=ft.FontWeight.W_500 if test_date != "N/A" else ft.FontWeight.NORMAL
                        ),
                        width=120  # Increased width for date display
                    ),
                    ft.Container(
                        content=ft.Text(f"{test.voltage}V" if test.voltage else "N/A", size=12),
                        width=80
                    ),
                    ft.Container(
                        content=ft.Text(
                            test.notes or "No notes",
                            size=11,
                            selectable=True,
                            tooltip=test.notes if test.notes else "No notes available",
                            max_lines=2,
                            overflow=ft.TextOverflow.ELLIPSIS
                        ),
                        expand=True,
                        padding=ft.padding.only(right=10)
                    ),
                ], alignment=ft.MainAxisAlignment.START),
                bgcolor=bg_color,
                padding=ft.padding.symmetric(horizontal=15, vertical=8),
                border_radius=3,
                margin=ft.margin.only(bottom=1),
                on_click=row_click_handler,
                border=ft.border.all(1, border_color)
            )
            rows.append(test_row)
        
        return rows


class SAPNavigationManager:
    """Manages navigation between different SAP codes in search results"""
    
    def __init__(self, gui_instance):
        self.gui = gui_instance
        self.current_sap_index = 0
        self.sap_codes = []
        
    def update_sap_codes(self, tests: list):
        """Update the list of available SAP codes from test results"""
        sap_set = set()
        for test in tests:
            sap_code = test.sap_code or "Unknown SAP"
            sap_set.add(sap_code)
        new_sap_codes = sorted(list(sap_set))
        
        # Reset index whenever the set of SAP codes changes
        if new_sap_codes != self.sap_codes:
            self.sap_codes = new_sap_codes
            self.current_sap_index = 0
        else:
            # SAP codes are the same, keep current index but ensure it is in range
            self.sap_codes = new_sap_codes
            if self.current_sap_index >= len(self.sap_codes):
                self.current_sap_index = 0
    
    def get_current_sap_tests(self, all_tests: list) -> list:
        """Get tests for the currently selected SAP code"""
        if not self.sap_codes or self.current_sap_index >= len(self.sap_codes):
            return all_tests
            
        current_sap = self.sap_codes[self.current_sap_index]
        return [test for test in all_tests if (test.sap_code or "Unknown SAP") == current_sap]
    
    def create_navigation_controls(self, update_callback: Callable) -> ft.Container:
        """Create SAP navigation controls with enhanced design"""
        color = self._color
        current_label = (
            f"Viewing SAP Code {self.current_sap_index + 1} of {len(self.sap_codes)}"
            if self.sap_codes else "No SAP codes available"
        )
        current_value = self.sap_codes[self.current_sap_index] if self.sap_codes else "None"
        prev_disabled = self.current_sap_index <= 0
        next_disabled = self.current_sap_index >= len(self.sap_codes) - 1

        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Icon(ft.Icons.NAVIGATE_BEFORE, color=color('primary', 'blue'), size=20),
                    ft.Text(current_label, size=14, weight=ft.FontWeight.W_500, color=color('primary', 'blue')),
                    ft.Icon(ft.Icons.NAVIGATE_NEXT, color=color('primary', 'blue'), size=20),
                ], alignment=ft.MainAxisAlignment.CENTER, spacing=5),
                ft.Row([
                    ft.Text(
                        f"Current: {current_value}",
                        size=16,
                        weight=ft.FontWeight.BOLD,
                        color=color('on_surface', 'darkblue')
                    ),
                ], alignment=ft.MainAxisAlignment.CENTER),
                ft.Row([
                    ft.ElevatedButton(
                        "◀ Previous SAP",
                        disabled=prev_disabled,
                        on_click=lambda e: self._navigate_sap(-1, update_callback),
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=8)
                        ),
                        height=40
                    ),
                    ft.Container(width=20),  # Spacer
                    ft.ElevatedButton(
                        "Next SAP ▶",
                        disabled=next_disabled,
                        on_click=lambda e: self._navigate_sap(1, update_callback),
                        style=ft.ButtonStyle(
                            shape=ft.RoundedRectangleBorder(radius=8)
                        ),
                        height=40
                    )
                ], alignment=ft.MainAxisAlignment.CENTER, spacing=10)
            ], spacing=8),
            bgcolor=color('primary_container', '#e3f2fd') if self.sap_codes else color('surface_variant', '#f0f0f0'),
            padding=ft.padding.all(15),
            border_radius=8,
            margin=ft.margin.only(bottom=15),
            border=ft.border.all(2, color('primary', '#1976d2') if self.sap_codes else color('outline', '#cccccc'))
        )
    
    def _navigate_sap(self, delta: int, update_callback: Callable):
        """Navigate to next/previous SAP code"""
        new_index = self.current_sap_index + delta
        if 0 <= new_index < len(self.sap_codes):
            self.current_sap_index = new_index
            logger.info(f"Navigated to SAP {self.current_sap_index + 1} of {len(self.sap_codes)}: {self.sap_codes[self.current_sap_index]}")
            try:
                update_callback()
            except Exception as e:
                logger.error(f"Error in SAP navigation update callback: {e}")
        else:
            logger.warning(f"Cannot navigate to SAP index {new_index}, valid range is 0 to {len(self.sap_codes) - 1}")
        
    def reset(self):
        """Reset navigation to first SAP code"""
        self.current_sap_index = 0

    def _color(self, token: str, fallback: str) -> str:
        """Convenience wrapper that defers to the GUI's theme resolver if present."""

        resolver = getattr(self.gui, "_themed_color", None)
        if callable(resolver):
            try:
                return resolver(token, fallback) or fallback
            except Exception as exc:
                logger.debug("SAPNavigationManager color fallback (%s): %s", token, exc)
        return fallback

