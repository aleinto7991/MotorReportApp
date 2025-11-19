"""
Optimized Search functionality for the Motor Report GUI
Handles search operations and results display with performance optimizations.
"""
import flet as ft
import logging
import threading
from typing import List, Dict, Optional, TYPE_CHECKING, Any, Callable
from ...data.models import Test
from ..utils.pagination import Paginator

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI

logger = logging.getLogger(__name__)


class SearchManager:
    """Manages search functionality and result display with performance optimizations"""
    
    def __init__(self, gui: 'MotorReportAppGUI'):
        self.gui = gui
        self._results_builder = SearchResultsBuilder(gui)
        self.filters: Dict[str, str] = {
            "test_lab": "",
            "date": "",
            "voltage": "",
            "notes": ""
        }
        self.filter_inputs: Dict[str, Any] = {}
        self._filter_refresh_timer: Optional[threading.Timer] = None
        self._refresh_lock = threading.Lock()
    
    @property
    def state(self):
        """Get state manager from GUI"""
        return self.gui.state_manager
    
    def display_search_results(self):
        """Display search results with optimized performance and SAP navigation"""
        try:
            all_tests = list(self.state.state.found_tests)
            if not all_tests:
                self._display_empty_results()
                self._hide_navigation()
                return

            filtered_tests = self._apply_filters(all_tests)
            filters_active = any(value for value in self.filters.values())

            self._results_builder.render(
                visible_tests=filtered_tests,
                all_tests=all_tests,
                filters_active=filters_active
            )
            
        except Exception as e:
            logger.error(f"Error displaying search results: {e}")
            self._display_error_fallback()
    
    def _update_navigation(self):
        """Update SAP navigation controls"""
        try:
            if hasattr(self.gui, 'sap_navigator') and self.gui.sap_navigator:
                # Create navigation controls with callback to refresh results
                navigation_controls = self.gui.sap_navigator.create_navigation_controls(
                    update_callback=self._refresh_current_sap_display
                )
                
                # Update the navigation container
                if hasattr(self.gui, 'sap_navigation_container'):
                    self.gui.sap_navigation_container.content = navigation_controls
                    self.gui.sap_navigation_container.visible = len(self.gui.sap_navigator.sap_codes) >= 1
                    
        except Exception as e:
            logger.error(f"Error updating navigation: {e}")
    
    def _hide_navigation(self):
        """Hide SAP navigation when not needed"""
        try:
            if hasattr(self.gui, 'sap_navigation_container'):
                self.gui.sap_navigation_container.visible = False
        except Exception as e:
            logger.error(f"Error hiding navigation: {e}")
    
    def _refresh_current_sap_display(self):
        """Refresh display for current SAP selection"""
        try:
            # Re-display results with new SAP selection
            self.display_search_results()
        except Exception as e:
            logger.error(f"Error refreshing SAP display: {e}")
    
    def _display_empty_results(self):
        """Display empty state efficiently"""
        self.gui.results_area.controls.clear()
        self.gui.results_area.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.SEARCH_OFF, size=48, color=self.gui._themed_color('on_surface', 'grey')),
                    ft.Text("No search results", size=16, weight=ft.FontWeight.W_500, color=self.gui._themed_color('on_surface', 'grey')),
                    ft.Text("Enter a SAP code or test number to search", size=12, color=self.gui._themed_color('on_surface', 'grey'))
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=ft.padding.all(20),
                alignment=ft.alignment.center
            )
        )
        self.gui._safe_page_update()
    
    def _display_error_fallback(self):
        """Display error state"""
        self.gui.results_area.controls.clear()
        self.gui.results_area.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.ERROR, size=48, color=self.gui._themed_color('error', 'red')),
                    ft.Text("Error loading results", size=16, weight=ft.FontWeight.W_500, color=self.gui._themed_color('error', 'red')),
                    ft.Text("Please try searching again", size=12, color=self.gui._themed_color('on_surface', 'grey'))
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=ft.padding.all(20),
                alignment=ft.alignment.center
            )
        )
        self.gui._safe_page_update()
    
    def _update_results_ui(self):
        """Update UI elements efficiently after displaying results"""
        # Update selection count
        if hasattr(self.gui, 'selected_count_text'):
            count = len(self.state.state.selected_tests)
            self.gui.selected_count_text.value = f"Selected: {count} tests"
            self.gui.selected_count_text.visible = count > 0
        
        # Update status with SAP information
        if hasattr(self.gui, 'status_manager'):
            count = len(self.state.state.selected_tests)
            total = len(self.state.state.found_tests)
            
            # Include SAP navigation info in status
            sap_info = ""
            if hasattr(self.gui, 'sap_navigator') and len(self.gui.sap_navigator.sap_codes) > 1:
                current_sap = self.gui.sap_navigator.sap_codes[self.gui.sap_navigator.current_sap_index]
                current_sap_tests = len(self.gui.sap_navigator.get_current_sap_tests(self.state.state.found_tests))
                sap_info = f" (Viewing SAP {current_sap}: {current_sap_tests} tests)"
            
            if count > 0:
                self.gui.status_manager.update_status(
                    f"✅ {count} of {total} tests selected{sap_info}. Click 'Configure' tab to continue.", 
                    "green"
                )
            else:
                self.gui.status_manager.update_status(
                    f"Found {total} tests{sap_info}. Select the ones to include in your report.", 
                    "blue"
                )
        
        # Single page update
        self.gui._safe_page_update()
    
    def search_tests(self, search_term: str):
        """Search for tests - delegated to backend"""
        try:
            if hasattr(self.gui, 'app') and self.gui.app:
                found_tests = self.gui.app.search_tests(search_term)
                self.state.state.found_tests = found_tests
                
                # Extract unique SAP codes
                sap_codes = list(set(test.sap_code for test in found_tests if test.sap_code))
                self.state.state.found_sap_codes = sap_codes
                
                # Reset SAP navigator to first SAP when new search is performed
                if hasattr(self.gui, 'sap_navigator') and self.gui.sap_navigator:
                    self.gui.sap_navigator.reset()
                
                logger.info(f"Search completed: {len(found_tests)} tests found, {len(sap_codes)} unique SAP codes")
                self.display_search_results()
                return found_tests
            else:
                logger.warning("Backend app not available for search")
                return []
        except Exception as e:
            logger.error(f"Error during search: {e}")
            return []

    # --- filter management -------------------------------------------------

    def register_filter_inputs(self, filter_inputs: Dict[str, Any]):
        """Track filter input controls so they can be reset programmatically."""
        self.filter_inputs = filter_inputs

    def generate_filter_handler(self, filter_key: str):
        """Create an on_change handler that updates a specific filter field."""
        def handler(event):
            self.filters[filter_key] = (event.control.value or "").strip()
            logger.debug(f"Filter '{filter_key}' updated to '{self.filters[filter_key]}'")
            self._schedule_results_refresh()
        return handler

    def clear_filters(self, event=None):
        """Clear all active filters and refresh the search results."""
        any_cleared = any(value for value in self.filters.values())
        for key in self.filters:
            self.filters[key] = ""
        for control in self.filter_inputs.values():
            try:
                control.value = ""
                control.update()
            except Exception as update_err:
                logger.debug(f"Filter control reset failed: {update_err}")
        if any_cleared:
            self._schedule_results_refresh(delay=0.05)

    def _apply_filters(self, tests: List[Test]) -> List[Test]:
        if not any(self.filters.values()):
            return tests
        filtered = [test for test in tests if self._matches_filters(test)]
        logger.info(f"Filter applied: {len(filtered)} of {len(tests)} tests match current filters")
        return filtered

    def _matches_filters(self, test: Test) -> bool:
        def contains(value: Optional[str], pattern: str) -> bool:
            return pattern.lower() in (value or "").lower()

        if self.filters["test_lab"] and not contains(getattr(test, "test_lab_number", ""), self.filters["test_lab"]):
            return False
        if self.filters["date"]:
            date_value = getattr(test, "date", None) or getattr(test, "test_date", None) or ""
            if not contains(str(date_value), self.filters["date"]):
                return False
        if self.filters["voltage"] and not contains(str(getattr(test, "voltage", "")), self.filters["voltage"]):
            return False
        if self.filters["notes"] and not contains(getattr(test, "notes", ""), self.filters["notes"]):
            return False
        return True

    def _schedule_results_refresh(self, delay: float = 0.25):
        """Debounce heavy UI refreshes when filters change."""
        with self._refresh_lock:
            if self._filter_refresh_timer:
                self._filter_refresh_timer.cancel()
            self._filter_refresh_timer = threading.Timer(delay, self._trigger_refresh)
            self._filter_refresh_timer.daemon = True
            self._filter_refresh_timer.start()

    def _trigger_refresh(self):
        run_thread = getattr(getattr(self.gui, 'page', None), 'run_thread', None)
        if callable(run_thread):
            run_thread(self.display_search_results)
        else:
            self.display_search_results()


class SearchResultsBuilder:
    """Builds and renders search results UI within the Search & Select tab."""

    def __init__(self, gui: 'MotorReportAppGUI'):
        self.gui = gui
        self.paginator = Paginator(items=[], page_size=50)  # 50 items per page for optimal performance

    def render(
        self,
        visible_tests: List[Test],
        all_tests: List[Test],
        filters_active: bool = False
    ):  # noqa: D401 - documentation inherited
        total_tests = len(all_tests)
        visible_total = len(visible_tests)
        logger.info(
            "SearchResultsBuilder: rendering %s visible tests (total=%s, filters=%s)",
            visible_total,
            total_tests,
            filters_active
        )

        if not all_tests:
            self._render_empty()
            return

        gui = self.gui
        gui.results_area.controls.clear()

        navigator = getattr(gui, "sap_navigator", None)
        current_sap = None

        if navigator:
            navigator.update_sap_codes(all_tests)
            if navigator.sap_codes:
                try:
                    current_sap = navigator.sap_codes[navigator.current_sap_index]
                except IndexError:
                    navigator.current_sap_index = 0
                    current_sap = navigator.sap_codes[0] if navigator.sap_codes else None

        if not current_sap:
            sample_list = visible_tests or all_tests
            if sample_list:
                current_sap = sample_list[0].sap_code or "Unknown SAP"
            else:
                current_sap = "Unknown SAP"

        def sap_key(test: Test) -> str:
            return (test.sap_code or "Unknown SAP")

        current_visible_tests = [test for test in visible_tests if sap_key(test) == current_sap]
        current_all_tests = [test for test in all_tests if sap_key(test) == current_sap]
        sap_total = len(current_all_tests)
        sap_visible = len(current_visible_tests) if filters_active else sap_total

        if not current_visible_tests and filters_active:
            self._render_filtered_empty(total_tests, current_sap)
            self._update_navigation(all_tests)
            return

        display_tests = current_visible_tests if current_visible_tests else current_all_tests

        if not display_tests:
            gui.results_area.controls.append(
                ft.Container(
                    content=ft.Text(
                        "No tests available for the current SAP selection.",
                        color=self._color('text_muted', 'grey')
                    ),
                    alignment=ft.alignment.center,
                    padding=20
                )
            )
            gui._safe_page_update()
            self._update_navigation(all_tests)
            return

        # Update paginator with current display tests
        self.paginator.items = display_tests
        
        gui.results_area.controls.append(
            self._build_header(
                total_tests,
                visible_total,
                current_sap,
                sap_total,
                sap_visible,
                filters_active
            )
        )
        
        # Add pagination controls at top if needed (more than one page)
        if self.paginator.total_pages > 1:
            gui.results_area.controls.append(
                ft.Container(
                    content=self.paginator.create_navigation_controls(
                        on_page_change=lambda: self._refresh_page_display(
                            total_tests, visible_total, current_sap, sap_total, sap_visible, filters_active, all_tests
                        )
                    ),
                    padding=ft.padding.symmetric(vertical=10),
                    bgcolor=self._color('surface_variant', '#fafafa'),
                    border_radius=8,
                    border=ft.border.all(1, self._color('outline', '#e0e0e0'))
                )
            )
        
        gui.results_area.controls.append(self._build_column_headers())

        # Render only current page of tests (50 at a time instead of all)
        page_tests = self.paginator.get_current_page()
        for index, test in enumerate(page_tests):
            # Use global index for alternating colors across pages
            global_index = self.paginator.current_page * self.paginator.page_size + index
            gui.results_area.controls.append(self._build_test_row(test, global_index))
        
        # Add pagination controls at bottom if needed
        if self.paginator.total_pages > 1:
            gui.results_area.controls.append(
                ft.Container(
                    content=self.paginator.create_navigation_controls(
                        on_page_change=lambda: self._refresh_page_display(
                            total_tests, visible_total, current_sap, sap_total, sap_visible, filters_active, all_tests
                        )
                    ),
                    padding=ft.padding.symmetric(vertical=10),
                    bgcolor=self._color('surface_variant', '#fafafa'),
                    border_radius=8,
                    border=ft.border.all(1, self._color('outline', '#e0e0e0')),
                    margin=ft.margin.only(top=10)
                )
            )

        self._update_selection_indicator()
        self._update_status(
            total_tests,
            visible_total,
            current_sap,
            sap_total,
            sap_visible,
            filters_active
        )
        self._update_navigation(all_tests)

        try:
            gui.results_area.update()
        except Exception as update_err:
            logger.debug(f"SearchResultsBuilder: results_area update failed - {update_err}")

        gui._safe_page_update()

    # --- helpers ---------------------------------------------------------
    
    def _refresh_page_display(
        self,
        total_tests: int,
        visible_total: int,
        current_sap: str,
        sap_total: int,
        sap_visible: int,
        filters_active: bool,
        all_tests: List[Test]
    ):
        """
        Refresh display when pagination changes.
        Only re-renders the test rows, not the entire results area.
        """
        gui = self.gui
        
        # Find the index where test rows start (after header, pagination, column headers)
        # Structure: [header, top_pagination?, column_headers, rows..., bottom_pagination?]
        header_controls_count = 2  # header + column headers
        if self.paginator.total_pages > 1:
            header_controls_count += 1  # top pagination
        
        # Remove old test rows and bottom pagination (if exists)
        # Keep header, top pagination, and column headers
        controls_to_keep = header_controls_count
        gui.results_area.controls = gui.results_area.controls[:controls_to_keep]
        
        # Render current page of tests
        page_tests = self.paginator.get_current_page()
        for index, test in enumerate(page_tests):
            global_index = self.paginator.current_page * self.paginator.page_size + index
            gui.results_area.controls.append(self._build_test_row(test, global_index))
        
        # Add bottom pagination if needed
        if self.paginator.total_pages > 1:
            gui.results_area.controls.append(
                ft.Container(
                    content=self.paginator.create_navigation_controls(
                        on_page_change=lambda: self._refresh_page_display(
                            total_tests, visible_total, current_sap, sap_total, sap_visible, filters_active, all_tests
                        )
                    ),
                    padding=ft.padding.symmetric(vertical=10),
                    bgcolor=self._color('surface_variant', '#fafafa'),
                    border_radius=8,
                    border=ft.border.all(1, self._color('outline', '#e0e0e0')),
                    margin=ft.margin.only(top=10)
                )
            )
        
        # Update status and navigation
        self._update_selection_indicator()
        self._update_status(total_tests, visible_total, current_sap, sap_total, sap_visible, filters_active)
        self._update_navigation(all_tests)
        
        # Update UI
        try:
            gui.results_area.update()
        except Exception as update_err:
            logger.debug(f"Page refresh update failed: {update_err}")
        
        gui._safe_page_update()

    def _render_empty(self):
        gui = self.gui
        gui.results_area.controls.clear()
        gui.results_area.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.SEARCH_OFF, size=48, color=gui._themed_color('on_surface', 'grey')),
                    ft.Text("No search results", size=16, weight=ft.FontWeight.W_500, color=gui._themed_color('on_surface', 'grey')),
                    ft.Text("Enter a SAP code or test number to search", size=12, color=gui._themed_color('on_surface', 'grey'))
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=ft.padding.all(20),
                alignment=ft.alignment.center
            )
        )
        self._update_selection_indicator()
        self._update_status(0, 0, "", 0, 0, False)
        self._update_navigation([])
        gui._safe_page_update()

    def _render_filtered_empty(self, total_tests: int, current_sap: Optional[str] = None):
        gui = self.gui
        gui.results_area.controls.clear()
        sap_note = f" for SAP {current_sap}" if current_sap else ""
        gui.results_area.controls.append(
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.Icons.FILTER_ALT_OFF, size=48, color=gui._themed_color('on_surface', 'grey')),
                    ft.Text(
                        f"No tests{sap_note} match the current filters",
                        size=16,
                        weight=ft.FontWeight.W_500,
                        color=gui._themed_color('on_surface', 'grey')
                    ),
                    ft.Text(f"Clear filters to view all {total_tests} test(s).", size=12, color=gui._themed_color('on_surface', 'grey'))
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=ft.padding.all(20),
                alignment=ft.alignment.center
            )
        )
        self._update_selection_indicator()
        self._update_status(total_tests, 0, current_sap or "", 0, 0, True)
        gui._safe_page_update()

    def _build_header(
        self,
        total_tests: int,
        visible_tests: int,
        current_sap: str,
        sap_total: int,
        sap_visible: int,
        filters_active: bool
    ) -> ft.Control:
        gui = self.gui
        selected_count = len(getattr(gui.state_manager.state, 'selected_tests', {}))
        filter_note = "" if total_tests == visible_tests else f" | Showing {visible_tests} after filters"
        sap_note = (
            f"Viewing SAP {current_sap} ({sap_visible} of {sap_total} test(s))"
            if filters_active and sap_total != sap_visible
            else f"Viewing SAP {current_sap} ({sap_total} test(s))"
        )
        return ft.Container(
            content=ft.Row([
                ft.Icon(ft.Icons.LIST_ALT, color=self._color('primary', 'blue')),
                ft.Text(f"Found {total_tests} test(s){filter_note}", size=16, weight=ft.FontWeight.W_500),
                ft.Text(sap_note, size=12, color=self._color('text_muted', 'grey')),
                ft.Container(expand=True),
                ft.Text(
                    f"Selected: {selected_count}",
                    size=12,
                    color=(self._color('primary', 'green') if selected_count > 0 else self._color('text_muted', 'grey'))
                )
            ], spacing=10),
            padding=ft.padding.symmetric(vertical=6, horizontal=10),
            bgcolor=self._color('surface', '#eef5ff'),
            border_radius=6,
            margin=ft.margin.only(bottom=8)
        )

    def _build_column_headers(self) -> ft.Control:
        header_color = self._color('on_surface_variant', '#0f172a')
        return ft.Container(
            content=ft.Row([
                ft.Container(width=40),
                ft.Container(ft.Text("Test Lab", size=11, weight=ft.FontWeight.W_500, color=header_color), width=110),
                ft.Container(ft.Text("Date", size=11, weight=ft.FontWeight.W_500, color=header_color), width=130),
                ft.Container(ft.Text("Voltage", size=11, weight=ft.FontWeight.W_500, color=header_color), width=90),
                ft.Container(ft.Text("Notes", size=11, weight=ft.FontWeight.W_500, color=header_color), expand=True),
                ft.Container(width=40)
            ]),
            bgcolor=self._color('surface_variant', '#f5f9ff'),
            padding=ft.padding.symmetric(horizontal=12, vertical=6),
            border_radius=6,
            margin=ft.margin.only(bottom=4),
            border=ft.border.all(1, self._color('outline', '#d0d7e5'))
        )

    def _build_test_row(self, test, index: int) -> ft.Control:
        base_bg = self._color('surface', '#ffffff')
        alt_bg = self._color('surface_variant', '#f8f8f8')
        bg_color = base_bg if index % 2 == 0 else alt_bg
        border_color = self._color('outline', '#e0e0e0')
        selected_border = self._color('primary', '#1d4ed8')
        lab_color = self._color('on_surface', '#111827')
        date_success = self._color('success', 'darkgreen')
        date_muted = self._color('text_muted', 'grey')
        meta_color = self._color('text_muted', '#475467')
        selected_tests = getattr(self.gui.state_manager.state, 'selected_tests', {})
        is_selected = test.test_lab_number in selected_tests

        date_text = getattr(test, 'date', None) or getattr(test, 'test_date', None) or "N/A"
        if isinstance(date_text, str) and len(date_text) > 16:
            date_text = date_text[:16] + "…"

        row_click_handler = self.gui.event_handlers.on_row_clicked(test) if hasattr(self.gui, 'event_handlers') else None

        return ft.Container(
            content=ft.Row([
                ft.Checkbox(
                    value=is_selected,
                    data=test,
                    on_change=self.gui.event_handlers.on_test_selected if hasattr(self.gui, 'event_handlers') else None
                ),
                ft.Container(
                    ft.Text(test.test_lab_number, size=12, weight=ft.FontWeight.W_500, color=lab_color),
                    width=110
                ),
                ft.Container(
                    ft.Text(
                        date_text,
                        size=12,
                        color=date_success if date_text != "N/A" else date_muted,
                        weight=ft.FontWeight.W_500 if date_text != "N/A" else ft.FontWeight.NORMAL
                    ),
                    width=130
                ),
                ft.Container(
                    ft.Text(
                        f"{test.voltage}V" if getattr(test, 'voltage', None) else "N/A",
                        size=12,
                        color=lab_color
                    ),
                    width=90
                ),
                ft.Container(
                    ft.Text(
                        (test.notes or "No notes"),
                        size=11,
                        color=meta_color,
                        tooltip=test.notes or "No notes",
                        max_lines=2,
                        overflow=ft.TextOverflow.ELLIPSIS
                    ),
                    expand=True
                ),
                ft.IconButton(
                    icon=ft.Icons.INFO_OUTLINE,
                    tooltip="Test details",
                    icon_color=self._color('primary', '#2563eb'),
                    on_click=(lambda e, t=test: self.gui._show_test_details(t)) if hasattr(self.gui, '_show_test_details') else None
                )
            ], alignment=ft.MainAxisAlignment.START),
            padding=ft.padding.symmetric(horizontal=12, vertical=6),
            bgcolor=bg_color,
            border=ft.border.all(1, selected_border if is_selected else border_color),
            border_radius=6,
            margin=ft.margin.only(bottom=4),
            on_click=row_click_handler if row_click_handler else None
        )

    def _color(self, token: str, fallback: str) -> str:
        """Helper to resolve semantic colors with graceful fallback."""

        resolver: Optional[Callable[[str, str], Optional[str]]] = getattr(self.gui, "_themed_color", None)
        if callable(resolver):
            try:
                resolved = resolver(token, fallback)
                if resolved:
                    return resolved
            except Exception as exc:
                logger.debug("SearchResultsBuilder color fallback (%s): %s", token, exc)
        return fallback

    def _update_selection_indicator(self):
        if hasattr(self.gui, 'selected_count_text') and self.gui.selected_count_text:
            count = len(getattr(self.gui.state_manager.state, 'selected_tests', {}))
            self.gui.selected_count_text.value = f"Selected: {count} tests"
            self.gui.selected_count_text.visible = count > 0

    def _update_status(
        self,
        total_tests: int,
        visible_tests: int,
        current_sap: str,
        sap_total: int,
        sap_visible: int,
        filters_active: bool
    ):
        if not hasattr(self.gui, 'status_manager') or not self.gui.status_manager:
            return

        selected_count = len(getattr(self.gui.state_manager.state, 'selected_tests', {}))
        sap_info = ""
        if current_sap:
            if filters_active and sap_total != sap_visible:
                sap_info = f" | SAP {current_sap}: {sap_visible} of {sap_total} test(s)"
            else:
                sap_info = f" | SAP {current_sap}: {sap_total} test(s)"
        if filters_active and total_tests:
            sap_info += f" | {visible_tests} after filters"

        if selected_count > 0:
            self.gui.status_manager.update_status(
                f"✅ {selected_count} of {total_tests} tests selected{sap_info}. Click 'Configure' to continue.",
                "green"
            )
        elif total_tests > 0:
            self.gui.status_manager.update_status(
                f"Found {total_tests} test(s){sap_info}. Select the ones to include in your report.",
                "blue"
            )
        else:
            self.gui.status_manager.update_status("No tests found for this search.", "orange")

    def _update_navigation(self, tests):
        if not hasattr(self.gui, 'sap_navigation_container'):
            return

        if hasattr(self.gui, 'sap_navigator') and self.gui.sap_navigator and tests:
            nav_controls = self.gui.sap_navigator.create_navigation_controls(
                update_callback=self.gui.search_manager.display_search_results
            )
            self.gui.sap_navigation_container.content = nav_controls
            self.gui.sap_navigation_container.visible = len(self.gui.sap_navigator.sap_codes) >= 1
        else:
            self.gui.sap_navigation_container.visible = False


