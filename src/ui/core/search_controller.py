"""Search workflow orchestration extracted from `EventHandlers`.

The controller keeps search-specific logic together so `EventHandlers`
can delegate without carrying the entire implementation. It also makes
unit testing easier by allowing a focused interface that relies on a
small, explicit set of collaborators.
"""

from __future__ import annotations

import logging
import time
import traceback
from typing import Callable, Optional, TYPE_CHECKING

import flet as ft

from ...core.telemetry import log_duration
from ..utils.thread_pool import run_in_background
from ..utils.error_boundary import (
    with_error_boundary,
    safe_ui_update,
    search_circuit_breaker
)
from ..utils.loading_states import LoadingState

if TYPE_CHECKING:  # pragma: no cover - runtime circular imports avoided
    from .state_manager import StateManager
    from ..main_gui import MotorReportAppGUI

logger = logging.getLogger(__name__)


class SearchController:
    """Coordinate search-related UI updates and backend calls."""

    def __init__(
        self,
        gui: "MotorReportAppGUI",
        state_manager: "StateManager",
        *,
    update_button_state: Callable[..., bool],
        safe_status_update: Callable[[str, str], bool],
        safe_results_update: Callable[[], bool],
    ) -> None:
        self.gui = gui
        self.state_manager = state_manager
        self._update_button_state = update_button_state
        self._safe_status_update = safe_status_update
        self._safe_results_update = safe_results_update

    # ------------------------------------------------------------------
    # Public API used by EventHandlers

    def on_search_clicked(self, event: ft.ControlEvent) -> None:
        """Kick off a search run with UI feedback and concurrency guard."""
        # Guard against overlapping operations
        if self.state_manager.is_operation_in_progress():
            current_op = self.state_manager.get_current_operation()
            logger.warning("Search blocked - %s is in progress", current_op)
            self.gui.status_manager.update_status(
                f"Cannot search while {current_op} is in progress. Please wait...",
                "orange",
            )
            return

        query = self._current_query()
        if not query:
            self.gui.status_manager.update_status(
                "Please enter a SAP code or test number to search.",
                "red",
            )
            return

        if not self.state_manager.start_operation("search"):
            return

        # Immediate visual feedback
        self._update_button_state("search_button", enabled=False, text="Searching...", icon=ft.Icons.HOURGLASS_EMPTY)
        self._update_button_state("search_input_field", enabled=False)
        self.gui.status_manager.update_status(f"🔍 Searching for '{query}'...", "blue")
        self.gui.status_manager.show_progress(f"Searching for '{query}'...")
        if hasattr(self.gui, "progress_indicators"):
            self.gui.progress_indicators.show_progress(1)

        self._show_loading_placeholder(query)
        self.gui._safe_page_update()
        
        # Use page.invoke_later for non-blocking UI update instead of time.sleep
        def _start_search():
            if not self._ensure_backend_ready():
                self._cleanup_after_search()
                return
            self._start_background_search(query)
        
        if hasattr(self.gui.page, 'invoke_later'):
            self.gui.page.invoke_later(_start_search, 0.05)  # 50ms - faster than 100ms sleep
        else:
            _start_search()  # Fallback: immediate execution

    # ------------------------------------------------------------------
    # Internal helpers

    def _current_query(self) -> str:
        text_field = getattr(self.gui, "search_input_field", None)
        if text_field and text_field.value:
            return text_field.value.strip()
        return ""

    def _show_loading_placeholder(self, query: str) -> None:
        """Show enhanced loading state with skeleton."""
        if not hasattr(self.gui, "results_area"):
            return
        
        # Use enhanced loading state with skeleton
        self.gui.results_area.controls.clear()
        color_resolver = getattr(self.gui, "_themed_color", None)
        self.gui.results_area.controls.append(
            LoadingState.create_search_loading(query=query, color_resolver=color_resolver)
        )

    def _ensure_backend_ready(self) -> bool:
        if hasattr(self.gui, "app") and self.gui.app:
            return True

        logger.warning("Backend not initialized, attempting synchronous initialization...")
        try:
            self.gui.report_manager.initialize_backend()
        except Exception as exc:  # pragma: no cover - defensive logging
            logger.error("Backend initialization failed: %s", exc)
            self.gui.status_manager.update_status(
                f"[ERROR] Backend initialization failed: {exc}",
                "red",
            )
            self.gui.status_manager.hide_progress()
            return False

        if hasattr(self.gui, "app") and self.gui.app:
            logger.info("Backend initialized successfully during search trigger")
            return True

        self.gui.status_manager.update_status("[ERROR] Backend initialization failed. Check paths.", "red")
        self.gui.status_manager.hide_progress()
        return False

    def _start_background_search(self, query: str) -> None:
        """Launch search in background thread using thread pool."""
        if hasattr(self.gui, "page") and hasattr(self.gui.page, "run_thread"):
            self.gui.page.run_thread(self._perform_search, query)
        else:
            # Use thread pool for better resource management
            run_in_background(self._perform_search, query)

    def _perform_search(self, query: str) -> None:
        """Perform enhanced search in a background thread."""
        try:
            logger.info("Starting enhanced search for query: '%s'", query)

            if not self._safe_status_update(f"🔍 Analyzing search input: '{query}'", "blue"):
                self.gui.status_manager.show_progress(f"Searching for '{query}'...")

            # Reset previous results
            state = self.state_manager.state
            state.found_tests.clear()
            state.found_sap_codes.clear()

            if hasattr(self.gui, "app") and self.gui.app:
                logger.info("Calling backend analyze_search_input method...")
                with log_duration(logger, "analyze_search_input", level=logging.DEBUG):
                    search_analysis = self.gui.app.analyze_search_input(query)
            else:
                raise RuntimeError("Backend not available for search analysis")

            logger.info("Search analysis: %s", search_analysis)
            found_tests = search_analysis.get("found_tests", [])
            state.found_tests = found_tests

            found_sap_codes = {test.sap_code for test in found_tests if getattr(test, "sap_code", None)}
            state.found_sap_codes = sorted(found_sap_codes)
            logger.info("Found SAP codes: %s", state.found_sap_codes)
            
            # Record success for circuit breaker
            search_circuit_breaker.record_success()

            if search_analysis.get("status") == "empty":
                self._safe_status_update("[ERROR] No search query provided", "red")
                return

            self._provide_search_feedback(search_analysis)

            if not self._safe_results_update():
                try:
                    self.gui._display_search_results()
                    self.gui._safe_page_update()
                except Exception as fallback_error:  # pragma: no cover - defensive logging
                    logger.error("Fallback display update failed: %s", fallback_error)
        except Exception as exc:  # pragma: no cover - defensive logging
            logger.error("Search failed with exception: %s", exc)
            traceback.print_exc()
            # Record failure for circuit breaker
            search_circuit_breaker.record_failure()
            message = f"❌ Search failed: {exc}"
            if not self._safe_status_update(message, "red"):
                self.gui.status_manager.update_status(message, "red")
        finally:
            self._cleanup_after_search()

    def _cleanup_after_search(self) -> None:
        logger.info("Cleaning up search UI state")
        self.state_manager.end_operation()

        self._update_button_state("search_button", enabled=True, text="Search", icon=ft.Icons.SEARCH)
        self._update_button_state("search_input_field", enabled=True)

        if hasattr(self.gui, "status_manager") and self.gui.status_manager:
            self.gui.status_manager.hide_progress()

        if hasattr(self.gui, "progress_indicators"):
            self.gui.progress_indicators.hide_progress(1)

        self.gui._safe_page_update()

    def _provide_search_feedback(self, analysis: dict) -> None:
        """Generate user-facing feedback based on search analysis results."""
        total_inputs = analysis.get("total_inputs", 0)
        total_found = analysis.get("total_found", 0)
        strategy = analysis.get("search_strategy", "mixed")

        feedback_parts = []
        if total_found > 0:
            feedback_parts.append(f"✅ Found {total_found} test(s)")
            test_inputs = analysis.get("test_number_inputs", [])
            if test_inputs:
                feedback_parts.append(f"📊 {len(test_inputs)} test number(s) matched")

            sap_inputs = analysis.get("sap_code_inputs", [])
            if sap_inputs:
                total_sap_tests = sum(len(item.get("matches", [])) for item in sap_inputs)
                feedback_parts.append(f"🏷️ {len(sap_inputs)} SAP code(s) with {total_sap_tests} test(s)")

        unmatched = analysis.get("unmatched_inputs", [])
        if unmatched:
            feedback_parts.append(f"⚠️ No matches for: {', '.join(unmatched)}")

        feedback_msg = " | ".join(feedback_parts) if feedback_parts else (
            f"⚠️ No tests found for any of the {total_inputs} input(s)"
        )

        color = "orange"
        if total_found > 0:
            if strategy == "sap_codes_only" and total_found > 1:
                feedback_msg += " | 👆 Select tests from the SAP code(s) below"
            elif strategy == "test_numbers_only":
                feedback_msg += " | ✅ Test numbers found - ready to proceed"
            elif strategy == "mixed":
                feedback_msg += " | 🔄 Mixed search - review and select tests"
            color = "green"

        self._safe_status_update(feedback_msg, color)

