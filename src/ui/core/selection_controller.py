"""Selection workflow controller extracted from `EventHandlers`.

This controller centralizes the logic responsible for updating the UI
when tests are selected or deselected. Keeping the behaviour isolated
makes it easier to reuse in future views and simplifies unit testing.
"""

from __future__ import annotations

import logging
from typing import Callable, TYPE_CHECKING

import flet as ft

from ..utils.debouncer import selection_debouncer

if TYPE_CHECKING:  # pragma: no cover - import hints only
    from ..main_gui import MotorReportAppGUI
    from .state_manager import StateManager
    from ...data.models import Test

logger = logging.getLogger(__name__)


class SelectionController:
    """Coordinate UI state for performance/noise test selections."""

    def __init__(
        self,
        gui: "MotorReportAppGUI",
        state_manager: "StateManager",
    ) -> None:
        self.gui = gui
        self.state_manager = state_manager

    # ------------------------------------------------------------------
    # Public API used by EventHandlers

    def on_test_selected(self, event: ft.ControlEvent) -> None:
        control = event.control
        test = getattr(control, "data", None)
        if test is None:
            logger.warning("Selection event without test data")
            return

        test_lab = getattr(test, "test_lab_number", None)
        if not test_lab:
            logger.warning("Selection event missing test_lab_number")
            return

        selected = bool(control.value)
        logger.debug("Test selection changed: %s -> %s", test_lab, "SELECTED" if selected else "DESELECTED")

        linked_tests = self._resolve_linked_tests(test) if selected else [test]

        for linked_test in linked_tests:
            linked_id = getattr(linked_test, "test_lab_number", None)
            if not linked_id:
                continue
            self.state_manager.update_test_selection(linked_id, linked_test, selected)

        self._apply_current_selection_state()
        self._refresh_search_results()
        # Debounce UI updates to prevent excessive re-renders during bulk selections
        self._debounced_update_selection_ui()

    def row_click_handler(self, test: "Test") -> Callable[[ft.ControlEvent], None]:
        def _handler(_: ft.ControlEvent) -> None:
            current_selected = test.test_lab_number in self.state_manager.state.selected_tests
            target_state = not current_selected

            linked_tests = self._resolve_linked_tests(test) if target_state else [test]
            for linked_test in linked_tests:
                linked_id = getattr(linked_test, "test_lab_number", None)
                if not linked_id:
                    continue
                self.state_manager.update_test_selection(linked_id, linked_test, target_state)

            self._apply_current_selection_state()
            self._refresh_search_results()
            # Debounce UI updates for row clicks too
            self._debounced_update_selection_ui()

        return _handler

    # ------------------------------------------------------------------
    # Internal helpers

    def _apply_current_selection_state(self) -> None:
        state = self.state_manager.state
        has_selections = bool(state.selected_tests)
        state.search_selection_applied = has_selections
        if has_selections:
            self.state_manager.apply_search_selection()
            logger.debug("Auto-applied selection: %d tests", len(state.selected_tests))
    
    @selection_debouncer.debounce
    def _debounced_update_selection_ui(self) -> None:
        """Debounced version of _update_selection_ui to prevent excessive updates."""
        self._update_selection_ui()

    def _update_selection_ui(self) -> None:
        state = self.state_manager.state
        try:
            if hasattr(self.gui, "selected_count_text") and self.gui.selected_count_text:
                count = len(state.selected_tests)
                self.gui.selected_count_text.value = f"Selected: {count} tests"
                self.gui.selected_count_text.visible = count > 0
                self.gui.selected_count_text.update()

            if hasattr(self.gui, "status_manager") and self.gui.status_manager:
                count = len(state.selected_tests)
                if count > 0:
                    message = "✅ {count} test(s) selected. Click 'Configure' tab to continue.".format(count=count)
                    self.gui.status_manager.update_status(message, "green")
                else:
                    self.gui.status_manager.update_status("Select tests to include in your report.", "blue")
        except Exception as exc:  # pragma: no cover - defensive logging
            logger.warning("Error while updating selection UI: %s", exc)
            # Fallback to a full refresh if incremental updates fail
            if hasattr(self.gui, "_display_search_results"):
                self.gui._display_search_results()
                self.gui._safe_page_update()

    # ------------------------------------------------------------------
    # Linked selection helpers

    def _resolve_linked_tests(self, primary_test: "Test") -> list["Test"]:
        """Return tests that should be auto-selected alongside *primary_test*."""
        results: dict[str, "Test"] = {}

        def _add(test_obj: "Test") -> None:
            test_id = getattr(test_obj, "test_lab_number", None)
            if not test_id:
                return
            normalized_id = test_id.strip().upper()
            if normalized_id and normalized_id not in results:
                results[normalized_id] = test_obj

        _add(primary_test)

        primary_id = getattr(primary_test, "test_lab_number", "")
        normalized = primary_id.strip().upper()
        if not normalized:
            return list(results.values())

        if normalized.endswith("A"):
            base = normalized[:-1]
        else:
            base = normalized

        candidate_ids = {normalized}
        if base:
            candidate_ids.add(base)
            candidate_ids.add(f"{base}A")

        for test in self.state_manager.state.found_tests:
            test_id = getattr(test, "test_lab_number", "")
            if not test_id:
                continue
            candidate = test_id.strip().upper()
            if candidate in candidate_ids:
                _add(test)

        return list(results.values())

    def _refresh_search_results(self) -> None:
        search_manager = getattr(self.gui, "search_manager", None)
        if not search_manager:
            return

        display_fn = getattr(search_manager, "display_search_results", None)
        if not callable(display_fn):
            return

        try:
            page = getattr(self.gui, "page", None)
            invoke_later = getattr(page, "invoke_later", None)
            if callable(invoke_later):
                invoke_later(display_fn)
            else:
                display_fn()
        except Exception as exc:  # pragma: no cover - defensive logging
            logger.debug("Failed to refresh search results after linked selection: %s", exc)

