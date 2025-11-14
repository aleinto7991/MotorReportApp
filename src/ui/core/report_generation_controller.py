"""Report generation workflow controller extracted from `EventHandlers`.

This controller centralizes report-generation logic, including temporary
file handling, background execution, and UI feedback. Separating the
orchestration from the event-dispatch layer makes the flow easier to
test and reduces interdependencies.
"""

from __future__ import annotations

import datetime
import logging
import os
import shutil
import subprocess
import platform
import tempfile
import threading  # Keep for Lock
from typing import Callable, Optional, TYPE_CHECKING

from ..utils.thread_pool import run_in_background
from ..utils.error_boundary import (
    with_error_boundary,
    report_circuit_breaker
)
from ..utils.loading_states import LoadingState

import flet as ft

from ...core.telemetry import log_duration

if TYPE_CHECKING:  # pragma: no cover
    from ..main_gui import MotorReportAppGUI
    from .state_manager import StateManager

logger = logging.getLogger(__name__)


class ReportGenerationController:
    """Coordinate background report generation and file-save workflows."""

    def __init__(
        self,
        gui: "MotorReportAppGUI",
        state_manager: "StateManager",
        *,
        update_button_state: Callable[..., bool],
    ) -> None:
        self.gui = gui
        self.state_manager = state_manager
        self._update_button_state = update_button_state
        self._temp_file_lock = threading.Lock()
        self._temp_files_created = set()
        self._temp_report_file: Optional[str] = None
        self._report_filename: Optional[str] = None

    # ------------------------------------------------------------------
    # Public API used by EventHandlers

    def on_generate_report_clicked(self, _event: ft.ControlEvent) -> None:
        """Kick off report generation with concurrency guard."""
        if self.state_manager.is_operation_in_progress():
            current_op = self.state_manager.get_current_operation()
            logger.warning("Report generation blocked - %s is in progress", current_op)
            self.gui.status_manager.update_status(
                f"Cannot generate report while {current_op} is in progress. Please wait...",
                "orange",
            )
            return

        if not self.state_manager.state.selected_tests:
            self.gui.status_manager.update_status(
                "No tests selected. Please select tests in the Search & Select tab.",
                "red",
            )
            return

        if not self.state_manager.start_operation("report_generation"):
            return

        self._update_generate_button_state(generating=True)
        self.gui.status_manager.update_status("🔄 Starting report generation...", "blue")

        # Use thread pool for better resource management
        run_in_background(self._generate_report)

    def save_report_to_folder(self, folder_path: str) -> None:
        """Copy the generated report from temp to the specified folder."""
        if not self._temp_report_file or not self._report_filename:
            logger.warning("No temp report file available for saving")
            return

        if not os.path.exists(self._temp_report_file):
            logger.error("Temp report file does not exist: %s", self._temp_report_file)
            self.gui.status_manager.update_status("❌ Temp file not found. Generate report again.", "red")
            return

        try:
            final_path = os.path.join(folder_path, self._report_filename)
            with log_duration(logger, "copy_report_to_final_location", level=logging.DEBUG):
                shutil.copy2(self._temp_report_file, final_path)

            self.gui.status_manager.update_status(
                f"✅ Report saved: {self._report_filename}",
                "green",
            )
            self.gui.status_manager.hide_progress()

            self._show_download_success_dialog(self._report_filename, final_path)

        except Exception as exc:  # pragma: no cover - defensive logging
            logger.error("Error saving report to folder %s: %s", folder_path, exc)
            self.gui.status_manager.update_status(f"❌ Save failed: {exc}", "red")

    # ------------------------------------------------------------------
    # Internal helpers

    def _update_generate_button_state(self, *, generating: bool) -> None:
        try:
            if generating:
                self._update_button_state(
                    "generate_button",
                    False,
                    "Generating...",
                    ft.Icons.HOURGLASS_EMPTY,
                    None,
                    None,
                )
            else:
                self._update_button_state(
                    "generate_button",
                    True,
                    "Generate Report",
                    ft.Icons.CREATE,
                    None,
                    None,
                )
            self.gui._safe_page_update()
        except Exception as exc:  # pragma: no cover - defensive logging
            logger.debug("Button state update failed: %s", exc)

    def _generate_report(self) -> None:
        """Execute report generation with error tracking."""
        try:
            self.gui.status_manager.show_progress("Preparing report generation...")

            tests_to_process = self.state_manager.get_tests_to_process()
            if not tests_to_process:
                raise RuntimeError("No tests selected for processing")

            noise_saps = list(self.state_manager.state.selected_noise_saps)
            comparison_saps = list(self.state_manager.state.selected_comparison_saps)

            logger.info("Report generation data:")
            logger.info("  Performance tests: %d tests", len(tests_to_process))
            logger.info("  Tests: %s", [(t.test_lab_number, t.sap_code) for t in tests_to_process])
            logger.info("  Noise SAPs: %s", noise_saps)
            logger.info("  Comparison SAPs: %s", comparison_saps)

            self._log_fine_grained_selections(comparison_saps, noise_saps)

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Motor_Performance_Report_{timestamp}.xlsx"
            self.gui.status_manager.update_status(f"🔄 Generating report: {filename}...", "blue")

            if not hasattr(self.gui, "report_manager") or not self.gui.report_manager:
                raise RuntimeError("Report manager not available. Please restart the application.")

            temp_path = self._create_temp_file_safely(filename)
            logger.info("Generating report to temporary location: %s", temp_path)

            multiple_comparisons = self._build_multiple_comparisons()

            with log_duration(logger, "generate_report_with_path"):
                self.gui.report_manager.generate_report_with_path(
                    tests_to_process=tests_to_process,
                    noise_saps=noise_saps,
                    comparison_saps=comparison_saps,
                    multiple_comparisons=multiple_comparisons,
                    output_path=temp_path,
                )
            
            # Record success for circuit breaker
            report_circuit_breaker.record_success()
            
            self.gui.status_manager.hide_progress()

        except Exception as exc:  # pragma: no cover - defensive logging
            logger.error("Report generation failed: %s", exc)
            self.gui.status_manager.update_status(f"❌ Report generation failed: {exc}", "red")
            self.gui.status_manager.hide_progress()
            self._show_report_error_dialog(str(exc))
        finally:
            self.state_manager.end_operation()
            self._update_generate_button_state(generating=False)

    def _log_fine_grained_selections(self, comparison_saps: list, noise_saps: list) -> None:
        if comparison_saps:
            for sap in comparison_saps:
                selected_labs = self.state_manager.state.selected_comparison_test_labs.get(sap, set())
                if selected_labs:
                    logger.info("  Comparison SAP %s: Selected test labs %s", sap, list(selected_labs))
                else:
                    logger.info("  Comparison SAP %s: All tests (no specific selection)", sap)

        if noise_saps:
            for sap in noise_saps:
                selected_noise_tests = self.state_manager.state.selected_noise_test_labs.get(sap, set())
                if selected_noise_tests:
                    logger.info("  Noise SAP %s: Selected tests %s", sap, list(selected_noise_tests))
                else:
                    logger.info("  Noise SAP %s: All tests (no specific selection)", sap)

    def _build_multiple_comparisons(self) -> list:
        multiple_comparisons = []
        state = self.state_manager.state

        if hasattr(state, "comparison_groups") and state.comparison_groups:
            logger.info("Converting new comparison_groups format to multiple_comparisons for report")
            for group_id, group_data in state.comparison_groups.items():
                if isinstance(group_data, dict) and group_data:
                    all_test_labs = []
                    for sap_code, test_labs in group_data.items():
                        if test_labs:
                            all_test_labs.extend(list(test_labs))

                    if all_test_labs:
                        comparison_group = {
                            "id": group_id,
                            "name": group_id,
                            "test_labs": all_test_labs,
                            "description": f"Comparison group with {len(group_data)} SAPs",
                            "sap_data": group_data,
                        }
                        multiple_comparisons.append(comparison_group)
                        logger.info(
                            "  Converted group %s: %d test labs from %d SAPs",
                            group_id,
                            len(all_test_labs),
                            len(group_data),
                        )

        if not multiple_comparisons and hasattr(state, "multiple_comparisons"):
            multiple_comparisons = state.multiple_comparisons
            logger.info("Using existing multiple_comparisons format")

        logger.info("Final multiple_comparisons for report: %d groups", len(multiple_comparisons))
        return multiple_comparisons

    def _create_temp_file_safely(self, filename: str) -> str:
        with self._temp_file_lock:
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, filename)
            self._temp_files_created.add(temp_path)
            self._temp_report_file = temp_path
            self._report_filename = filename
            return temp_path

    def _show_download_success_dialog(self, filename: str, full_path: str) -> None:
        def on_open_folder(_: ft.ControlEvent) -> None:
            folder_path = os.path.dirname(full_path)
            try:
                if platform.system() == "Windows":
                    subprocess.run(f'explorer /select,"{full_path}"', shell=True, check=False)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", "-R", full_path], check=False)
                else:
                    subprocess.run(["xdg-open", folder_path], check=False)
                success_dialog.open = False
                self.gui.page.update()
            except Exception as exc:  # pragma: no cover - defensive logging
                logger.error("Error opening folder: %s", exc)

        def on_close_success(_: ft.ControlEvent) -> None:
            success_dialog.open = False
            self.gui.page.update()

        success_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("✅ Download Complete!", weight=ft.FontWeight.BOLD, color="green"),
            content=ft.Container(
                content=ft.Column(
                    [
                        ft.Text("Your report has been saved successfully!", size=14),
                        ft.Container(height=10),
                        ft.Container(
                            content=ft.Column(
                                [
                                    ft.Text("📄 Filename:", size=12, weight=ft.FontWeight.BOLD),
                                    ft.Text(filename, size=13, selectable=True),
                                    ft.Container(height=8),
                                    ft.Text("📁 Location:", size=12, weight=ft.FontWeight.BOLD),
                                    ft.Text(os.path.dirname(full_path), size=12, selectable=True, color="#666666"),
                                ],
                                spacing=4,
                            ),
                            padding=ft.padding.all(12),
                            bgcolor="#f0f8f0",
                            border_radius=8,
                            border=ft.border.all(1, "#4CAF50"),
                        ),
                    ],
                    spacing=10,
                    tight=True,
                ),
                width=450,
                padding=ft.padding.all(20),
            ),
            actions=[
                ft.TextButton("Close", on_click=on_close_success),
                ft.ElevatedButton(
                    "📂 Open Folder",
                    on_click=on_open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(bgcolor="#4CAF50", color="#FFFFFF"),
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        self.gui.page.overlay.append(success_dialog)
        success_dialog.open = True
        self.gui.page.update()

    def _show_report_error_dialog(self, error_message: str) -> None:
        def on_close(_: ft.ControlEvent) -> None:
            error_dialog.open = False
            self.gui.page.update()

        error_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("❌ Report Generation Failed", weight=ft.FontWeight.BOLD, color="red"),
            content=ft.Container(
                content=ft.Column(
                    [
                        ft.Text("An error occurred while generating the report:", size=14),
                        ft.Container(height=10),
                        ft.Container(
                            content=ft.Text(error_message, size=12, selectable=True),
                            padding=ft.padding.all(12),
                            bgcolor="#fff0f0",
                            border_radius=8,
                            border=ft.border.all(1, "#f44336"),
                        ),
                    ],
                    spacing=10,
                    tight=True,
                ),
                width=450,
                padding=ft.padding.all(20),
            ),
            actions=[ft.TextButton("Close", on_click=on_close)],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        self.gui.page.overlay.append(error_dialog)
        error_dialog.open = True
        self.gui.page.update()

