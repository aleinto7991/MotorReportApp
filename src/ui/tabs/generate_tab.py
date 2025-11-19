"""
Generate Tab - Fourth tab for generating the report
"""
import flet as ft
import logging
import datetime
import time
import traceback
from typing import Optional, Any, Dict, List
from ..components.base import BaseTab
from ...data.models import Test

logger = logging.getLogger(__name__)


class GenerateTab(BaseTab):
    """Tab for generating the motor performance report"""
    
    def __init__(self, parent_gui=None):
        super().__init__(parent_gui)
        self.tab_name = "4. Generate"
        self.tab_icon = ft.Icons.CREATE

    def _color(self, token: str, fallback: str) -> str:
        """Shorthand for resolving themed colors with safe fallbacks."""

        return self.theme_color(token, fallback)
    
    def get_tab_content(self) -> ft.Control:
        """Build the generation tab content"""
        # Initialize progress indicators
        step3_progress = None
        step3_status = None
        step4_progress = None
        step4_status = None
        
        try:
            logger.info("🔧 Building fresh Generate tab content...")
            
            # Get progress indicators from parent
            if self.parent_gui and hasattr(self.parent_gui, 'progress_indicators'):
                step3_progress, step3_status = self.parent_gui.progress_indicators.get_indicators_for_step(3)
                step4_progress, step4_status = self.parent_gui.progress_indicators.get_indicators_for_step(4)
            
                # Always create fresh summary container to avoid reference issues
            self.summary_container = ft.Container(
                content=ft.Column([
                    ft.Text(
                        "Selected Tests Summary:",
                        weight=ft.FontWeight.W_500,
                        size=16,
                        color=self._color('on_surface', '#1f2933')
                    ),
                    self._build_tests_summary(),
                ], spacing=10),
                visible=True,
                expand=False,
                bgcolor=self._color('surface', '#ffffff'),
                border_radius=8,
                border=ft.border.all(1, self._color('outline', '#d0d7e5')),
                padding=ft.padding.all(12)
            )
            
            logger.info("✅ Generate tab content built successfully")
        except Exception as e:
            logger.error(f"❌ Error building Generate tab content: {e}")
            # Return a simple error content instead of failing
            self.summary_container = ft.Container(
                content=ft.Column([
                    ft.Text(
                        "Error loading content",
                        weight=ft.FontWeight.W_500,
                        size=16,
                        color=self._color('error', '#c62828')
                    ),
                    ft.Text(f"Error: {str(e)}", size=12, color=self._color('text_muted', '#5f6b7a')),
                ], spacing=10),
                visible=True,
                expand=False,
                bgcolor=self._color('error_container', '#ffebee'),
                border_radius=8,
                padding=ft.padding.all(12)
            )
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Generate Report", size=20, weight=ft.FontWeight.BOLD),
                    step3_progress or ft.Container(),
                    step3_status or ft.Container(),
                    step4_progress or ft.Container(),
                    step4_status or ft.Container(),
                ], spacing=10),
                
                ft.Text(
                    "Generate your motor performance report with the selected tests and configuration.", 
                    color=self._color('text_muted', '#5f6b7a'),
                    size=14
                ),
                ft.Divider(),
                
                self.summary_container,
                
                ft.Divider(),
                self.parent_gui.generate_button if self.parent_gui else ft.ElevatedButton("Generate Report"),
                
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.WARNING, color=self._color('warning', '#f57c00')),
                        ft.Text(
                            "Make sure all previous steps are completed before generating the report.", 
                            color=self._color('warning', '#f57c00'),
                            size=14
                        )
                    ], spacing=10),
                    padding=ft.padding.all(15),
                    bgcolor=self._color('warning_container', '#fff3e0'),
                    border_radius=5,
                    border=ft.border.all(1, self._color('outline', '#ffcc80'))
                )
            ], spacing=15),
            padding=ft.padding.all(20),
            expand=True
        )
    
                
    def _build_tests_summary(self) -> ft.Control:
        """Build the selected tests summary display with interactive controls"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            logger.warning("🔍 _build_tests_summary: No state manager found")
            return ft.Text("❌ No state manager found", color=self._color('error', '#c62828'), size=18)

        state_manager = self.parent_gui.state_manager
        state = state_manager.state
        state_manager.refresh_carichi_matches()
        selected_tests = state.selected_tests

        logger.info(f"🔍 GenerateTab Debug - selected_tests count: {len(selected_tests)}")
        logger.info(f"🔍 GenerateTab Debug - selected_tests keys: {list(selected_tests.keys())}")
        logger.info(f"🔍 GenerateTab Debug - selected_noise_saps: {state.selected_noise_saps}")
        logger.info(f"🔍 GenerateTab Debug - selected_comparison_saps: {state.selected_comparison_saps}")
        logger.info(f"🔍 GenerateTab Debug - config_selection_applied: {state.config_selection_applied}")

        if not selected_tests:
            logger.warning("🔍 _build_tests_summary: No selected tests found")
            return ft.Text(
                "⚠️ No tests selected yet. Go to Search & Select tab and select tests.",
                color=self._color('warning', '#f57c00'),
                size=16
            )

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        header = ft.Text(
            f"🔄 Last Updated: {timestamp}",
            size=14,
            color=self._color('success', '#2e7d32'),
            weight=ft.FontWeight.BOLD
        )

        performance_section = self._build_performance_summary_section(state)
        carichi_section = self._build_carichi_summary_section(state)
        noise_section = self._build_noise_summary_section(state)
        lf_section = self._build_lf_summary_section(state)
        comparison_section = self._build_comparison_summary_section(state)
        data_flow_section = self._build_data_flow_section(state)

        return ft.Column(
            controls=[
                header,
                performance_section,
                carichi_section,
                noise_section,
                lf_section,
                comparison_section,
                data_flow_section,
            ],
            spacing=12,
        )

    def _build_performance_summary_section(self, state) -> ft.Control:
        selected_tests = state.selected_tests
        sap_groups: Dict[str, List[Test]] = {}
        for test in selected_tests.values():
            sap_code = test.sap_code or "Unknown"
            sap_groups.setdefault(sap_code, []).append(test)

        sap_rows: List[ft.Control] = []
        for sap_code in sorted(sap_groups.keys()):
            tests_for_sap = sap_groups[sap_code]

            sap_header = ft.Row(
                controls=[
                    ft.Text(
                        f"{sap_code}: {len(tests_for_sap)} test(s)",
                        size=13,
                        color=self._color('primary', '#0f4c81'),
                        weight=ft.FontWeight.W_500,
                    ),
                    ft.IconButton(
                        icon=ft.Icons.DELETE_FOREVER,
                        icon_color=self._color('error', '#d32f2f'),
                        tooltip=f"Remove all tests for SAP {sap_code}",
                        on_click=lambda e, sap=sap_code: self._handle_remove_performance_sap(sap),
                    ),
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            )

            test_rows: List[ft.Control] = []
            for test in sorted(tests_for_sap, key=lambda t: t.test_lab_number):
                voltage_display = f"{test.voltage}V" if test.voltage and test.voltage.strip() else "N/A"
                notes_display = test.notes.strip() if test.notes and test.notes.strip() else "No notes"
                if len(notes_display) > 40:
                    notes_display = notes_display[:37] + "..."

                test_rows.append(
                    ft.Row(
                        controls=[
                            ft.Text(
                                f"• Test {test.test_lab_number}: {voltage_display} | {notes_display}",
                                size=12,
                                color=self._color('on_surface', '#1f2933'),
                            ),
                            ft.IconButton(
                                icon=ft.Icons.CANCEL,
                                icon_color=self._color('error', '#f44336'),
                                tooltip=f"Remove test {test.test_lab_number}",
                                on_click=lambda e, test_id=test.test_lab_number: self._handle_remove_performance_test(test_id),
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    )
                )

            sap_rows.append(
                ft.Container(
                    content=ft.Column(controls=[sap_header, *test_rows], spacing=3),
                    padding=ft.padding.symmetric(horizontal=10, vertical=6),
                    bgcolor=self._color('surface_variant', '#f0f4ff'),
                    border_radius=6,
                    border=ft.border.all(1, self._color('outline', '#d0d7e5')),
                )
            )

        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text(
                        "📊 PERFORMANCE SHEET",
                        size=18,
                        color=self._color('primary', '#1565c0'),
                        weight=ft.FontWeight.BOLD
                    ),
                    ft.Text(
                        f"All {len(selected_tests)} selected tests from Step 1 will be included",
                        size=14,
                        color=self._color('success', '#2e7d32'),
                    ),
                    *sap_rows,
                ],
                spacing=6,
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('primary_container', '#e3f2fd'),
            border_radius=8,
            border=ft.border.all(1, self._color('outline', '#90caf9')),
        )

    def _build_carichi_summary_section(self, state) -> ft.Control:
        state_manager = getattr(self.parent_gui, 'state_manager', None)
        if not state_manager:
            return ft.Container(
                content=ft.Text(
                    "Carichi precheck unavailable - no state manager",
                    color=self._color('error', '#c62828'),
                    size=14,
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('error_container', '#ffebee'),
                border_radius=8,
            )

        status = state_manager.get_carichi_status()
        accent = self._color('tertiary', '#6a1b9a')
        container_bg = self._color('tertiary_container', '#f3e5f5')
        muted = self._color('text_muted', '#5f6b7a')

        header = ft.Text(
            "⚙️ CARICHI NOMINALI PRECHECK",
            size=18,
            color=accent,
            weight=ft.FontWeight.BOLD,
        )

        if not status.get('enabled'):
            return ft.Container(
                content=ft.Column(
                    controls=[
                        header,
                        ft.Text(
                            "Configure the Test Lab directory in the Setup tab to enable Carichi precheck.",
                            size=14,
                            color=muted,
                        ),
                    ],
                    spacing=6,
                ),
                padding=ft.padding.all(10),
                bgcolor=container_bg,
                border_radius=8,
                border=ft.border.all(1, self._color('outline', '#ce93d8')),
            )

        path_text = ft.Text(
            f"Directory: {status.get('path') or 'Not set'}",
            size=12,
            color=muted,
            selectable=True,
        )

        if status.get('errors'):
            error_controls = [
                ft.Text(
                    f"❌ {message}",
                    size=13,
                    color=self._color('error', '#c62828'),
                )
                for message in status['errors']
            ]

            return ft.Container(
                content=ft.Column(
                    controls=[header, path_text, *error_controls],
                    spacing=6,
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('error_container', '#ffebee'),
                border_radius=8,
                border=ft.border.all(1, self._color('outline', '#ffcdd2')),
            )

        total_tests = status.get('total_tests', 0)
        coverage_line = ft.Text(
            f"Coverage: {status.get('resolved_count', 0)}/{total_tests} test(s) ({status.get('coverage_percent', 0.0)}%)",
            size=14,
            color=accent,
            weight=ft.FontWeight.W_500,
        )

        last_checked_display = None
        last_checked = status.get('last_checked')
        if last_checked:
            try:
                checked_dt = datetime.datetime.fromisoformat(last_checked)
                last_checked_display = checked_dt.strftime("%H:%M:%S")
            except ValueError:
                last_checked_display = last_checked

        info_controls: List[ft.Control] = [header, path_text]

        if total_tests == 0:
            info_controls.append(
                ft.Text(
                    "No performance tests selected yet. Carichi files will be resolved after Step 1.",
                    size=14,
                    color=muted,
                )
            )
        else:
            info_controls.append(coverage_line)

            if last_checked_display:
                info_controls.append(
                    ft.Text(
                        f"Last checked: {last_checked_display}",
                        size=12,
                        color=muted,
                    )
                )

            year_counts = status.get('year_counts') or {}
            if year_counts:
                year_summary = ", ".join(
                    f"{year}: {count}"
                    for year, count in sorted(year_counts.items(), reverse=True)
                )
                info_controls.append(
                    ft.Text(
                        f"Year folders: {year_summary}",
                        size=12,
                        color=muted,
                    )
                )

            missing_details = status.get('missing_details') or []
            if missing_details:
                info_controls.append(
                    ft.Text(
                        "Missing workbook(s):",
                        size=13,
                        color=self._color('error', '#c62828'),
                        weight=ft.FontWeight.W_500,
                    )
                )
                max_rows = 5
                for entry in missing_details[:max_rows]:
                    info_controls.append(
                        ft.Text(
                            f"• {entry.get('test_number')} ({entry.get('sap_code')})",
                            size=12,
                            color=self._color('error', '#c62828'),
                        )
                    )
                if len(missing_details) > max_rows:
                    info_controls.append(
                        ft.Text(
                            f"…and {len(missing_details) - max_rows} more",
                            size=12,
                            color=self._color('error', '#c62828'),
                        )
                    )
            else:
                info_controls.append(
                    ft.Text(
                        "✅ All selected tests have Carichi workbook matches",
                        size=14,
                        color=self._color('success', '#2e7d32'),
                        weight=ft.FontWeight.W_500,
                    )
                )

        return ft.Container(
            content=ft.Column(controls=info_controls, spacing=6),
            padding=ft.padding.all(10),
            bgcolor=container_bg,
            border_radius=8,
            border=ft.border.all(1, self._color('outline', '#ce93d8')),
        )

    def _build_noise_summary_section(self, state) -> ft.Control:
        if not state.include_noise:
            return ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text(
                            "🔊 NOISE ANALYSIS",
                            size=18,
                            color=self._color('success', '#2e7d32'),
                            weight=ft.FontWeight.BOLD
                        ),
                        ft.Text(
                            "Noise analysis disabled in configuration",
                            size=14,
                            color=self._color('text_muted', '#5f6b7a')
                        ),
                    ],
                    spacing=6,
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('surface_variant', '#f5f5f5'),
                border_radius=8,
                border=ft.border.all(1, self._color('outline', '#e0e0e0')),
            )

        sap_rows: List[ft.Control] = []
        for sap_code in sorted(state.selected_noise_saps):
            test_labs = sorted(state.selected_noise_test_labs.get(sap_code, set()))

            sap_header = ft.Row(
                controls=[
                    ft.Text(
                        f"{sap_code}: {len(test_labs)} noise test(s)",
                        size=13,
                        color=self._color('success', '#2e7d32'),
                        weight=ft.FontWeight.W_500,
                    ),
                    ft.IconButton(
                        icon=ft.Icons.DELETE_FOREVER,
                        icon_color=self._color('error', '#c62828'),
                        tooltip=f"Remove SAP {sap_code} from noise analysis",
                        on_click=lambda e, sap=sap_code: self._handle_remove_noise_sap(sap),
                    ),
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            )

            if test_labs:
                test_rows = [
                    ft.Row(
                        controls=[
                            ft.Text(f"• Test {lab}", size=12, color=self._color('on_surface', '#1f2933')),
                            ft.IconButton(
                                icon=ft.Icons.CANCEL,
                                icon_color=self._color('error', '#d32f2f'),
                                tooltip=f"Remove noise test {lab}",
                                on_click=lambda e, sap=sap_code, lab=lab: self._handle_remove_noise_test(sap, lab),
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    )
                    for lab in test_labs
                ]
            else:
                test_rows = [ft.Text("• No noise tests selected", size=12, color=self._color('warning', '#f57c00'))]

            sap_rows.append(
                ft.Container(
                    content=ft.Column(controls=[sap_header, *test_rows], spacing=3),
                    padding=ft.padding.symmetric(horizontal=10, vertical=6),
                    bgcolor=self._color('success_container', '#e8f5e9'),
                    border_radius=6,
                )
            )

        info_text = (
            ft.Text(
                f"✅ Noise analysis enabled for {len(state.selected_noise_saps)} SAP code(s)",
                size=14,
                color=self._color('success', '#2e7d32'),
            )
            if state.selected_noise_saps
            else ft.Text(
                "⚠️ No SAP codes selected for noise analysis",
                size=14,
                color=self._color('warning', '#f57c00')
            )
        )

        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text(
                        "🔊 NOISE ANALYSIS",
                        size=18,
                        color=self._color('success', '#2e7d32'),
                        weight=ft.FontWeight.BOLD
                    ),
                    info_text,
                    *sap_rows,
                ],
                spacing=6,
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('success_container', '#e8f5e8'),
            border_radius=8,
            border=ft.border.all(1, self._color('outline', '#a5d6a7')),
        )

    def _build_lf_summary_section(self, state) -> ft.Control:
        """Build summary section for Life Test (LF) data"""
        
        if not hasattr(state, 'selected_lf_test_numbers') or not state.selected_lf_test_numbers:
            return ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Text(
                            "🔬 LIFE TEST (LF) DATA",
                            size=18,
                            color=self._color('primary', '#1976d2'),
                            weight=ft.FontWeight.BOLD
                        ),
                        ft.Text(
                            "⚠️ No LF tests selected",
                            size=14,
                            color=self._color('warning', '#f57c00')
                        ),
                    ],
                    spacing=6,
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('primary_container', '#e3f2fd'),
                border_radius=8,
                border=ft.border.all(1, self._color('outline', '#90caf9')),
            )
        
        # Build list of selected LF tests
        sap_rows = []
        total_tests = 0
        
        for sap_code, test_numbers in state.selected_lf_test_numbers.items():
            if not test_numbers:
                continue
            
            total_tests += len(test_numbers)
            
            sap_header = ft.Text(
                f"SAP {sap_code}: {len(test_numbers)} LF test(s)",
                size=14,
                weight=ft.FontWeight.BOLD,
                color=self._color('primary', '#1565c0'),
            )
            
            # Create rows for each test
            test_rows = [
                ft.Row(
                    controls=[
                        ft.Text(f"• {test_num}", size=12),
                        ft.IconButton(
                            icon=ft.Icons.REMOVE_CIRCLE_OUTLINE,
                            icon_size=16,
                            icon_color=self._color('error', '#c62828'),
                            tooltip=f"Remove LF test {test_num}",
                            on_click=lambda e, s=sap_code, t=test_num: self._remove_lf_test(s, t),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                )
                for test_num in sorted(test_numbers)
            ]
            
            sap_rows.append(
                ft.Container(
                    content=ft.Column(controls=[sap_header, *test_rows], spacing=3),
                    padding=ft.padding.symmetric(horizontal=10, vertical=6),
                    bgcolor=self._color('primary_container', '#e3f2fd'),
                    border_radius=6,
                )
            )
        
        info_text = ft.Text(
            f"✅ {total_tests} LF test(s) selected across {len(state.selected_lf_test_numbers)} SAP code(s)",
            size=14,
            color=self._color('primary', '#1565c0'),
        )
        
        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text(
                        "🔬 LIFE TEST (LF) DATA",
                        size=18,
                        color=self._color('primary', '#1976d2'),
                        weight=ft.FontWeight.BOLD,
                    ),
                    
                    info_text,
                    *sap_rows,
                ],
                spacing=6,
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('primary_container', '#e3f2fd'),
            border_radius=8,
            border=ft.border.all(1, self._color('outline', '#90caf9')),
        )
    
    def _remove_lf_test(self, sap_code: str, test_number: str):
        """Remove an LF test from selection"""
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        
        state = self.parent_gui.state_manager.state
        if sap_code in state.selected_lf_test_numbers:
            state.selected_lf_test_numbers[sap_code].discard(test_number)
            
            # If no more tests for this SAP, remove the SAP entry
            if not state.selected_lf_test_numbers[sap_code]:
                del state.selected_lf_test_numbers[sap_code]
                state.selected_lf_saps.discard(sap_code)
        
        # Refresh the generate tab
        self.refresh_content()

    def _build_comparison_summary_section(self, state) -> ft.Control:
        selected_tests = state.selected_tests
        has_groups = bool(state.comparison_groups)
        has_legacy = state.include_comparison and bool(state.selected_comparison_saps)

        group_controls: List[ft.Control] = []

        if has_groups:
            for group_id in sorted(state.comparison_groups.keys()):
                sap_map = state.comparison_groups[group_id]
                group_header = ft.Row(
                    controls=[
                        ft.Text(
                            f"{group_id}: {sum(len(t) for t in sap_map.values())} test lab(s)",
                            size=13,
                            color=self._color('warning', '#f57c00'),
                            weight=ft.FontWeight.W_500,
                        ),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_FOREVER,
                            icon_color=self._color('error', '#c62828'),
                            tooltip=f"Remove comparison group {group_id}",
                            on_click=lambda e, gid=group_id: self._handle_remove_comparison_group(gid),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                )

                sap_entries: List[ft.Control] = []
                for sap_code in sorted(sap_map.keys()):
                    test_labs = sorted(sap_map[sap_code])
                    sap_row_header = ft.Row(
                        controls=[
                            ft.Text(
                                f"• {sap_code}: {len(test_labs)} test(s)",
                                size=12,
                                color=self._color('warning', '#f57c00'),
                            ),
                            ft.IconButton(
                                icon=ft.Icons.DELETE,
                                icon_color=self._color('warning', '#fb8c00'),
                                tooltip=f"Remove {sap_code} from {group_id}",
                                on_click=lambda e, gid=group_id, sap=sap_code: self._handle_remove_comparison_group_sap(gid, sap),
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    )

                    test_rows = []
                    for lab in test_labs:
                        matching_test = next(
                            (t for t in selected_tests.values() if t.sap_code == sap_code and t.test_lab_number == lab),
                            None,
                        )
                        voltage_display = f"{matching_test.voltage}V" if matching_test and matching_test.voltage else "N/A"
                        notes_display = (matching_test.notes or "No notes") if matching_test else "From config"
                        if len(notes_display) > 35:
                            notes_display = notes_display[:32] + "..."

                        test_rows.append(
                            ft.Row(
                                controls=[
                                    ft.Text(
                                        f"   └ Test {lab}: {voltage_display} | {notes_display}",
                                        size=11,
                                        color=self._color('on_surface', '#1f2933'),
                                    ),
                                    ft.IconButton(
                                        icon=ft.Icons.CANCEL,
                                        icon_color=self._color('error', '#e65100'),
                                        tooltip=f"Remove test {lab} from {group_id}",
                                        on_click=lambda e, gid=group_id, sap=sap_code, lab=lab: self._handle_remove_comparison_group_test(gid, sap, lab),
                                    ),
                                ],
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                            )
                        )

                    sap_entries.append(ft.Column(controls=[sap_row_header, *test_rows], spacing=2))

                group_controls.append(
                    ft.Container(
                        content=ft.Column(controls=[group_header, *sap_entries], spacing=4),
                        padding=ft.padding.symmetric(horizontal=10, vertical=6),
                        bgcolor=self._color('warning_container', '#fff3e0'),
                        border_radius=6,
                    )
                )

        legacy_controls: List[ft.Control] = []
        if has_legacy:
            for sap_code in sorted(state.selected_comparison_saps):
                selected_test_labs = sorted(state.selected_comparison_test_labs.get(sap_code, set()))
                sap_header = ft.Row(
                    controls=[
                        ft.Text(
                            f"{sap_code}: {len(selected_test_labs) if selected_test_labs else 'All'} test(s)",
                            size=13,
                            color=self._color('warning', '#f57c00'),
                            weight=ft.FontWeight.W_500,
                        ),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_FOREVER,
                            icon_color=self._color('error', '#c62828'),
                            tooltip=f"Remove SAP {sap_code} from comparison",
                            on_click=lambda e, sap=sap_code: self._handle_remove_comparison_sap(sap),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                )

                if selected_test_labs:
                    test_rows = [
                        ft.Row(
                            controls=[
                                ft.Text(
                                    f"• Test {lab}",
                                    size=12,
                                    color=self._color('warning', '#f57c00')
                                ),
                                
                                ft.IconButton(
                                    icon=ft.Icons.CANCEL,
                                    icon_color=self._color('error', '#e65100'),
                                    tooltip=f"Remove test {lab} from comparison",
                                    on_click=lambda e, sap=sap_code, lab=lab: self._handle_remove_comparison_test(sap, lab),
                                ),
                            ],
                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        )
                        for lab in selected_test_labs
                    ]
                else:
                    test_rows = [
                        ft.Text(
                            "• All selected tests included",
                            size=12,
                            color=self._color('warning', '#f57c00')
                        )
                    ]

                legacy_controls.append(
                    ft.Container(
                        content=ft.Column(controls=[sap_header, *test_rows], spacing=3),
                        padding=ft.padding.symmetric(horizontal=10, vertical=6),
                        bgcolor=self._color('warning_container', '#fff8e1'),
                        border_radius=6,
                    )
                )

        summary_controls: List[ft.Control] = []
        if group_controls:
            summary_controls.append(
                ft.Text(
                    "✅ Comparison groups configured",
                    size=14,
                    color=self._color('warning', '#f57c00')
                )
            )
            summary_controls.extend(group_controls)
        if legacy_controls:
            summary_controls.append(
                ft.Text(
                    "✅ Legacy comparison selection active",
                    size=14,
                    color=self._color('warning', '#f57c00')
                )
            )
            summary_controls.extend(legacy_controls)
        if not summary_controls:
            summary_controls.append(
                ft.Text(
                    "❌ No comparison configurations selected",
                    size=14,
                    color=self._color('text_muted', '#5f6b7a')
                )
            )

        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text(
                        "📈 COMPARISON SHEET",
                        size=18,
                        color=self._color('warning', '#fb8c00'),
                        weight=ft.FontWeight.BOLD
                    ),
                    *summary_controls,
                ],
                spacing=6,
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('warning_container', '#fff3e0'),
            border_radius=8,
            border=ft.border.all(1, self._color('outline', '#ffe0b2')),
        )

    def _build_data_flow_section(self, state) -> ft.Control:
        selected_tests = state.selected_tests
        has_comparison_groups = bool(state.comparison_groups)
        has_traditional_comparison = state.include_comparison and bool(state.selected_comparison_saps)

        accent_color = self._color('secondary', '#6a1b9a')
        muted_color = self._color('text_muted', '#5f6b7a')

        summary_lines = [
            ft.Text(
                f"Step 1: Selected {len(selected_tests)} test(s) → Performance Sheet (all tests)",
                size=12,
                color=accent_color,
            ),
            ft.Text(
                f"Step 2: Selected {len(state.selected_noise_saps)} noise SAP(s) → Noise Analysis",
                size=12,
                color=accent_color,
            ),
        ]

        if has_comparison_groups:
            summary_lines.append(
                ft.Text(
                    f"Step 2: {len(state.comparison_groups)} comparison group(s) → Comparison Sheet",
                    size=12,
                    color=accent_color,
                )
            )
        elif has_traditional_comparison:
            summary_lines.append(
                ft.Text(
                    f"Step 2: {len(state.selected_comparison_saps)} comparison SAP(s) → Comparison Sheet",
                    size=12,
                    color=accent_color,
                )
            )
        else:
            summary_lines.append(
                ft.Text(
                    "Step 2: No comparison configured → No comparison sheet",
                    size=12,
                    color=muted_color,
                )
            )

        return ft.Container(
            content=ft.Column(
                controls=[
                    ft.Text(
                        "📋 DATA FLOW SUMMARY",
                        size=16,
                        color=accent_color,
                        weight=ft.FontWeight.BOLD,
                    ),
                    *summary_lines,
                ],
                spacing=4,
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('secondary_container', '#f3e5f5'),
            border_radius=8,
            border=ft.border.all(1, self._color('outline_variant', '#ce93d8')),
        )

    def _handle_remove_performance_test(self, test_id: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        removed = self.parent_gui.state_manager.remove_selected_test(test_id)
        if removed:
            self._after_selection_change()

    def _handle_remove_performance_sap(self, sap_code: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        removed_count = self.parent_gui.state_manager.remove_tests_for_sap(sap_code)
        if removed_count:
            self._after_selection_change()

    def _handle_remove_noise_sap(self, sap_code: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_noise_selection(sap_code):
            self._after_selection_change()

    def _handle_remove_noise_test(self, sap_code: str, test_lab: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_noise_selection(sap_code, test_lab):
            self._after_selection_change()

    def _handle_remove_comparison_sap(self, sap_code: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_comparison_selection(sap_code):
            self._after_selection_change()

    def _handle_remove_comparison_test(self, sap_code: str, test_lab: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_comparison_selection(sap_code, test_lab):
            self._after_selection_change()

    def _handle_remove_comparison_group(self, group_id: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_comparison_group_entry(group_id):
            self._after_selection_change()

    def _handle_remove_comparison_group_sap(self, group_id: str, sap_code: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_comparison_group_entry(group_id, sap_code=sap_code):
            self._after_selection_change()

    def _handle_remove_comparison_group_test(self, group_id: str, sap_code: str, test_lab: str):
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            return
        if self.parent_gui.state_manager.remove_comparison_group_entry(group_id, sap_code=sap_code, test_lab=test_lab):
            self._after_selection_change()

    def _after_selection_change(self):
        try:
            if self.parent_gui and hasattr(self.parent_gui, 'event_handlers') and self.parent_gui.event_handlers:
                self.parent_gui.event_handlers.on_apply_config_selection()
            if self.parent_gui and hasattr(self.parent_gui, 'workflow_manager') and self.parent_gui.workflow_manager:
                # Refresh both config and generate tabs to reflect latest state
                self.parent_gui.workflow_manager.refresh_tab('config')
            self.refresh_content()
            if self.parent_gui:
                self.parent_gui._safe_page_update()
        except Exception as exc:
            logger.error(f"❌ Error after selection change: {exc}")
    
    def _build_performance_section(self, selected_tests) -> ft.Control:
        """Build the Performance sheet section (always included)"""
        logger.info(f"🔍 Building performance section with {len(selected_tests)} tests")
        logger.info(f"🔍 Performance section test details: {[(k, v.sap_code) for k, v in list(selected_tests.items())[:3]]}")
        
        # Group tests by SAP code
        sap_groups = {}
        for test in selected_tests.values():
            sap_code = test.sap_code or "Unknown SAP"
            if sap_code not in sap_groups:
                sap_groups[sap_code] = []
            sap_groups[sap_code].append(test.test_lab_number)
        
        logger.info(f"🔍 Performance section: {len(sap_groups)} SAP groups found")
        logger.info(f"🔍 Performance section SAP groups: {dict(list(sap_groups.items())[:3])}")
        
        # Build SAP code summary
        sap_summaries = []
        for sap_code, test_numbers in sap_groups.items():
            sap_text = ft.Text(
                f"  • {sap_code}: {len(test_numbers)} test(s) ({', '.join(sorted(test_numbers))})",
                size=12,
                color=self._color('on_surface', '#1f2933'),
            )
            sap_summaries.append(sap_text)
        
        content_items = [
            ft.Row([
                ft.Icon(ft.Icons.ASSESSMENT, color=self._color('primary', '#1976d2'), size=16),
                ft.Text(
                    "Performance Sheet",
                    weight=ft.FontWeight.W_500,
                    color=self._color('primary', '#1976d2'),
                    size=14,
                ),
            ], spacing=5),
            ft.Text(
                f"All {len(selected_tests)} selected tests will be included",
                size=12,
                color=self._color('text_muted', '#5f6b7a'),
            ),
        ]
        
        # Add SAP summaries
        content_items.extend(sap_summaries)
        
        # Enhanced container with visibility settings
        return ft.Container(
            content=ft.Column(
                controls=content_items,
                spacing=5,
                tight=True,
                visible=True      # Explicitly visible
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('primary_container', '#e3f2fd'),
            border_radius=5,
            border=ft.border.all(1, self._color('outline', '#90caf9')),
            expand=False,
            visible=True,         # Container explicitly visible
            height=None,          # Natural height
            width=None            # Natural width
        )
    
    def _build_noise_section(self, selected_noise_saps, selected_tests) -> ft.Control:
        """Build the Noise section"""
        logger.info(f"🔍 Building noise section with SAPs: {selected_noise_saps}")
        logger.info(f"🔍 Noise section - selected_tests type: {type(selected_tests)}, count: {len(selected_tests)}")
        
        # Get state to access noise test selections
        if not self.parent_gui or not hasattr(self.parent_gui, 'state_manager'):
            logger.warning("🔍 _build_noise_section: No state manager found")
            return ft.Text("❌ No state manager found", color=self._color('error', '#c62828'), size=14)
        
        state = self.parent_gui.state_manager.state
        
        if not selected_noise_saps:
            logger.info("🔍 Noise section: No SAP codes selected, showing disabled state")
            # Enhanced disabled noise section with visibility
            return ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Row([
                            ft.Icon(ft.Icons.VOLUME_OFF, color=self._color('text_disabled', '#9e9e9e'), size=16),
                            ft.Text(
                                "Noise Analysis",
                                weight=ft.FontWeight.W_500,
                                color=self._color('text_disabled', '#9e9e9e'),
                                size=14,
                            )
                        ], spacing=5),
                        ft.Text(
                            "Disabled - No SAP codes selected",
                            size=12,
                            color=self._color('text_disabled', '#9e9e9e'),
                        )
                    ],
                    spacing=3,
                    tight=True,
                    visible=True      # Explicitly visible
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('surface_container_low', '#f5f5f5'),
                border_radius=5,
                border=ft.border.all(1, self._color('outline_variant', '#cfcfcf')),
                expand=False,
                visible=True,         # Container explicitly visible
                height=None,          # Natural height
                width=None            # Natural width
            )
        
        # Find noise tests using the correct state data
        # Don't use selected_tests (performance tests) - use selected_noise_test_labs instead
        logger.info(f"🔍 Noise section: Using selected_noise_test_labs: {state.selected_noise_test_labs}")
        
        # Build SAP code summary for noise
        sap_summaries = []
        for sap_code in selected_noise_saps:
            matching_tests = list(state.selected_noise_test_labs.get(sap_code, set()))
            logger.info(f"🔍 Noise section: SAP {sap_code} has {len(matching_tests)} noise tests: {matching_tests}")
            sap_text = ft.Text(
                f"  • {sap_code}: {len(matching_tests)} test(s) ({', '.join(sorted(matching_tests)) if matching_tests else 'None selected'})",
                size=12,
                color=self._color('on_surface', '#1f2933'),
            )
            sap_summaries.append(sap_text)
        
        content_items = [
            ft.Row([
                ft.Icon(ft.Icons.VOLUME_UP, color=self._color('success', '#388e3c'), size=16),
                ft.Text(
                    "Noise Analysis",
                    weight=ft.FontWeight.W_500,
                    color=self._color('success', '#388e3c'),
                    size=14,
                )
            ], spacing=5),
            ft.Text(
                f"Enabled for {len(selected_noise_saps)} SAP code(s)",
                size=12,
                color=self._color('text_muted', '#5f6b7a'),
            ),
        ]
        content_items.extend(sap_summaries)
        
        # Enhanced noise section container with visibility
        return ft.Container(
            content=ft.Column(
                controls=content_items,
                spacing=5,
                tight=True,
                visible=True      # Explicitly visible
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('success_container', '#e8f5e8'),
            border_radius=5,
            border=ft.border.all(1, self._color('outline', '#a5d6a7')),
            expand=False,
            visible=True,         # Container explicitly visible
            height=None,          # Natural height
            width=None            # Natural width
        )
    
    def _build_comparison_section(self, selected_comparison_saps, selected_tests) -> ft.Control:
        """Build the Comparison section"""
        if not selected_comparison_saps:
            # Enhanced disabled comparison section with visibility
            return ft.Container(
                content=ft.Column(
                    controls=[
                        ft.Row([
                            ft.Icon(ft.Icons.COMPARE, color=self._color('text_disabled', '#9e9e9e'), size=16),
                            ft.Text(
                                "Comparison Sheet",
                                weight=ft.FontWeight.W_500,
                                color=self._color('text_disabled', '#9e9e9e'),
                                size=14,
                            )
                        ], spacing=5),
                        ft.Text(
                            "Disabled - No SAP codes selected",
                            size=12,
                            color=self._color('text_disabled', '#9e9e9e'),
                        )
                    ],
                    spacing=3,
                    tight=True,
                    visible=True      # Explicitly visible
                ),
                padding=ft.padding.all(10),
                bgcolor=self._color('surface_container_low', '#f5f5f5'),
                border_radius=5,
                border=ft.border.all(1, self._color('outline_variant', '#cfcfcf')),
                expand=False,
                visible=True,         # Container explicitly visible
                height=None,          # Natural height
                width=None            # Natural width
            )
        
        # Find tests that match selected comparison SAP codes
        comparison_tests = {k: v for k, v in selected_tests.items() if v.sap_code in selected_comparison_saps}
        
        # Build SAP code summary for comparison
        sap_summaries = []
        for sap_code in selected_comparison_saps:
            matching_tests = [test.test_lab_number for test in comparison_tests.values() if test.sap_code == sap_code]
            sap_text = ft.Text(
                f"  • {sap_code}: {len(matching_tests)} test(s) ({', '.join(sorted(matching_tests)) if matching_tests else 'None selected'})",
                size=12,
                color=self._color('on_surface', '#1f2933'),
            )
            sap_summaries.append(sap_text)
        
        content_items = [
            ft.Row([
                ft.Icon(ft.Icons.COMPARE_ARROWS, color=self._color('warning', '#fbc02d'), size=16),
                ft.Text(
                    "Comparison Sheet",
                    weight=ft.FontWeight.W_500,
                    color=self._color('warning', '#fbc02d'),
                    size=14,
                )
            ], spacing=5),
            ft.Text(
                f"Enabled for {len(selected_comparison_saps)} SAP code(s)",
                size=12,
                color=self._color('text_muted', '#5f6b7a'),
            ),
        ]
        content_items.extend(sap_summaries)
        
        # Enhanced comparison section container with visibility
        return ft.Container(
            content=ft.Column(
                controls=content_items,
                spacing=5,
                tight=True,
                visible=True      # Explicitly visible
            ),
            padding=ft.padding.all(10),
            bgcolor=self._color('warning_container', '#fffde7'),
            border_radius=5,
            border=ft.border.all(1, self._color('outline', '#ffe082')),
            expand=False,
            visible=True,         # Container explicitly visible
            height=None,          # Natural height
            width=None            # Natural width
        )
    

    

    def refresh_content(self):
        """Enhanced method to refresh the Generate tab content when configuration changes"""
        import time
        import traceback
        
        try:
            logger.info("🔄 Starting Enhanced Generate tab refresh...")
            
            # Validate prerequisites
            if not self._validate_refresh_prerequisites():
                return False
            
            # Method 1: Try in-place content update (faster)
            if self._try_in_place_refresh():
                logger.info("✅ In-place refresh successful")
                return True
            
            # Method 2: Full tab content replacement (more reliable)
            logger.info("🔄 Falling back to full tab refresh...")
            return self._perform_full_tab_refresh()
            
        except Exception as e:
            logger.error(f"❌ Error in refresh_content: {e}")
            logger.error(f"❌ Traceback: {traceback.format_exc()}")
            return False
    
    def _validate_refresh_prerequisites(self) -> bool:
        """Validate that all necessary components are available for refresh"""
        if not (self.parent_gui and hasattr(self.parent_gui, 'state_manager')):
            logger.warning("❌ Cannot refresh - no state manager available")
            return False
        
        if not (self.parent_gui and hasattr(self.parent_gui, 'tabs') and 
                self.parent_gui.tabs and len(self.parent_gui.tabs.tabs) > 3):
            logger.warning("❌ Cannot refresh - tabs not properly initialized") 
            return False
        
        return True
    
    def _try_in_place_refresh(self) -> bool:
        """Try to refresh just the summary content without replacing the entire tab"""
        try:
            import time
            
            logger.info("🔧 Attempting in-place content refresh...")
            
            # Check if we have a valid summary container reference
            if not hasattr(self, 'summary_container') or not self.summary_container:
                logger.info("� No summary container reference, skipping in-place refresh")
                return False
            
            # Build new summary content
            new_summary_content = self._build_tests_summary()
            
            # Try to update the summary container content with proper type checking
            if (hasattr(self.summary_container, 'content') and 
                self.summary_container.content and
                hasattr(self.summary_container.content, 'controls')):
                
                # Find the summary content in the controls list with safety checks
                try:
                    # Cast to Any to avoid type checking issues
                    content_obj: Any = self.summary_container.content
                    controls = getattr(content_obj, 'controls', None)
                    if controls:  # Ensure controls is not None
                        for i, control in enumerate(controls):
                            if (hasattr(control, 'value') and isinstance(control.value, str) and 
                                "Selected Tests Summary" in control.value):
                                # Found the summary header, replace the next control (the actual summary)
                                if i + 1 < len(controls):
                                    controls[i + 1] = new_summary_content
                                    
                                    # Force visibility and update
                                    new_summary_content.visible = True
                                    self.summary_container.visible = True
                                    
                                    # Perform strategic updates with safety checks
                                    if (self.parent_gui and 
                                        hasattr(self.parent_gui, '_safe_page_update')):
                                        self.parent_gui._safe_page_update()
                                        time.sleep(0.02)
                                        self.parent_gui._safe_page_update()
                                    
                                    logger.info("✅ In-place summary refresh completed")
                                    return True
                except (AttributeError, TypeError) as e:
                    logger.warning(f"⚠️ Error accessing controls: {e}")
                    return False
            
            logger.info("📋 In-place refresh not possible with current structure")
            return False
            
        except Exception as e:
            logger.warning(f"⚠️ In-place refresh failed: {e}")
            return False
    
    def _perform_full_tab_refresh(self) -> bool:
        """Perform a complete tab content replacement with enhanced timing and error handling"""
        try:
            import time
            
            logger.info("🔧 Starting full tab content replacement...")
            
            # Validate parent_gui and required attributes
            if not (self.parent_gui and hasattr(self.parent_gui, 'tabs') and
                    self.parent_gui.tabs and hasattr(self.parent_gui.tabs, 'tabs') and
                    len(self.parent_gui.tabs.tabs) > 3):
                logger.error("❌ Cannot refresh - tabs not properly initialized")
                return False
            
            # Step 1: Get the current tab reference
            current_tab = self.parent_gui.tabs.tabs[3]
            original_selected_index = self.parent_gui.tabs.selected_index
            
            # Step 2: Build completely fresh content
            logger.info("🔧 Building fresh content...")
            new_content = self.get_tab_content()
            
            # Step 3: Enhanced content replacement strategy
            logger.info("🔧 Replacing tab content...")
            
            # Store current visibility state
            was_visible = getattr(current_tab.content, 'visible', True) if current_tab.content else True
            
            # Method A: Direct replacement with forced visibility
            current_tab.content = new_content
            new_content.visible = True
            
            # Step 4: Efficient single-update sequence
            logger.info("🔧 Executing optimized update...")
            
            # Single comprehensive update - no excessive strategies or sleeps
            if hasattr(self.parent_gui, '_safe_page_update'):
                self.parent_gui._safe_page_update()
            elif hasattr(self.parent_gui, 'page') and hasattr(self.parent_gui.page, 'update'):
                try:
                    self.parent_gui.page.update()
                except Exception as page_error:
                    logger.warning(f"⚠️ Page update failed: {page_error}")
            
            # Ensure visibility
            new_content.visible = True
            current_tab.visible = True
            
            # Step 5: Update our container reference
            self._update_summary_container_reference(new_content)
            
            logger.info("✅ Full tab refresh completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"❌ Full tab refresh failed: {e}")
            return False
    
    def _update_summary_container_reference(self, new_content):
        """Update the summary container reference after content replacement"""
        try:
            # Search for the new summary container in the fresh content
            def find_summary_container(control, depth=0):
                if depth > 5:  # Prevent infinite recursion
                    return None
                
                # Check if this is our summary container
                if (hasattr(control, 'content') and 
                    hasattr(control.content, 'controls') and 
                    len(control.content.controls) >= 2):
                    
                    first_control = control.content.controls[0]
                    if (hasattr(first_control, 'value') and 
                        isinstance(first_control.value, str) and
                        "Selected Tests Summary" in first_control.value):
                        return control
                
                # Recursively search children
                if hasattr(control, 'content'):
                    if hasattr(control.content, 'controls'):
                        for child in control.content.controls:
                            result = find_summary_container(child, depth + 1)
                            if result:
                                return result
                    elif hasattr(control.content, 'content'):
                        return find_summary_container(control.content, depth + 1)
                
                return None
            
            found_container = find_summary_container(new_content)
            if found_container:
                self.summary_container = found_container
                logger.info("📋 Updated summary container reference")
            else:
                logger.info("📋 Summary container not found in new content")
                
        except Exception as e:
            logger.warning(f"⚠️ Could not update summary container reference: {e}")
    
    def on_tab_visible(self):
        """Called when the Generate tab becomes visible - ensures content is up-to-date"""
        try:
            logger.info("👁️ Generate tab became visible - checking if refresh needed...")
            
            if not self._validate_refresh_prerequisites():
                return
            
            # Validate parent_gui and tabs
            if not (self.parent_gui and hasattr(self.parent_gui, 'tabs') and
                    self.parent_gui.tabs and hasattr(self.parent_gui.tabs, 'tabs') and
                    len(self.parent_gui.tabs.tabs) > 3):
                logger.warning("❌ Cannot check tab visibility - tabs not properly initialized")
                return
            
            # Check if we have valid content
            current_tab = self.parent_gui.tabs.tabs[3]
            if not current_tab.content:
                logger.info("🔄 No content found, triggering refresh...")
                self.refresh_content()
                return
            
            # Check if content seems outdated (simple heuristic)
            if (self.parent_gui and hasattr(self.parent_gui, 'state_manager') and
                self.parent_gui.state_manager):
                state = self.parent_gui.state_manager.state
                if hasattr(self, '_last_known_test_count'):
                    if self._last_known_test_count != len(state.selected_tests):
                        logger.info("🔄 Test count changed, triggering refresh...")
                        self.refresh_content()
                        return
                
                # Store current test count for future comparisons
                self._last_known_test_count = len(state.selected_tests)
            
            logger.info("👁️ Content appears current, no refresh needed")
            
        except Exception as e:
            logger.warning(f"⚠️ Error in on_tab_visible: {e}")

    def force_content_visibility(self):
        """Force all content to be visible and updated - for cases where normal refresh isn't enough"""
        try:
            logger.info("🔧 Forcing content visibility...")
            
            if not (self.parent_gui and hasattr(self.parent_gui, 'tabs') and 
                    self.parent_gui.tabs and len(self.parent_gui.tabs.tabs) > 3):
                return False
            
            current_tab = self.parent_gui.tabs.tabs[3]
            if current_tab.content:
                # Force visibility on all levels
                current_tab.visible = True
                current_tab.content.visible = True
                
                # If it's a container, force visibility on its content too
                if hasattr(current_tab.content, 'content') and current_tab.content.content:
                    current_tab.content.content.visible = True
                    
                    # If the content has controls, make them visible too
                    if hasattr(current_tab.content.content, 'controls'):
                        for control in current_tab.content.content.controls:
                            if hasattr(control, 'visible'):
                                control.visible = True
                
                # Apply multiple updates
                if hasattr(self.parent_gui, 'page') and hasattr(self.parent_gui.page, 'update'):
                    self.parent_gui.page.update()
                
                self.parent_gui._safe_page_update()
                
                logger.info("✅ Content visibility forced")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"❌ Error forcing content visibility: {e}")
            return False

    # Consolidated aliases for backward compatibility - all delegate to refresh_content
    def force_refresh(self): return self.refresh_content()
    def update_tests_summary(self): return self.refresh_content()
    def test_ui_refresh_capability(self): return self.refresh_content()

