"""
Utility functions for the Motor Report GUI
Contains standalone utility functions that don't require circular dependencies.
"""
import flet as ft
import logging
from typing import List, Dict, Callable, Optional
from ...data.models import Test

logger = logging.getLogger(__name__)

ColorResolver = Optional[Callable[[str, str], Optional[str]]]


def _resolve_color(resolver: ColorResolver, token: str, fallback: str) -> str:
    if callable(resolver):
        try:
            resolved = resolver(token, fallback)
            if resolved:
                return resolved
        except Exception as exc:
            logger.debug("display_utils color resolver failed for %s: %s", token, exc)
    return fallback


def create_test_row(
    test: Test,
    is_selected: bool,
    on_checkbox_change: Callable,
    on_row_click: Callable,
    *,
    color_resolver: ColorResolver = None,
) -> ft.Container:
    """Create a test row for display"""
    color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
    notes_value = (test.notes or "").strip()
    truncated_notes = notes_value if len(notes_value) <= 50 else f"{notes_value[:50]}..."
    checkbox = ft.Checkbox(
        value=is_selected,
        data=test,
        on_change=on_checkbox_change
    )
    
    row = ft.Container(
        content=ft.Row([
            checkbox,
            ft.Column([
                ft.Text(f"Test Lab: {test.test_lab_number}", weight=ft.FontWeight.W_500),
                ft.Text(f"SAP Code: {test.sap_code}", size=12, color=color('text_muted', 'grey')),
                ft.Text(f"Voltage: {test.voltage}V", size=12, color=color('text_muted', 'grey')),
                ft.Text(
                    f"Notes: {truncated_notes}" if truncated_notes else "Notes: N/A",
                    size=12,
                    color=color('text_muted', 'grey')
                ),
            ], spacing=2, expand=True),
            ft.IconButton(
                ft.Icons.INFO_OUTLINE,
                tooltip="Click row to toggle selection",
                icon_color=color('primary', 'blue')
            )
        ], alignment=ft.MainAxisAlignment.START),
        padding=ft.padding.all(8),
        margin=ft.margin.symmetric(vertical=2),
        border=ft.border.all(1, color('primary', 'lightblue')) if is_selected else ft.border.all(1, color('outline', 'lightgrey')),
        border_radius=5,
        bgcolor=color('primary_container', '#F0F8FF') if is_selected else color('surface', 'white'),
        on_click=on_row_click
    )
    
    return row


def create_sap_group_header(
    sap_code: str,
    test_count: int,
    *,
    color_resolver: ColorResolver = None,
) -> ft.Container:
    """Create a SAP group header"""
    color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
    return ft.Container(
        content=ft.Row([
            ft.Text(
                f"SAP Code: {sap_code}",
                size=16,
                weight=ft.FontWeight.BOLD,
                color=color('primary', 'blue')
            ),
            ft.Text(
                f"({test_count} test{'s' if test_count != 1 else ''})",
                size=14,
                color=color('text_muted', 'grey')
            )
        ]),
        bgcolor=color('primary_container', '#E3F2FD'),
        padding=ft.padding.all(10),
        border_radius=5,
        margin=ft.margin.only(top=10, bottom=5)
    )


def create_action_buttons(
    selected_count: int,
    on_apply: Callable,
    on_clear: Callable,
    *,
    color_resolver: ColorResolver = None,
) -> ft.Row:
    """Create action buttons for search results"""
    color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
    return ft.Row([
        ft.Text(
            f"{selected_count} test{'s' if selected_count != 1 else ''} selected",
            size=14,
            weight=ft.FontWeight.W_500,
            color=color('success', 'green')
        ),
        ft.ElevatedButton(
            "Apply Selection",
            icon=ft.Icons.CHECK,
            bgcolor=color('success', 'green'),
            color=color('on_success', 'white'),
            on_click=on_apply
        ),
        ft.ElevatedButton(
            "Clear Selection",
            icon=ft.Icons.CLEAR,
            bgcolor=color('warning', 'orange'),
            color=color('on_warning', 'white'),
            on_click=on_clear
        )
    ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, wrap=True)


def group_tests_by_sap(tests: List[Test]) -> Dict[str, List[Test]]:
    """Group tests by SAP code"""
    sap_groups = {}
    for test in tests:
        sap = test.sap_code or "Unknown SAP"
        if sap not in sap_groups:
            sap_groups[sap] = []
        sap_groups[sap].append(test)
    return sap_groups


def create_sap_checkbox(sap_code: str, is_selected: bool, on_change: Callable) -> ft.Checkbox:
    """Create a SAP selection checkbox"""
    return ft.Checkbox(
        label=sap_code,
        value=is_selected,
        data=sap_code,
        on_change=on_change
    )


def create_pagination_controls(
    current_page: int,
    total_pages: int,
    on_previous: Callable,
    on_next: Callable,
    *,
    color_resolver: ColorResolver = None,
) -> ft.Row:
    """Create pagination controls"""
    color = lambda token, fallback: _resolve_color(color_resolver, token, fallback)
    return ft.Row([
        ft.ElevatedButton(
            "Previous",
            icon=ft.Icons.ARROW_BACK,
            disabled=current_page == 0,
            on_click=on_previous
        ),
        ft.Text(
            f"Page {current_page + 1} of {total_pages}",
            color=color('text_muted', 'grey')
        ),
        ft.ElevatedButton(
            "Next",
            icon=ft.Icons.ARROW_FORWARD,
            disabled=current_page >= total_pages - 1,
            on_click=on_next
        )
    ], alignment=ft.MainAxisAlignment.CENTER)

