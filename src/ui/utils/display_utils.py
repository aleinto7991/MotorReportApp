"""
Utility functions for the Motor Report GUI
Contains standalone utility functions that don't require circular dependencies.
"""
import flet as ft
import logging
from typing import List, Dict, Callable
from ...data.models import Test

logger = logging.getLogger(__name__)


def create_test_row(test: Test, is_selected: bool, on_checkbox_change: Callable, on_row_click: Callable) -> ft.Container:
    """Create a test row for display"""
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
                ft.Text(f"SAP Code: {test.sap_code}", size=12, color="grey"),
                ft.Text(f"Voltage: {test.voltage}V", size=12, color="grey"),
                ft.Text(f"Notes: {test.notes[:50]}..." if len(test.notes) > 50 else f"Notes: {test.notes}", size=12, color="grey"),
            ], spacing=2, expand=True),
            ft.IconButton(
                ft.Icons.INFO_OUTLINE,
                tooltip="Click row to toggle selection",
                icon_color="blue"
            )
        ], alignment=ft.MainAxisAlignment.START),
        padding=ft.padding.all(8),
        margin=ft.margin.symmetric(vertical=2),
        border=ft.border.all(1, "lightblue") if is_selected else ft.border.all(1, "lightgrey"),
        border_radius=5,
        bgcolor="#F0F8FF" if is_selected else "white",
        on_click=on_row_click
    )
    
    return row


def create_sap_group_header(sap_code: str, test_count: int) -> ft.Container:
    """Create a SAP group header"""
    return ft.Container(
        content=ft.Row([
            ft.Text(
                f"SAP Code: {sap_code}",
                size=16,
                weight=ft.FontWeight.BOLD,
                color="blue"
            ),
            ft.Text(
                f"({test_count} test{'s' if test_count != 1 else ''})",
                size=14,
                color="grey"
            )
        ]),
        bgcolor="#E3F2FD",
        padding=ft.padding.all(10),
        border_radius=5,
        margin=ft.margin.only(top=10, bottom=5)
    )


def create_action_buttons(selected_count: int, on_apply: Callable, on_clear: Callable) -> ft.Row:
    """Create action buttons for search results"""
    return ft.Row([
        ft.Text(
            f"{selected_count} test{'s' if selected_count != 1 else ''} selected",
            size=14,
            weight=ft.FontWeight.W_500,
            color="green"
        ),
        ft.ElevatedButton(
            "Apply Selection",
            icon=ft.Icons.CHECK,
            bgcolor="green",
            color="white",
            on_click=on_apply
        ),
        ft.ElevatedButton(
            "Clear Selection",
            icon=ft.Icons.CLEAR,
            bgcolor="orange",
            color="white",
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


def create_pagination_controls(current_page: int, total_pages: int, on_previous: Callable, on_next: Callable) -> ft.Row:
    """Create pagination controls"""
    return ft.Row([
        ft.ElevatedButton(
            "Previous",
            icon=ft.Icons.ARROW_BACK,
            disabled=current_page == 0,
            on_click=on_previous
        ),
        ft.Text(f"Page {current_page + 1} of {total_pages}"),
        ft.ElevatedButton(
            "Next",
            icon=ft.Icons.ARROW_FORWARD,
            disabled=current_page >= total_pages - 1,
            on_click=on_next
        )
    ], alignment=ft.MainAxisAlignment.CENTER)

