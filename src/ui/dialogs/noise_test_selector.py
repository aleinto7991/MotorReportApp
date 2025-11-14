"""
Noise Test Selector Dialog - integrates with existing GUI dialog patterns.

This dialog allows users to select which noise tests to include in the report
after validation shows which files are available, missing, or invalid.
"""

import flet as ft
import os
from typing import List, Callable, Optional
import logging

logger = logging.getLogger(__name__)

class NoiseTestSelectorDialog:
    """Dialog for selecting which noise tests to include in the report."""
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.selected_tests = []
        self.dialog: Optional[ft.AlertDialog] = None
        self.checkboxes = []
        
    def show(self, noise_tests, on_confirm: Callable):
        """Show the selection dialog matching existing dialog patterns."""
        from ...validators.noise_test_validator import NoiseTestValidationInfo
        
        # Group tests by status
        valid_tests = [t for t in noise_tests if t.is_valid]
        missing_tests = [t for t in noise_tests if not t.exists]
        invalid_tests = [t for t in noise_tests if t.exists and not t.is_valid]
        
        # Reset checkboxes
        self.checkboxes = []
        
        # Build content - simplified version
        content_controls = []
        
        # Header
        content_controls.append(
            ft.Text("Noise Test Validation Results", 
                   size=16, 
                   weight=ft.FontWeight.BOLD)
        )
        content_controls.append(ft.Divider())
        
        # Summary
        summary_text = f"Found {len(noise_tests)} tests in registry: {len(valid_tests)} valid, {len(missing_tests)} missing, {len(invalid_tests)} invalid"
        content_controls.append(ft.Text(summary_text, size=12))
        content_controls.append(ft.Divider())
        
        # Valid tests section
        if valid_tests:
            content_controls.append(
                ft.Text("Select tests to include:", 
                       size=14, 
                       weight=ft.FontWeight.BOLD)
            )
            
            # Select all controls
            select_all_cb = ft.Checkbox(
                label="Select All",
                value=True,
                on_change=lambda e: self._toggle_all_tests(e.control.value)
            )
            content_controls.append(select_all_cb)
            content_controls.append(ft.Divider())
            
            # Individual test checkboxes
            for test in valid_tests:
                file_info = f" ({test.file_size} bytes)" if test.file_size else ""
                date_info = f" - {test.date}" if test.date else ""
                
                cb = ft.Checkbox(
                    label=f"{test.sap_code} - {test.test_no}{date_info}{file_info}",
                    value=True,  # Default to selected
                    data=test
                )
                self.checkboxes.append(cb)
                content_controls.append(cb)
                
        # Missing tests section
        if missing_tests:
            content_controls.append(ft.Divider())
            content_controls.append(
                ft.Text("Missing test files (will be skipped):", 
                       size=14, 
                       weight=ft.FontWeight.BOLD,
                       color="red")
            )
            
            for test in missing_tests:
                content_controls.append(
                    ft.Text(f"❌ {test.sap_code} - {test.test_no}: {test.error_message}",
                           size=12, color="red")
                )
                
        # Invalid tests section  
        if invalid_tests:
            content_controls.append(ft.Divider())
            content_controls.append(
                ft.Text("Invalid test files (will be skipped):", 
                       size=14, 
                       weight=ft.FontWeight.BOLD,
                       color="orange")
            )
            
            for test in invalid_tests:
                content_controls.append(
                    ft.Text(f"⚠️ {test.sap_code} - {test.test_no}: {test.error_message}",
                           size=12, color="orange")
                )
        
        # If no valid tests, show message
        if not valid_tests:
            content_controls.append(
                ft.Text("No valid noise test files found. Report will be generated without noise data.", 
                       size=14, color="blue")
            )
        
        # Create scrollable content
        content = ft.Container(
            content=ft.Column(
                content_controls,
                scroll=ft.ScrollMode.AUTO,
                spacing=5
            ),
            height=400,
            width=500
        )
        
        def handle_confirm(e):
            """Handle confirm button click."""
            # Collect selected tests
            self.selected_tests = []
            for cb in self.checkboxes:
                if cb.value and cb.data:
                    self.selected_tests.append(cb.data)
            
            logger.info(f"User selected {len(self.selected_tests)} noise tests")
            
            self.page.dialog = None  # type: ignore
            self.page.update()
            on_confirm(self.selected_tests)
            
        def handle_cancel(e):
            """Handle cancel button click."""
            logger.info("User cancelled noise test selection")
            self.page.dialog = None  # type: ignore
            self.page.update()
            on_confirm([])  # Empty list means cancelled
        
        # Create dialog
        self.dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Select Noise Tests"),
            content=content,
            actions=[
                ft.TextButton("Cancel", on_click=handle_cancel),
                ft.ElevatedButton("Continue", on_click=handle_confirm),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
            open=True
        )
        
        self.page.dialog = self.dialog  # type: ignore
        self.page.update()
        
    def _toggle_all_tests(self, select_all: bool):
        """Toggle all test checkboxes."""
        for cb in self.checkboxes:
            cb.value = select_all
        self.page.update()

