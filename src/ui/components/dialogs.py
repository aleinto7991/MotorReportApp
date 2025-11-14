"""
Dialog components for the GUI
"""
import flet as ft
import logging
from typing import List, Callable, Optional
from .base import BaseComponent

logger = logging.getLogger(__name__)


class SAPSelectionDialog(BaseComponent):
    """Dialog for selecting which SAP codes to compare"""
    
    def __init__(self, parent_gui=None, sap_codes: Optional[List[str]] = None, 
                 tests_to_process: Optional[List] = None, on_confirm: Optional[Callable] = None, 
                 on_cancel: Optional[Callable] = None):
        super().__init__(parent_gui)
        self.sap_codes = sap_codes or []
        self.tests_to_process = tests_to_process or []
        self.on_confirm = on_confirm
        self.on_cancel = on_cancel
        self.sap_checkboxes = []
        
    def build(self) -> ft.AlertDialog:
        """Build the SAP selection dialog"""
        # Create checkboxes for each SAP code
        self.sap_checkboxes = []
        for sap in self.sap_codes:
            # Count how many tests have this SAP code
            test_count = sum(1 for test in self.tests_to_process if test.sap_code == sap)
            checkbox = ft.Checkbox(
                label=f"{sap} ({test_count} test{'s' if test_count != 1 else ''})",
                value=True,  # Default to selected
                data=sap
            )
            self.sap_checkboxes.append(checkbox)

        dialog_content = ft.Column([
            ft.Text("Select SAP codes to include in comparison:", size=16, weight=ft.FontWeight.BOLD),
            ft.Text("Multiple SAP codes found in your selected tests. Choose which ones to compare:"),
            ft.Container(
                content=ft.Column(self.sap_checkboxes, spacing=5, scroll=ft.ScrollMode.AUTO),
                height=200,  # Limit height so it doesn't overflow
            ),
            ft.Row([
                ft.TextButton("Cancel", on_click=self._on_cancel),
                ft.ElevatedButton("Generate Report", on_click=self._on_confirm)
            ], alignment=ft.MainAxisAlignment.END)
        ], spacing=10, scroll=ft.ScrollMode.AUTO)

        return ft.AlertDialog(
            title=ft.Text("SAP Code Selection"),
            content=dialog_content,
            actions_alignment=ft.MainAxisAlignment.END,
        )
    
    def _on_confirm(self, e):
        """Handle dialog confirmation"""
        selected_saps = [cb.data for cb in self.sap_checkboxes if cb.value]
        logger.info(f"User selected SAPs for comparison: {selected_saps}")
        
        if not selected_saps:
            if self.parent_gui and hasattr(self.parent_gui, 'update_status'):
                self.parent_gui.update_status("Error: Please select at least one SAP code for comparison.", color="red")
            return
        
        # Close dialog first
        self._close_dialog()
        
        # Call the confirmation callback
        if self.on_confirm:
            self.on_confirm(selected_saps)
    
    def _on_cancel(self, e):
        """Handle dialog cancellation"""
        logger.info("User cancelled SAP selection")
        self._close_dialog()
        
        if self.on_cancel:
            self.on_cancel()
    
    def _close_dialog(self):
        """Close the dialog"""
        if self.parent_gui and hasattr(self.parent_gui, 'page'):
            for overlay in self.parent_gui.page.overlay[:]:
                if isinstance(overlay, ft.AlertDialog):
                    overlay.open = False
                    self.parent_gui.page.overlay.remove(overlay)
            self.safe_update()
    
    def show(self):
        """Show the dialog"""
        if self.parent_gui and hasattr(self.parent_gui, 'page'):
            dialog = self.build()
            self.parent_gui.page.overlay.append(dialog)
            dialog.open = True
            self.safe_update()
            logger.info("SAP selection dialog displayed")


class NotesDialog(BaseComponent):
    """Dialog for displaying full test notes"""
    
    def __init__(self, parent_gui=None, test=None):
        super().__init__(parent_gui)
        self.test = test
    
    def build(self) -> ft.AlertDialog:
        """Build the notes dialog"""
        if not self.test:
            return ft.AlertDialog(
                title=ft.Text("Notes"),
                content=ft.Text("No test information available."),
            )
        
        dialog_content = ft.Column([
            ft.Text(f"Test: {self.test.test_lab_number} | SAP: {self.test.sap_code}", 
                   weight=ft.FontWeight.BOLD),
            ft.Container(
                content=ft.Text(self.test.notes or "No notes available", selectable=True, size=14),
                height=300,
                border=ft.border.all(1, "grey"),
                border_radius=5,
                padding=10,
                bgcolor="#f9f9f9"
            ),
            ft.Row([
                ft.ElevatedButton("Close", on_click=self._close_dialog)
            ], alignment=ft.MainAxisAlignment.END)
        ], spacing=10)

        return ft.AlertDialog(
            title=ft.Text("Full Notes"),
            content=dialog_content,
            actions_alignment=ft.MainAxisAlignment.END,
        )
    
    def _close_dialog(self, e):
        """Close the dialog"""
        if self.parent_gui and hasattr(self.parent_gui, 'page'):
            for overlay in self.parent_gui.page.overlay[:]:
                if isinstance(overlay, ft.AlertDialog):
                    overlay.open = False
                    self.parent_gui.page.overlay.remove(overlay)
            self.safe_update()
    
    def show(self):
        """Show the dialog"""
        if self.parent_gui and hasattr(self.parent_gui, 'page'):
            dialog = self.build()
            self.parent_gui.page.overlay.append(dialog)
            dialog.open = True
            self.safe_update()

