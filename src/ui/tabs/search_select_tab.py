"""
Search and Select Tab - Second tab for finding and selecting tests
"""
import flet as ft
from ..components.base import BaseTab


class SearchSelectTab(BaseTab):
    """Tab for searching and selecting tests"""
    
    def __init__(self, parent_gui=None):
        super().__init__(parent_gui)
        self.tab_name = "2. Search & Select"
        self.tab_icon = ft.Icons.SEARCH
    
    def get_tab_content(self) -> ft.Control:
        """Build the search and select tab content"""
        # Get progress indicators from parent
        step1_progress = None
        step1_status = None
        if self.parent_gui and hasattr(self.parent_gui, 'progress_indicators'):
            step1_progress, step1_status = self.parent_gui.progress_indicators.get_indicators_for_step(1)
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.Text("Search & Select Tests", size=20, weight=ft.FontWeight.BOLD),
                    step1_progress or ft.Container(),
                    step1_status or ft.Container(),
                ], spacing=10),
                
                ft.Text("Enter SAP codes or test numbers to find related tests in the registry.", 
                       color="grey", size=14),
                ft.Divider(),
                
                ft.Row([
                    self.parent_gui.search_input_field if self.parent_gui else ft.TextField(label="Search"),
                    self.parent_gui.search_button if self.parent_gui else ft.ElevatedButton("Search")
                ], alignment=ft.MainAxisAlignment.START),
                
                # SAP Navigation (appears when multiple SAP codes found)
                self.parent_gui.sap_navigation_container if self.parent_gui else ft.Container(),
                
                ft.Text("Search Results:", weight=ft.FontWeight.BOLD),
                self.parent_gui.results_filters_row if self.parent_gui else ft.Container(),
                ft.Container(
                    content=self.parent_gui.results_area if self.parent_gui else ft.Column([ft.Text("No results")]),
                    border=ft.border.all(1, "grey"),
                    border_radius=ft.border_radius.all(5),
                    padding=10,
                    expand=True,
                ),
                
                # Selected count display only (no action buttons needed)
                self.parent_gui.selected_count_text if self.parent_gui else ft.Container(),
                
            ], spacing=15),
            padding=ft.padding.all(20),
            expand=True
        )

