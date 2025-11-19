"""
Event handlers for the Motor Report GUI
Contains all event handling logic separated from the main GUI class.
"""
import flet as ft
import logging
import threading  # Keep for Lock
import time
import os
import shutil
import subprocess
import platform
import datetime
import tempfile
from typing import Optional, TYPE_CHECKING

from .search_controller import SearchController
from .selection_controller import SelectionController
from .report_generation_controller import ReportGenerationController
from .file_picker_controller import FilePickerController
from .configuration_controller import ConfigurationController
from ..utils.thread_pool import run_in_background

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI

logger = logging.getLogger(__name__)


class EventHandlers:
    """Handles all GUI events for the Motor Report App"""
    
    def __init__(self, gui: 'MotorReportAppGUI'):
        self.gui = gui
        # Thread lock for temp file operations
        self._temp_file_lock = threading.Lock()
        self._temp_files_created = set()  # Track temp files to prevent double deletion
        # Report save state
        self._temp_report_file = None
        self._report_filename = None

        self.search_controller = SearchController(
            gui,
            self.state,
            update_button_state=self._update_button_state,
            safe_status_update=self._safe_status_update,
            safe_results_update=self._safe_results_update,
        )
        self.selection_controller = SelectionController(gui, self.state)
        self.report_generation_controller = ReportGenerationController(
            gui,
            self.state,
            update_button_state=self._update_button_state,
        )
        self.file_picker_controller = FilePickerController(gui, self.state)
        self.configuration_controller = ConfigurationController(gui, self.state)
    
    @property
    def state(self):
        """Get state manager from GUI"""
        return self.gui.state_manager
    
    def _update_button_state(self, button_name, enabled, text=None, icon=None, bgcolor=None, color=None):
        """Helper method to update button state safely - centralized button management"""
        button = getattr(self.gui, button_name, None)
        if button:
            try:
                button.disabled = not enabled
                if text:
                    button.text = text
                if icon:
                    button.icon = icon
                if bgcolor:
                    button.bgcolor = bgcolor
                if color:
                    button.color = color
                logger.debug(f"Updated {button_name}: enabled={enabled}, text='{text}'")
                return True
            except Exception as e:
                logger.error(f"Error updating {button_name}: {e}")
                return False
        else:
            logger.debug(f"Button {button_name} not found")
            return False

    def _has_gui_component(self, component_name):
        """Helper to check if GUI component exists - replaces hasattr chains"""
        return hasattr(self.gui, component_name) and getattr(self.gui, component_name) is not None

    def _has_state_property(self, property_path):
        """Helper to safely check nested state properties"""
        try:
            if not self.state:
                return False
            parts = property_path.split('.')
            obj = self.state.state
            for part in parts:
                if not hasattr(obj, part):
                    return False
                obj = getattr(obj, part)
            return obj is not None
        except Exception:
            return False

    def _has_gui_property(self, property_path):
        """Helper to safely check nested GUI properties"""
        try:
            parts = property_path.split('.')
            obj = self.gui
            for part in parts:
                if not hasattr(obj, part) or getattr(obj, part) is None:
                    return False
                obj = getattr(obj, part)
            return True
        except Exception:
            return False
    
    def _safe_status_update(self, message: str, color: str = "black") -> bool:
        """Safely update status message with error handling"""
        try:
            if hasattr(self.gui, 'status_manager') and self.gui.status_manager:
                self.gui.status_manager.update_status(message, color)
                return True
            else:
                logger.warning(f"Status manager not available: {message}")
                return False
        except Exception as e:
            logger.error(f"Error updating status: {e}")
            return False
    
    def _safe_results_update(self) -> bool:
        """Safely update search results display with error handling"""
        try:
            display_method = None
            if hasattr(self.gui, '_enhanced_search_results_display'):
                display_method = self.gui._enhanced_search_results_display
            elif hasattr(self.gui, '_display_search_results'):
                display_method = self.gui._display_search_results

            if display_method is None:
                logger.warning("Search results display method not available")
                return False

            # Flet requires UI updates to occur on the main thread. When searches
            # run in a background thread we schedule the UI refresh using
            # page.invoke_later to avoid silent failures that left the results
            # area empty.
            invoke_later = getattr(getattr(self.gui, 'page', None), 'invoke_later', None)
            if callable(invoke_later):
                def _update_results():
                    try:
                        display_method()
                    finally:
                        self.gui._safe_page_update()

                invoke_later(_update_results)
                return True

            # Fallback for environments where invoke_later is not available.
            display_method()
            self.gui._safe_page_update()
            return True
        except Exception as e:
            logger.error(f"Error updating search results: {e}")
            return False
    
    def on_folder_picked(self, e: ft.FilePickerResultEvent):
        """Handle folder selection event - delegated to FilePickerController"""
        self.file_picker_controller.on_folder_picked(e)
    
    def on_registry_file_picked(self, e: ft.FilePickerResultEvent):
        """Handle registry file selection event - delegated to FilePickerController"""
        self.file_picker_controller.on_registry_file_picked(e)
    
    def on_tab_change(self, e):
        """Handle tab change events"""
        if hasattr(self.gui, 'workflow_manager'):
            self.gui.workflow_manager.handle_tab_change(e.control.selected_index)
    
    def on_go_to_configure_tab(self, e=None):
        """Handle Next Step button click to go to configure tab"""
        logger.info("🔗 Next Step button clicked - navigating to configure tab")
        workflow_manager = getattr(self.gui, 'workflow_manager', None)
        if workflow_manager and hasattr(workflow_manager, 'next_step'):
            # Use next_step() instead of go_to_step() for proper validation and visual feedback
            success = workflow_manager.next_step()
            if success:
                self.gui.status_manager.update_status(
                    "Moved to configuration step. Configure your report settings.", 
                    "blue"
                )
                logger.info("Successfully navigated to configure tab")
            else:
                self.gui.status_manager.update_status(
                    "Cannot proceed to configuration. Please complete the selection first.", 
                    "red"
                )
                logger.warning("Navigation to configure tab failed - selection not complete")
        else:
            logger.error("No workflow manager available for navigation")
    
    def on_search_clicked(self, e):
        """Delegate search workflow to the dedicated controller."""
        self.search_controller.on_search_clicked(e)
    
    def on_test_selected(self, e):
        """Handle test selection checkbox with automatic application - optimized for performance"""
        self.selection_controller.on_test_selected(e)
                
    def on_row_clicked(self, test):
        """Handle test row click with automatic selection application - optimized"""
        return self.selection_controller.row_click_handler(test)
    
    def on_generate_report_clicked(self, e):
        """Delegate report generation workflow to the dedicated controller."""
        self.report_generation_controller.on_generate_report_clicked(e)
    
    def _update_generate_button_state(self, generating: bool):
        """Update generate button state efficiently in a single operation"""
        try:
            if generating:
                self._update_button_state('generate_button', enabled=False, text="Generating...", icon=ft.Icons.HOURGLASS_EMPTY)
            else:
                self._update_button_state('generate_button', enabled=True, text="Generate Report", icon=ft.Icons.CREATE)
            
            # Single page update for button state
            self.gui._safe_page_update()
            
        except Exception as ui_err:
            logger.debug(f"Button state update failed: {ui_err}")
    
    def _generate_report(self):
        """Generate report in background thread - optimized for non-blocking operation"""
        try:
            self.gui.status_manager.show_progress("Preparing report generation...")
            
            # Get tests to process and SAP codes to validate
            tests_to_process = self.state.get_tests_to_process()
            
            if not tests_to_process:
                raise Exception("No tests selected for processing")
            
            # Get SAP selections from configuration step
            noise_saps = list(self.state.state.selected_noise_saps)
            comparison_saps = list(self.state.state.selected_comparison_saps)
            
            logger.info(f"Report generation data:")
            logger.info(f"  Performance tests: {len(tests_to_process)} tests")
            logger.info(f"  Tests: {[(t.test_lab_number, t.sap_code) for t in tests_to_process]}")
            logger.info(f"  Noise SAPs: {noise_saps}")
            logger.info(f"  Comparison SAPs: {comparison_saps}")
            
            # Log fine-grained comparison test lab selection
            if comparison_saps:
                for sap in comparison_saps:
                    selected_labs = self.state.state.selected_comparison_test_labs.get(sap, set())
                    if selected_labs:
                        logger.info(f"  Comparison SAP {sap}: Selected test labs {list(selected_labs)}")
                    else:
                        logger.info(f"  Comparison SAP {sap}: All tests (no specific selection)")
            
            # Log fine-grained noise test lab selection
            if noise_saps:
                for sap in noise_saps:
                    selected_noise_tests = self.state.state.selected_noise_test_labs.get(sap, set())
                    if selected_noise_tests:
                        logger.info(f"  Noise SAP {sap}: Selected tests {list(selected_noise_tests)}")
                    else:
                        logger.info(f"  Noise SAP {sap}: All tests (no specific selection)")
            
            # Generate filename with timestamp
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Motor_Performance_Report_{timestamp}.xlsx"
            
            # Update status without blocking UI
            self.gui.status_manager.update_status(f"🔄 Generating report: {filename}...", "blue")
            
            # Check if report manager exists
            if not hasattr(self.gui, 'report_manager') or not self.gui.report_manager:
                raise Exception("Report manager not available. Please restart the application.")
            
            # Create temp file safely (optimized)
            temp_path = self._create_temp_file_safely(filename)
            
            logger.info(f"Generating report to temporary location: {temp_path}")
            
            # Get multiple comparisons from state - Convert new comparison_groups format
            multiple_comparisons = []
            
            # Check if we have the new comparison_groups structure
            if hasattr(self.state.state, 'comparison_groups') and self.state.state.comparison_groups:
                logger.info("Converting new comparison_groups format to multiple_comparisons for report")
                
                for group_id, group_data in self.state.state.comparison_groups.items():
                    if isinstance(group_data, dict) and group_data:
                        # Extract test labs from all SAPs in this group
                        all_test_labs = []
                        for sap_code, test_labs in group_data.items():
                            if test_labs:
                                all_test_labs.extend(list(test_labs))
                        
                        if all_test_labs:
                            comparison_group = {
                                "id": group_id,
                                "name": group_id,  # Use group_id as name for now
                                "test_labs": all_test_labs,
                                "description": f"Comparison group with {len(group_data)} SAPs",
                                "sap_data": group_data  # Include the SAP-specific data
                            }
                            multiple_comparisons.append(comparison_group)
                            logger.info(f"  Converted group {group_id}: {len(all_test_labs)} test labs from {len(group_data)} SAPs")
            
            # Fallback to old multiple_comparisons if new format is not available
            if not multiple_comparisons and hasattr(self.state.state, 'multiple_comparisons'):
                multiple_comparisons = self.state.state.multiple_comparisons
                logger.info("Using existing multiple_comparisons format")
            
            logger.info(f"Final multiple_comparisons for report: {len(multiple_comparisons)} groups")
            
            # OPTIMIZED: Generate report with minimal status updates to avoid UI blocking
            try:
                self.gui.report_manager.generate_report_with_path(
                    tests_to_process=tests_to_process,
                    noise_saps=noise_saps,
                    comparison_saps=comparison_saps,
                    multiple_comparisons=multiple_comparisons,
                    output_path=temp_path
                )
                self.gui.status_manager.hide_progress()
            except Exception as report_err:
                # Handle specific report generation errors
                logger.error(f"Report generation process failed: {report_err}")
                raise Exception(f"Report generation failed: {str(report_err)}")
            
        except Exception as ex:
            logger.error(f"Report generation failed: {ex}")
            self.gui.status_manager.update_status(f"❌ Report generation failed: {str(ex)}", "red")
            self.gui.status_manager.hide_progress()
            
            # Show error dialog
            self._show_report_error_dialog(str(ex))
            
        finally:
            # End the report generation operation
            self.state.end_operation()
            
            # Re-enable generate report button efficiently
            self._update_generate_button_state(generating=False)
    
    def _show_download_dialog(self, filename: str, file_path: str):
        """Show simple save location picker - like folder/registry pickers"""
        
        # Store the source file for later copying
        self._temp_report_file = file_path
        self._report_filename = filename
        
        # Set picker context for save operation
        if self._has_state_property('picker_context'):
            self.gui.state_manager.state.picker_context = "save_report_via_file"
        
        # Use file picker to select any existing file in target directory
        self.gui.file_picker.pick_files(
            dialog_title="Select any file in the folder where you want to save the report",
            allow_multiple=False
        )
    
    def _show_download_success_dialog(self, filename: str, full_path: str):
        """Show success dialog after file has been downloaded/saved"""
        
        def on_open_folder(e):
            """Open the folder containing the report"""
            
            folder_path = os.path.dirname(full_path)
            try:
                if platform.system() == "Windows":
                    subprocess.run(f'explorer /select,"{full_path}"', shell=True)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", "-R", full_path])
                else:  # Linux
                    subprocess.run(["xdg-open", folder_path])
                success_dialog.open = False
                self.gui.page.update()
            except Exception as ex:
                logger.error(f"Error opening folder: {ex}")
        
        def on_close_success(e):
            success_dialog.open = False
            self.gui.page.update()
        
        success_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("✅ Download Complete!", weight=ft.FontWeight.BOLD, color="green"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Your report has been saved successfully!", size=14),
                    ft.Container(height=10),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("📄 Filename:", size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(filename, size=13, selectable=True),
                            ft.Container(height=8),
                            ft.Text("📁 Location:", size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(os.path.dirname(full_path), size=12, selectable=True, color="#666666"),
                        ], spacing=4),
                        padding=ft.padding.all(12),
                        bgcolor="#f0f8f0",
                        border_radius=8,
                        border=ft.border.all(1, "#4CAF50")
                    )
                ], spacing=10, tight=True),
                width=450,
                padding=ft.padding.all(20)
            ),
            actions=[
                ft.TextButton("Close", on_click=on_close_success),
                ft.ElevatedButton(
                    "📂 Open Folder", 
                    on_click=on_open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(bgcolor="#4CAF50", color="#FFFFFF")
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.gui.page.overlay.append(success_dialog)
        success_dialog.open = True
        self.gui.page.update()
    
    def _show_filename_input_dialog_with_file(self, default_filename: str, source_file_path: str):
        """Show save location dialog - uses file picker approach for browser compatibility"""
        
        def on_use_file_picker_for_save(e):
            """Use file picker to select save location (browser-compatible)"""
            # Store the source file for later copying
            self._temp_report_file = source_file_path
            self._report_filename = default_filename
            
            # Set picker context for save operation
            if self._has_state_property('picker_context'):
                self.gui.state_manager.state.picker_context = "save_report_via_file"
            
            # Use file picker to select any existing file in target directory
            self.gui.file_picker.pick_files(
                dialog_title="Select any file in the folder where you want to save the report",
                allow_multiple=False
            )
            
            # Close the current dialog and update once
            save_dialog.open = False
            self.gui.page.update()
        
        def on_manual_path_entry(e):
            """Show manual path entry dialog for advanced users"""
            self._show_manual_save_path_dialog(default_filename, source_file_path)
            # Batch dialog close and update
            save_dialog.open = False
            self.gui.page.update()
        
        def on_use_downloads(e):
            """Save to Downloads folder directly"""
            try:
                download_folder = os.path.expanduser("~/Downloads")
                final_path = os.path.join(download_folder, default_filename)
                
                # Copy file to Downloads folder
                shutil.copy2(source_file_path, final_path)
                
                # Close dialog and update UI
                save_dialog.open = False
                
                # Update status
                self.gui.status_manager.update_status(
                    f"✅ Report saved to Downloads: {default_filename}", 
                    "green"
                )
                self.gui.status_manager.hide_progress()
                
                # Single page update after all changes
                self.gui.page.update()
                
                # Show success dialog
                self._show_download_success_dialog(default_filename, final_path)
                
            except Exception as ex:
                logger.error(f"Error saving to Downloads: {ex}")
                self.gui.status_manager.update_status(f"❌ Save failed: {str(ex)}", "red")
        
        # Create save location dialog
        save_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Save Report"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text(f"Report generated: {default_filename}", size=14, weight=ft.FontWeight.W_500),
                    ft.Container(height=10),
                    ft.Text("Choose how to save the report:", size=13),
                    ft.Container(height=15),
                    
                    # Option 1: Save to Downloads folder (simplest)
                    ft.ElevatedButton(
                        "💾 Save to Downloads Folder",
                        icon=ft.Icons.DOWNLOAD,
                        on_click=on_use_downloads,
                        style=ft.ButtonStyle(bgcolor="#4CAF50", color="#FFFFFF"),
                        width=300
                    ),
                    
                    ft.Container(height=10),
                    
                    # Option 2: Use file picker (browser-compatible)
                    ft.ElevatedButton(
                        "📁 Choose Save Location",
                        icon=ft.Icons.FOLDER_OPEN,
                        on_click=on_use_file_picker_for_save,
                        width=300
                    ),
                    
                    ft.Container(height=10),
                    
                    # Option 3: Manual path entry
                    ft.ElevatedButton(
                        "✏️ Enter Path Manually",
                        icon=ft.Icons.EDIT,
                        on_click=on_manual_path_entry,
                        width=300
                    ),
                    
                ], spacing=8, tight=True, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                width=400,
                padding=ft.padding.all(20)
            ),
            actions=[
                ft.TextButton("Cancel", on_click=lambda e: self._close_dialog(save_dialog)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.gui.page.overlay.append(save_dialog)
        save_dialog.open = True
        self.gui.page.update()
    
    def _show_manual_save_path_dialog(self, default_filename: str, source_file_path: str):
        """Show manual path entry dialog as fallback"""
        def on_save_clicked(e):
            folder_path = path_input.value.strip() if path_input.value else ""
            filename = filename_input.value.strip() if filename_input.value else default_filename
            
            if not folder_path:
                error_text.value = "❌ Please enter a folder path"
                error_text.color = "red"
                self.gui.page.update()
                return
                
            if not filename.lower().endswith('.xlsx'):
                filename += '.xlsx'
                
            final_path = os.path.join(folder_path, filename)
            
            try:
                # Copy file to final location
                os.makedirs(folder_path, exist_ok=True)
                shutil.copy2(source_file_path, final_path)
                
                manual_dialog.open = False
                self.gui.page.update()
                
                self.gui.status_manager.update_status(f"✅ Report saved: {final_path}", "green")
                self._show_download_success_dialog(filename, final_path)
                
            except Exception as ex:
                error_text.value = f"❌ Save failed: {str(ex)}"
                error_text.color = "red"
                self.gui.page.update()
        
        path_input = ft.TextField(
            label="Folder Path", 
            hint_text="C:\\Users\\Username\\Documents",
            width=400,
            value=os.path.expanduser("~/Desktop")
        )
        filename_input = ft.TextField(
            label="Filename", 
            value=default_filename,
            width=400
        )
        error_text = ft.Text("", size=12)
        
        manual_dialog = ft.AlertDialog(
            title=ft.Text("Enter Save Location"),
            content=ft.Column([
                ft.Text("Enter the folder path and filename:"),
                path_input,
                filename_input,
                error_text
            ], tight=True),
            actions=[
                ft.TextButton("Cancel", on_click=lambda e: self._close_dialog(manual_dialog)),
                ft.ElevatedButton("Save", on_click=on_save_clicked)
            ]
        )
        
        self.gui.page.overlay.append(manual_dialog)
        manual_dialog.open = True
        self.gui.page.update()
    
    def _close_dialog(self, dialog):
        """Helper to close a dialog"""
        dialog.open = False
        self.gui.page.update()

    def on_sap_checked(self, sap_type: str):
        """Delegate SAP checkbox handler to the configuration controller."""
        return self.configuration_controller.on_sap_checked(sap_type)
    
    def on_confirm_test_selection(self, e):
        """Handle test selection confirmation"""
        if not self.state.state.selected_tests:
            self.gui.status_manager.update_status(
                "No tests selected. Please select at least one test.",
                "red"
            )
            return
        
        # Provide immediate visual feedback
        self.gui.status_manager.update_status("✅ Confirming test selection...", "blue")
        
        # Disable button temporarily to prevent double-clicks
        self._update_button_state('confirm_selection_button', enabled=False, text="Confirming...")
        
        self.gui._safe_page_update()
        
        self.state.apply_search_selection()
        
        # Show success message
        selected_count = len(self.state.state.selected_tests)
        self.gui.status_manager.update_status(
            f"✅ Selection confirmed: {selected_count} test(s) selected. Proceeding to configuration.",
            "green"
        )
        
        # Re-enable button and restore appearance
        self._update_button_state('confirm_selection_button', enabled=True, text="Confirm Selection")
        
        self.gui._safe_page_update()
        
        self.gui._go_to_configure_tab()
    
    def on_modify_test_selection(self, e):
        """Handle modify test selection"""
        # Provide immediate visual feedback
        self.gui.status_manager.update_status("🔄 Returning to search & select tab...", "blue")
        self.gui._safe_page_update()
        
        self.gui._go_to_search_select_tab()
    
    def on_start_new_search(self, e):
        """Handle start new search"""
        # Provide immediate visual feedback
        self.gui.status_manager.update_status("🔄 Starting new search...", "blue")
        
        # Disable button temporarily to prevent double-clicks
        self._update_button_state('new_search_button', enabled=False, text="Resetting...")
        
        self.gui._safe_page_update()
        
        self.state.reset_search()
        self.gui._display_search_results()
        self.gui.search_input_field.value = ""
        
        # Re-enable button and restore appearance
        self._update_button_state('new_search_button', enabled=True, text="New Search")
            # Note: This button doesn't have an initial icon
        
        self.gui.status_manager.update_status("✅ Ready for new search. Enter a SAP code or test number.", "green")
        
        self.gui._safe_page_update()
    
    def on_apply_search_selection(self, e=None):
        """Handle apply search selection"""
        # Provide immediate visual feedback
        self.gui.status_manager.update_status("⚙️ Applying selection...", "blue")
        
        # Disable button temporarily to prevent double-clicks
        self._update_button_state('apply_selection_button', enabled=False, text="Applying...", icon=ft.Icons.HOURGLASS_EMPTY)
        
        self.gui._safe_page_update()
        
        self.state.apply_search_selection()
        
        # Update UI
        selected_count = len(self.state.state.selected_tests)
        if selected_count > 0:
            self.gui.status_manager.update_status(
                f"✅ Selection applied: {selected_count} test(s) selected. "
                "You may now proceed to configuration.", 
                "green"
            )
        else:
            self.gui.status_manager.update_status(
                "❌ No tests selected. Please select at least one test to continue.", 
                "red"
            )
        
        # Re-enable button and restore appearance
        # Re-enable button and restore appearance
        self._update_button_state('apply_selection_button', enabled=True, text="Apply Selection", icon=ft.Icons.CHECK)
        
        # Refresh search select tab first
        self.gui.refresh_components(['search'])
        
        # Update workflow state after refresh
        if self._has_gui_component('workflow_manager'):
            self.gui.workflow_manager.update_workflow_state()
        
        # Full page update to ensure all controls are updated
        self.gui._safe_page_update()
    
    def on_clear_search_selection(self, e=None):
        """Handle clear search selection"""
        # Provide immediate visual feedback
        self.gui.status_manager.update_status("🧹 Clearing selection...", "blue")
        
        # Disable button temporarily to prevent double-clicks
        self._update_button_state('clear_selection_button', enabled=False, text="Clearing...", icon=ft.Icons.HOURGLASS_EMPTY)
        
        self.gui._safe_page_update()
        
        self.state.clear_search_selection()
        self.gui._display_search_results()
        
        # Re-enable button and restore appearance
        self._update_button_state('clear_selection_button', enabled=True, text="Clear Selection", icon=ft.Icons.CLEAR)
        
        self.gui.status_manager.update_status(
            "🧹 Selection cleared. Please select tests to continue.", 
            "orange"
        )
        
        # Refresh search select tab
        self.gui.refresh_components(['search'])
        
        # Update workflow state
        if self._has_gui_component('workflow_manager'):
            self.gui.workflow_manager.update_workflow_state()
        
        self.gui._safe_page_update()
    
    def on_apply_config_selection(self, e=None):
        """Handle apply configuration selection with enhanced visual feedback"""
        self._log_action("config_apply_started")
        
        # Provide immediate visual feedback and disable button
        self.gui.status_manager.update_status("⚙️ Applying configuration...", "blue")
        
        # Phase 1: Show "Applying..." state
        self._log_action("config_apply_button_updating", extra_info="Phase 1: Applying")
        self._update_button_state('config_apply_button', enabled=False, text="⏳ Applying", 
                                 icon=ft.Icons.HOURGLASS_EMPTY, bgcolor="#ff9800", color="white")
        
        # Force immediate page update to show button change
        self.gui._safe_page_update()
        
        # Mark configuration as applied
        self.state.state.config_selection_applied = True
        
        # Debug: Log what's in the state
        logger.info(f"🔍 Config Debug - selected_noise_saps: {self.state.state.selected_noise_saps}")
        logger.info(f"🔍 Config Debug - selected_comparison_saps: {self.state.state.selected_comparison_saps}")
        
        # Update UI
        selected_noise = len(self.state.state.selected_noise_saps)
        selected_comparison = len(self.state.state.selected_comparison_saps)
        
        # Show processing status for a moment
        self.gui.status_manager.update_status(
            f"⚙️ Processing configuration: Noise ({selected_noise} SAP codes), "
            f"Comparison ({selected_comparison} SAP codes)...", 
            "blue"
        )
        
        # Phase 2: Show "Success" state with persistence
        # Phase 2: Show success state 
        self._log_action("config_apply_button_updating", extra_info="Phase 2: Success")
        self._update_button_state('config_apply_button', enabled=True, text="✅ Configuration Applied!", 
                                 icon=ft.Icons.CHECK_CIRCLE, bgcolor="#4caf50", color="white")
        
        # Show success status
        self.gui.status_manager.update_status(
            f"✅ Configuration applied successfully: Noise ({selected_noise} SAP codes), "
            f"Comparison ({selected_comparison} SAP codes). Ready to generate report.", 
            "green"
        )
        
        # Update workflow state
        if hasattr(self.gui, 'workflow_manager'):
            self.gui.workflow_manager.update_workflow_state()
        
        # Refresh the Generate tab to show updated configuration
        if hasattr(self.gui, 'generate_tab'):
            self._log_action("generate_tab_refresh", extra_info="after config application")
            self.gui.generate_tab.refresh_content()
        
        # Final page update using proper Flet method
        if hasattr(self.gui, 'page'):
            self.gui.page.update()
        else:
            self.gui._safe_page_update()
        
        self._log_action("config_apply_completed")
        
        # Schedule a delayed update to keep the success state visible for a longer time
        def delayed_final_update():
            # Keep success state for user feedback but allow re-application
            if hasattr(self.gui, 'config_apply_button'):
                self.gui.config_apply_button.text = "✅ Applied - Click to Re-apply"
            self.gui._safe_page_update()
        
        # Use page's run_task if available for delayed execution, otherwise immediate
        if hasattr(self.gui.page, 'run_task'):
            import asyncio
            self.gui.page.run_task(self._delayed_update_after_config_apply)
        else:
            delayed_final_update()
    
    async def _delayed_update_after_config_apply(self):
        """Delayed update after config apply for persistent feedback"""
        import asyncio
        await asyncio.sleep(2)  # Keep success state visible for 2 seconds
        
        if hasattr(self.gui, 'config_apply_button'):
            self.gui.config_apply_button.text = "✅ Applied - Click to Re-apply"
            self.gui.config_apply_button.bgcolor = "#4caf50"  # Keep green
            self.gui.config_apply_button.color = "white"
            self.gui.config_apply_button.icon = ft.Icons.CHECK_CIRCLE
        
        self.gui._safe_page_update()
    
    def on_clear_config_selection(self, e=None):
        """Handle clear configuration selection"""
        # Clear all SAP selections
        self.state.state.selected_noise_saps.clear()
        self.state.state.selected_comparison_saps.clear()
        self.state.state.config_selection_applied = False
        
        self.gui.status_manager.update_status(
            "🧹 Configuration cleared. Please select SAP codes for features.", 
            "orange"
        )
        
        # Refresh the Configure tab to update checkboxes
        if hasattr(self.gui, 'workflow_manager'):
            self.gui.workflow_manager.refresh_tab('config')
            self.gui.workflow_manager.update_workflow_state()
        
        self.gui._safe_page_update()
    
    def on_go_to_search_select_tab(self, e=None):
        """Navigate back to Search & Select tab"""
        if hasattr(self.gui, 'workflow_manager'):
            self.gui.workflow_manager.go_to_step(1)  # Step 1 is Search & Select
    
    def on_go_to_generate_tab(self, e=None):
        """Navigate to Generate tab"""
        logger.info("🔗 Navigating to Generate tab...")
        
        # Navigate to the tab - workflow manager will handle the refresh
        if hasattr(self.gui, 'workflow_manager'):
            self.gui.workflow_manager.go_to_step(3)  # Step 3 is Generate
            
        self._log_action("generate_tab_navigation_success")
    
    def on_disconnect(self, e):
        """Handle page disconnect"""
        self.gui._cleanup_and_exit()
    
    def on_configuration_changed(self, config_key: str, value):
        """Handle configuration changes"""
        self.state.update_configuration(**{config_key: value})
        logger.debug(f"Configuration updated: {config_key} = {value}")
    
    def handle_post_build_setup(self):
        """Handle post-build setup tasks"""
        # Update initial paths from directory config
        from ...config.directory_config import PERFORMANCE_TEST_DIR, LAB_REGISTRY_FILE, TEST_LAB_CARICHI_DIR
        
        if PERFORMANCE_TEST_DIR:
            self.state.update_paths(tests_folder=str(PERFORMANCE_TEST_DIR))
            self.gui.tests_folder_path_text.value = f"Auto-detected: {PERFORMANCE_TEST_DIR}"
        
        if LAB_REGISTRY_FILE:
            self.state.update_paths(registry_file=str(LAB_REGISTRY_FILE))
            self.gui.registry_file_path_text.value = f"Auto-detected: {LAB_REGISTRY_FILE}"
        
        if TEST_LAB_CARICHI_DIR:
            self.state.update_paths(test_lab_dir=str(TEST_LAB_CARICHI_DIR))

        # Initialize backend in background using thread pool
        run_in_background(self._initialize_backend)
        
        self.gui._safe_page_update()
    
    def _initialize_backend(self):
        """Initialize the backend MotorReportApp in a background thread"""
        try:
            from ...core.motor_report_engine import MotorReportApp
            from ...config.app_config import AppConfig
            from ...config.directory_config import LOGO_PATH, NOISE_REGISTRY_FILE, NOISE_TEST_DIR
            
            self.gui.status_manager.status_text.value = "Initializing backend... Loading registry files..."
            self.gui.status_manager.progress_bar.visible = True
            self.gui._safe_page_update()
            
            config = AppConfig(
                tests_folder=self.state.state.selected_tests_folder,
                registry_path=self.state.state.selected_registry_file,
                output_path=".",
                logo_path=str(LOGO_PATH) if LOGO_PATH else None,
                noise_registry_path=str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else None,
                noise_dir=str(NOISE_TEST_DIR) if NOISE_TEST_DIR else None,
                test_lab_root=self.state.state.test_lab_directory or None,
                pressure_unit=self.state.state.pressure_unit,
                flow_unit=self.state.state.flow_unit,
                speed_unit=self.state.state.speed_unit,
                power_unit=self.state.state.power_unit,
                selected_lf_test_numbers=self.state.state.selected_lf_test_numbers,
            )
            
            self.gui.app = MotorReportApp(config)
            logger.info("Backend initialized successfully.")
            
            # Thread-safe UI update from background thread
            def update_ui_success():
                self.gui.status_manager.status_text.value = "Backend initialized successfully. Ready to search for tests!"
                self.gui.status_manager.progress_bar.visible = False
                self.gui._safe_page_update()
            
            if hasattr(self.gui.page, 'run_thread') and callable(self.gui.page.run_thread):
                self.gui.page.run_thread(update_ui_success)
            else:
                update_ui_success()
            
        except Exception as e:
            logger.error(f"Backend initialization failed: {e}")
            
            # Thread-safe UI update from background thread
            def update_ui_error():
                self.gui.status_manager.status_text.value = f"Backend initialization failed: {str(e)}"
                self.gui.status_manager.progress_bar.visible = False
                self.gui._safe_page_update()
            
            if hasattr(self.gui.page, 'run_thread') and callable(self.gui.page.run_thread):
                self.gui.page.run_thread(update_ui_error)
            else:
                update_ui_error()
    
    def on_save_file_picked(self, e):
        """Handle save file picker result"""
        try:
            if e.files and len(e.files) > 0:
                # Get the selected file path
                selected_path = e.files[0].path
                logger.info(f"User selected save location: {selected_path}")
                
                # Store the selected path and proceed with report generation
                self.selected_save_path = selected_path
                self._proceed_with_report_generation()
            else:
                logger.info("User cancelled save file selection")
                self.gui.status_manager.update_status("Report generation cancelled by user.", "orange")
                
                # Re-enable generate button
                self._update_button_state('generate_button', enabled=True, text="Generate Report", icon=ft.Icons.CREATE)
                self.gui._safe_page_update()
                    
        except Exception as ex:
            logger.error(f"Error handling save file picker result: {ex}")
            self.gui.status_manager.update_status(f"Error selecting save location: {str(ex)}", "red")
    
    def _proceed_with_report_generation(self):
        """Proceed with actual report generation after file location is selected"""
        import os
        
        try:
            if not hasattr(self, 'selected_save_path') or not self.selected_save_path:
                raise Exception("No save path selected")
                
            logger.info(f"Proceeding with report generation to: {self.selected_save_path}")
            
            # Get tests to process and SAP codes
            tests_to_process = self.state.get_tests_to_process()
            noise_saps = list(self.state.state.selected_noise_saps)
            comparison_saps = list(self.state.state.selected_comparison_saps)
            
            if not tests_to_process:
                self.gui.status_manager.update_status("❌ No tests selected for processing.", "red")
                self.gui.status_manager.hide_progress()
                return
            
            logger.info(f"Tests to process: {len(tests_to_process)}")
            logger.info(f"Noise SAPs: {noise_saps}")
            logger.info(f"Comparison SAPs: {comparison_saps}")
            
            # Update progress with detailed information
            filename = os.path.basename(self.selected_save_path)
            self.gui.status_manager.update_status(
                f"🔄 Generating report with {len(tests_to_process)} test(s)...", 
                "blue"
            )
            
            # Check if report manager exists
            if not hasattr(self.gui, 'report_manager') or not self.gui.report_manager:
                raise Exception("Report manager not available. Please restart the application.")
            
            # Use the report manager to generate the report with the selected path
            self.gui.report_manager.generate_report_with_path(
                tests_to_process=tests_to_process,
                noise_saps=noise_saps,
                comparison_saps=comparison_saps,
                output_path=self.selected_save_path
            )
            
            # Show success message
            self.gui.status_manager.update_status(
                f"✅ Report generated successfully: {filename}", 
                "green"
            )
            self.gui.status_manager.hide_progress()
            
            # Show success dialog
            self._show_report_success_dialog(filename, self.selected_save_path)
            
        except Exception as ex:
            logger.error(f"Error proceeding with report generation: {ex}")
            error_msg = f"❌ Report generation failed: {str(ex)}"
            self.gui.status_manager.update_status(error_msg, "red")
            self.gui.status_manager.hide_progress()
            
            # Show error dialog
            self._show_report_error_dialog(str(ex))
            
            # Re-enable generate button
            self._update_button_state('generate_button', enabled=True, text="Generate Report", icon=ft.Icons.CREATE)
            self.gui._safe_page_update()
    
    def _show_report_success_dialog(self, filename: str, full_path: str):
        """Show success dialog after report generation"""
        import os
        
        def on_open_folder(e):
            """Open the folder containing the report"""
            import subprocess
            import platform
            
            folder_path = os.path.dirname(full_path)
            try:
                if platform.system() == "Windows":
                    subprocess.run(f'explorer /select,"{full_path}"', shell=True)
                elif platform.system() == "Darwin":  # macOS
                    subprocess.run(["open", "-R", full_path])
                else:  # Linux
                    subprocess.run(["xdg-open", folder_path])
                success_dialog.open = False
                self.gui.page.update()
            except Exception as ex:
                logger.error(f"Error opening folder: {ex}")
        
        def on_close_success(e):
            success_dialog.open = False
            self.gui.page.update()
        
        success_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("✅ Report Generated Successfully!", weight=ft.FontWeight.BOLD, color="green"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("Your motor performance report has been created successfully!", size=14),
                    ft.Container(height=10),
                    ft.Container(
                        content=ft.Column([
                            ft.Text("📄 Filename:", size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(filename, size=13, selectable=True),
                            ft.Container(height=8),
                            ft.Text("📁 Location:", size=12, weight=ft.FontWeight.BOLD),
                            ft.Text(os.path.dirname(full_path), size=12, selectable=True, color="#666666"),
                        ], spacing=4),
                        padding=ft.padding.all(12),
                        bgcolor="#f0f8f0",
                        border_radius=8,
                        border=ft.border.all(1, "#4CAF50")
                    )
                ], spacing=10, tight=True),
                width=450,
                padding=ft.padding.all(20)
            ),
            actions=[
                ft.TextButton("Close", on_click=on_close_success),
                ft.ElevatedButton(
                    "📂 Open Folder", 
                    on_click=on_open_folder,
                    icon=ft.Icons.FOLDER_OPEN,
                    style=ft.ButtonStyle(bgcolor="#4CAF50", color="#FFFFFF")
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.gui.page.overlay.append(success_dialog)
        success_dialog.open = True
        self.gui.page.update()
    
    def _show_report_error_dialog(self, error_message: str):
        """Show error dialog if report generation fails"""
        def on_close_error(e):
            error_dialog.open = False
            self.gui.page.update()
        
        def on_retry(e):
            error_dialog.open = False
            self.gui.page.update()
            # Trigger the generate report process again
            self.on_generate_report_clicked(None)
        
        error_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("❌ Report Generation Failed", weight=ft.FontWeight.BOLD, color="red"),
            content=ft.Container(
                content=ft.Column([
                    ft.Text("An error occurred while generating your report:", size=14),
                    ft.Container(height=10),
                    ft.Container(
                        content=ft.Text(error_message, size=12, selectable=True),
                        padding=ft.padding.all(12),
                        bgcolor="#fff5f5",
                        border_radius=8,
                        border=ft.border.all(1, "#f44336")
                    ),
                    ft.Container(height=10),
                    ft.Text("Please check your settings and try again.", size=12, color="#666666")
                ], spacing=8, tight=True),
                width=450,
                padding=ft.padding.all(20)
            ),
            actions=[
                ft.TextButton("Close", on_click=on_close_error),
                ft.ElevatedButton(
                    "🔄 Try Again", 
                    on_click=on_retry,
                    icon=ft.Icons.REFRESH,
                    style=ft.ButtonStyle(bgcolor="#2196F3", color="#FFFFFF")
                ),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        
        self.gui.page.overlay.append(error_dialog)
        error_dialog.open = True
        self.gui.page.update()
    
    def _create_temp_file_safely(self, filename: str) -> str:
        """
        Create a temporary file safely with race condition protection.
        
        Args:
            filename: Desired filename for the temp file
            
        Returns:
            Path to the created temporary file
        """
        from ...utils.common import sanitize_filename, validate_directory_path
        
        with self._temp_file_lock:
            # Sanitize filename
            safe_filename = sanitize_filename(filename)
            
            # Create temp directory if needed
            temp_dir = validate_directory_path(tempfile.gettempdir())
            if not temp_dir:
                raise OSError("Cannot access temporary directory")
            
            # Generate unique temp file path
            counter = 0
            while True:
                if counter == 0:
                    temp_path = temp_dir / safe_filename
                else:
                    name, ext = os.path.splitext(safe_filename)
                    temp_path = temp_dir / f"{name}_{counter}{ext}"
                
                if not temp_path.exists():
                    break
                counter += 1
                if counter > 1000:  # Prevent infinite loop
                    raise OSError("Cannot create unique temporary file")
            
            # Create the file and track it
            temp_path.touch()
            self._temp_files_created.add(str(temp_path))
            logger.info(f"Created temp file: {temp_path}")
            return str(temp_path)

    def _cleanup_temp_file_safely(self, file_path: str):
        """
        Clean up temporary file safely with race condition protection.
        
        Args:
            file_path: Path to the temporary file to clean up
        """
        with self._temp_file_lock:
            if file_path not in self._temp_files_created:
                logger.debug(f"Temp file {file_path} not in our tracking set, skipping cleanup")
                return
                
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.info(f"Cleaned up temp file: {file_path}")
                self._temp_files_created.discard(file_path)
            except Exception as ex:
                logger.warning(f"Could not clean up temp file {file_path}: {ex}")
    
    def _save_report_to_folder(self, folder_path: str):
        """Save the generated report to the specified folder"""
        try:
            if not hasattr(self, '_temp_report_file') or not self._temp_report_file:
                raise Exception("No report file to save")
            
            if not hasattr(self, '_report_filename') or not self._report_filename:
                raise Exception("No report filename specified")
                
            final_path = os.path.join(folder_path, self._report_filename)
            
            # Copy file to final location
            shutil.copy2(self._temp_report_file, final_path)
            
            # Show success message
            self.gui.status_manager.update_status(
                f"✅ Report saved: {final_path}", 
                "green"
            )
            self.gui.status_manager.hide_progress()
            
            # Show success dialog
            self._show_download_success_dialog(self._report_filename, final_path)
            
            # Clean up temp file reference
            self._temp_report_file = None
            self._report_filename = None
            
        except Exception as ex:
            logger.error(f"Error saving report: {ex}")
            self.gui.status_manager.update_status(f"❌ Save failed: {str(ex)}", "red")
    
    def _log_action(self, action, level="info", extra_info=None):
        """
        Standardized logging method to replace excessive emoji logging.
        
        Args:
            action: The action being performed (e.g., "config_apply", "search_start")
            level: Log level ("info", "debug", "warning", "error")
            extra_info: Optional additional information
        """
        message = f"Action: {action}"
        if extra_info:
            message += f" - {extra_info}"
        
        log_method = getattr(logger, level, logger.info)
        log_method(message)

