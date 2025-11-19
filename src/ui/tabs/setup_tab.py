"""
Setup Tab - First tab for verifying input paths with manual entry support
"""
import flet as ft
import os
import logging
from pathlib import Path
from ..components.base import BaseTab
from ...config.directory_config import (
    PROJECT_ROOT, invalidate_directory_cache, 
    update_manual_paths, refresh_directory_cache, get_current_paths
)
from ...config.directory_cache import get_directory_cache

logger = logging.getLogger(__name__)


class SetupTab(BaseTab):
    """Tab for setting up and verifying input paths with manual editing"""
    
    def __init__(self, parent_gui=None):
        super().__init__(parent_gui)
        self.tab_name = "1. Setup"
        self.tab_icon = ft.Icons.SETTINGS
        
        # Create text field controls for manual path entry
        self.performance_path_field = ft.TextField(
            label="Performance Test Directory Path",
            hint_text="C:\\path\\to\\ProveEffettuate",
            expand=True,
            multiline=False
        )
        
        self.noise_path_field = ft.TextField(
            label="Noise Test Directory Path",
            hint_text="C:\\path\\to\\Tests Rumore",
            expand=True,
            multiline=False
        )
        
        self.lab_registry_field = ft.TextField(
            label="Lab Registry File Path",
            hint_text="C:\\path\\to\\Registro LAB.xlsx",
            expand=True,
            multiline=False
        )
        
        self.noise_registry_field = ft.TextField(
            label="Noise Registry File Path",
            hint_text="C:\\path\\to\\REGISTRO RUMORE.xlsx",
            expand=True,
            multiline=False
        )
        
        # Test Lab (CARICHI NOMINALI) directory
        self.test_lab_dir_field = ft.TextField(
            label="Test Lab Directory Path (CARICHI NOMINALI)",
            hint_text="C:\\path\\to\\CARICHI NOMINALI",
            expand=True,
            multiline=False
        )
        
        # LF (Life Test) fields
        self.lf_registry_field = ft.TextField(
            label="LF Registry File Path",
            hint_text="C:\\path\\to\\REGISTRO LF .xlsx",
            expand=True,
            multiline=False
        )
        
        self.lf_base_dir_field = ft.TextField(
            label="LF Base Directory Path (contains year folders)",
            hint_text="C:\\path\\to\\LF\\RELIABIL",
            expand=True,
            multiline=False
        )
        
        # Output directory field
        self.output_path_field = ft.TextField(
            label="Output Directory Path",
            hint_text="C:\\path\\to\\output",
            expand=True,
            multiline=False
        )
        
        # Status text controls
        self.status_text = ft.Text(
            "Ready to configure paths",
            color=self.theme_color('on_surface', '#fefefe')
        )
        self.cache_status_text = ft.Text("", color=self.theme_color('on_surface', '#fefefe'))
        
        # Create aliases for the fields to match what the code expects
        self.performance_field = self.performance_path_field
        self.noise_field = self.noise_path_field
        self.lab_registry_field = self.lab_registry_field
        self.noise_registry_field = self.noise_registry_field
        self.output_field = self.output_path_field
        
        self._path_fields = [
            self.performance_path_field,
            self.noise_path_field,
            self.lab_registry_field,
            self.noise_registry_field,
            self.test_lab_dir_field,
            self.lf_registry_field,
            self.lf_base_dir_field,
            self.output_path_field,
        ]

        self.apply_textfield_theme()

        # Load current paths
        self._load_current_paths()
    
    def get_tab_content(self) -> ft.Control:
        """Build the setup tab content with manual path entry"""
        color = self.theme_color
        onsurface = color('on_surface', '#fefefe')
        muted = color('text_muted', '#cfd8e3')
        return ft.Container(
            content=ft.Column([
                ft.Text("Path Configuration", size=20, weight=ft.FontWeight.BOLD, color=onsurface),
                ft.Text(
                    "Configure data paths manually or use auto-detection. Changes will be cached for future runs.", 
                    color=muted,
                    size=14
                ),
                ft.Divider(),
                
                # Manual Path Entry Section
                ft.Text("Manual Path Entry", size=16, weight=ft.FontWeight.W_500, color=onsurface),
                
                # Performance Test Directory
                ft.Text("Performance Test Directory:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.performance_path_field,
                
                # Noise Test Directory  
                ft.Text("Noise Test Directory:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.noise_path_field,
                
                # Lab Registry File
                ft.Text("Lab Registry File:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.lab_registry_field,
                
                # Noise Registry File
                ft.Text("Noise Registry File:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.noise_registry_field,
                
                # Test Lab Directory (CARICHI NOMINALI)
                ft.Text("Test Lab Directory (CARICHI NOMINALI):", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.test_lab_dir_field,
                
                # LF Registry File
                ft.Text("🔬 Life Test (LF) Registry File:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.lf_registry_field,
                
                # LF Base Directory
                ft.Text("🔬 Life Test (LF) Base Directory:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.lf_base_dir_field,
                
                # Output Directory
                ft.Text("Output Directory:", weight=ft.FontWeight.W_400, size=14, color=onsurface),
                self.output_path_field,
                
                ft.Divider(),
                
                # Action Buttons
                ft.Row([
                    ft.ElevatedButton(
                        "Validate Paths",
                        icon=ft.Icons.CHECK_CIRCLE,
                        on_click=self._on_validate_paths,
                        bgcolor=color('info', '#1976d2'),
                        color=color('on_info', 'white')
                    ),
                    ft.ElevatedButton(
                        "Save Paths",
                        icon=ft.Icons.SAVE,
                        on_click=self._on_save_paths,
                        bgcolor=color('success', 'green'),
                        color=color('on_success', 'white')
                    ),
                    ft.ElevatedButton(
                        "Refresh Cache",
                        icon=ft.Icons.REFRESH,
                        on_click=self._on_refresh_cache,
                        bgcolor=color('warning', 'orange'),
                        color=color('on_warning', 'black')
                    ),
                ], spacing=10, wrap=True),
                
                ft.Divider(),
                
                # Status Section
                ft.Text("Status", size=16, weight=ft.FontWeight.W_500, color=onsurface),
                self.status_text,
                self.cache_status_text,
                
                ft.Container(
                    content=ft.Row([
                        ft.Icon(ft.Icons.INFO, color=color('info', 'blue')),
                        ft.Text(
                            "Enter paths manually above and click 'Save Paths' to remember them for future use.", 
                            color=color('info', 'blue'),
                            size=12
                        )
                    ], spacing=10),
                    padding=ft.padding.all(15),
                    bgcolor=color('info_container', '#e3f2fd'),
                    border_radius=5
                )
            ], spacing=15, scroll=ft.ScrollMode.AUTO),
            padding=ft.padding.all(20),
            expand=True
        )

    def apply_textfield_theme(self):
        if not hasattr(self, "parent_gui") or not self.parent_gui:
            return
        style_helper = getattr(self.parent_gui, "_style_text_field", None)
        if not callable(style_helper):
            return
        for field in getattr(self, "_path_fields", []):
            style_helper(field)
    
    # File picker methods removed - manual entry only






    
    def _load_current_paths(self):
        """Load current paths from directory_config into the text fields"""
        try:
            from ...config.directory_config import get_current_paths
            
            paths = get_current_paths()
            
            # Update text fields with current paths
            if paths.get('performance_dir'):
                self.performance_field.value = paths['performance_dir']
            if paths.get('noise_dir'):
                self.noise_field.value = paths['noise_dir']
            if paths.get('lab_registry'):
                self.lab_registry_field.value = paths['lab_registry']
            if paths.get('noise_registry'):
                self.noise_registry_field.value = paths['noise_registry']
            if paths.get('test_lab_dir'):
                self.test_lab_dir_field.value = paths['test_lab_dir']
            if paths.get('lf_registry'):
                self.lf_registry_field.value = paths['lf_registry']
            if paths.get('lf_base_dir'):
                self.lf_base_dir_field.value = paths['lf_base_dir']
            if paths.get('output_dir'):
                self.output_field.value = paths['output_dir']
            
            # Update cache status
            self._update_cache_status()
            
        except Exception as e:
            self.status_text.value = f"Error loading paths: {str(e)}"
            self.status_text.color = self.theme_color('error', 'red')
    
    def _update_cache_status(self):
        """Update cache status display"""
        try:
            from ...config.directory_config import get_cache_status
            
            cache_info = get_cache_status()
            if cache_info.get('is_valid'):
                self.cache_status_text.value = f"Cache: Valid ({cache_info.get('registry_directories', 0)} registry, {cache_info.get('inf_directories', 0)} inf dirs)"
                self.cache_status_text.color = self.theme_color('success', 'green')
            else:
                self.cache_status_text.value = "Cache: Invalid or empty"
                self.cache_status_text.color = self.theme_color('warning', 'orange')
                
        except Exception as e:
            self.cache_status_text.value = f"Cache error: {str(e)}"
            self.cache_status_text.color = self.theme_color('error', 'red')
    
    def _validate_path(self, path_str: str) -> bool:
        """Validate if a path exists and is accessible"""
        if not path_str or not path_str.strip():
            return False
        
        try:
            from pathlib import Path
            path = Path(path_str.strip())
            return path.exists()
        except Exception:
            return False
    
    def _on_validate_paths(self, e):
        """Validate all entered paths"""
        try:
            # Validate each path
            validations = [
                ("Performance Dir", self.performance_field.value),
                ("Noise Dir", self.noise_field.value),
                ("Lab Registry", self.lab_registry_field.value),
                ("Noise Registry", self.noise_registry_field.value),
                ("Output Dir", self.output_field.value)
            ]
            
            all_valid = True
            messages = []
            
            for name, path in validations:
                if path and path.strip():
                    if self._validate_path(path):
                        messages.append(f"✅ {name}: Valid")
                    else:
                        messages.append(f"❌ {name}: Not found or inaccessible")
                        all_valid = False
                else:
                    messages.append(f"⚠️ {name}: Empty")
            
            # Update status
            if all_valid:
                self.status_text.value = "All paths validated successfully!"
                self.status_text.color = self.theme_color('success', 'green')
            else:
                self.status_text.value = "Some paths are invalid or missing"
                self.status_text.color = self.theme_color('warning', 'orange')
            
            # Show detailed validation in a dialog
            validation_text = "\n".join(messages)
            
            def close_dialog(e):
                dialog.open = False
                self._safe_page_update()
            
            dialog = ft.AlertDialog(
                title=ft.Text("Path Validation Results"),
                content=ft.Text(validation_text, selectable=True),
                actions=[ft.TextButton("OK", on_click=close_dialog)],
                actions_alignment=ft.MainAxisAlignment.END,
            )
            
            if self.parent_gui and hasattr(self.parent_gui, 'page'):
                self.parent_gui.page.dialog = dialog
                dialog.open = True
                self._safe_page_update()
            
        except Exception as ex:
            self.status_text.value = f"Validation error: {str(ex)}"
            self.status_text.color = self.theme_color('error', 'red')
            self._safe_page_update()
    
    def _on_save_paths(self, e):
        """Save the manually entered paths to cache"""
        try:
            from ...config.directory_config import update_cached_paths
            from pathlib import Path
            
            # Collect paths from form
            paths_to_save = {}
            
            if self.performance_field.value and self.performance_field.value.strip():
                paths_to_save['performance_dir'] = self.performance_field.value.strip()
            
            if self.noise_field.value and self.noise_field.value.strip():
                paths_to_save['noise_dir'] = self.noise_field.value.strip()
            
            if self.lab_registry_field.value and self.lab_registry_field.value.strip():
                paths_to_save['lab_registry'] = self.lab_registry_field.value.strip()
            
            if self.noise_registry_field.value and self.noise_registry_field.value.strip():
                paths_to_save['noise_registry'] = self.noise_registry_field.value.strip()
            
            if self.test_lab_dir_field.value and self.test_lab_dir_field.value.strip():
                paths_to_save['test_lab_dir'] = self.test_lab_dir_field.value.strip()
            
            if self.lf_registry_field.value and self.lf_registry_field.value.strip():
                paths_to_save['lf_registry'] = self.lf_registry_field.value.strip()
            
            if self.lf_base_dir_field.value and self.lf_base_dir_field.value.strip():
                paths_to_save['lf_base_dir'] = self.lf_base_dir_field.value.strip()
            
            if self.output_field.value and self.output_field.value.strip():
                paths_to_save['output_dir'] = self.output_field.value.strip()
            
            # Validate paths before saving
            invalid_paths = []
            for name, path in paths_to_save.items():
                if not self._validate_path(path):
                    invalid_paths.append(f"{name}: {path}")
            
            if invalid_paths:
                error_msg = "Cannot save invalid paths:\n" + "\n".join(invalid_paths)
                self.status_text.value = error_msg
                self.status_text.color = self.theme_color('error', 'red')
                self._safe_page_update()
                return
            
            # Save paths
            result = update_cached_paths(paths_to_save)
            
            if result.get('status') == 'success':
                self.status_text.value = f"Paths saved successfully! Updated {result.get('updated_count', 0)} paths."
                self.status_text.color = self.theme_color('success', 'green')
                self._update_cache_status()
                
                # Update state manager with new paths and load noise registry if available
                if self.parent_gui and hasattr(self.parent_gui, 'state_manager'):
                    try:
                        state_manager = self.parent_gui.state_manager
                        state_manager.update_paths(
                            tests_folder=paths_to_save.get('performance_dir'),
                            registry_file=paths_to_save.get('lab_registry'),
                            noise_folder=paths_to_save.get('noise_dir'),
                            noise_registry=paths_to_save.get('noise_registry'),
                            test_lab_dir=paths_to_save.get('test_lab_dir')
                        )
                        
                        # Try to load noise registry data if noise registry path was provided
                        if paths_to_save.get('noise_registry'):
                            if state_manager.load_noise_registry_data():
                                self.status_text.value += " Noise registry loaded successfully."
                            else:
                                self.status_text.value += " Warning: Failed to load noise registry."
                                
                    except Exception as e:
                        logger.warning(f"Failed to update state manager: {e}")
                
            else:
                self.status_text.value = f"Save failed: {result.get('message', 'Unknown error')}"
                self.status_text.color = self.theme_color('error', 'red')
            
        except Exception as ex:
            self.status_text.value = f"Save error: {str(ex)}"
            self.status_text.color = self.theme_color('error', 'red')
        
        self._safe_page_update()
    
    def _on_refresh_cache(self, e):
        """Refresh the directory cache by re-scanning"""
        try:
            from ...config.directory_config import refresh_directory_cache
            
            self.status_text.value = "🔄 Refreshing directory cache..."
            self.status_text.color = self.theme_color('info', 'blue')
            self._safe_page_update()
            
            # Refresh cache
            result = refresh_directory_cache()
            
            if result.get('status') == 'success':
                self.status_text.value = f"Cache refreshed! Found {result.get('success_count', 0)}/4 targets."
                self.status_text.color = self.theme_color('success', 'green')
                
                # Reload paths into form
                self._load_current_paths()
            else:
                self.status_text.value = f"Cache refresh failed: {result.get('message', 'Unknown error')}"
                self.status_text.color = self.theme_color('error', 'red')
            
            self._update_cache_status()
            
        except Exception as ex:
            self.status_text.value = f"Refresh error: {str(ex)}"
            self.status_text.color = self.theme_color('error', 'red')
        
        self._safe_page_update()
    
    def _safe_page_update(self):
        """Safely update the page if possible"""
        try:
            if self.parent_gui and hasattr(self.parent_gui, 'page') and self.parent_gui.page:
                self.parent_gui.page.update()
        except Exception as e:
            # Silently ignore update errors to prevent crashes
            pass

