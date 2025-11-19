"""
Report generation manager for the Motor Report GUI
Handles all report generation operations and backend integration.
"""
import logging
import threading
from typing import List, Optional, TYPE_CHECKING
from pathlib import Path
from ...data.models import Test
from ...config.app_config import AppConfig
from ...config.directory_config import LOGO_PATH

if TYPE_CHECKING:
    from ..main_gui import MotorReportAppGUI

logger = logging.getLogger(__name__)


class ReportManager:
    """Manages report generation and backend integration"""
    
    def __init__(self, gui: 'MotorReportAppGUI'):
        self.gui = gui
        self.selected_noise_tests = []  # Store selected noise tests
    
    @property
    def state(self):
        """Get state manager from GUI"""
        return self.gui.state_manager
    
    def initialize_backend(self):
        """Initialize the backend MotorReportApp"""
        try:
            # Update status to show initialization is starting
            self.gui.status_manager.update_status(
                "Initializing backend... Loading registry files...", 
                "blue"
            )
            
            # Load noise registry data if not already loaded
            if not self.state.is_noise_registry_loaded():
                if self.state.state.selected_noise_registry:
                    self.gui.status_manager.update_status(
                        "Loading noise registry data...", 
                        "blue"
                    )
                    if not self.state.load_noise_registry_data():
                        logger.warning("Failed to load noise registry data")
            
            # DEBUG: Log LF selections before creating config
            logger.info("=" * 60)
            logger.info("🔬 FRONTEND: LF DATA BEFORE CONFIG CREATION")
            logger.info("=" * 60)
            logger.info(f"📊 selected_lf_test_numbers: {self.state.state.selected_lf_test_numbers}")
            logger.info(f"📊 selected_lf_saps: {self.state.state.selected_lf_saps}")
            logger.info("=" * 60)
            
            logo_path = None
            if LOGO_PATH:
                logo_candidate = Path(LOGO_PATH)
                if logo_candidate.exists():
                    logo_path = str(logo_candidate)

            config = AppConfig(
                tests_folder=self.state.state.selected_tests_folder,
                registry_path=self.state.state.selected_registry_file,
                output_path=".",
                logo_path=logo_path,
                noise_registry_path=self.state.state.selected_noise_registry,
                noise_dir=self.state.state.selected_noise_folder,
                test_lab_root=self.state.state.test_lab_directory or None,
                pressure_unit=self.state.state.pressure_unit,
                flow_unit=self.state.state.flow_unit,
                speed_unit=self.state.state.speed_unit,
                power_unit=self.state.state.power_unit,
                selected_lf_test_numbers=self.state.state.selected_lf_test_numbers,  # ✅ FIX: Add LF selections!
            )
            
            # Create simplified noise handler with pre-loaded registry data
            noise_handler = None
            if self.state.is_noise_registry_loaded():
                # Simplified noise handler removed in refactor; fallback to default handler passed later
                noise_handler = None
                logger.info("Simplified noise handler module removed; using default backend handler")

            from ...core.motor_report_engine import MotorReportApp
            self.gui.app = MotorReportApp(config, noise_handler=noise_handler)
            logger.info("Backend initialized successfully.")
            
            # Update status to show success
            self.gui.status_manager.update_status(
                "Backend initialized successfully. Ready to search for tests!", 
                "green"
            )
            
        except Exception as e:
            logger.error(f"Backend initialization failed: {e}")
            self.gui.status_manager.update_status(
                f"Backend initialization failed: {str(e)}", 
                "red"
            )
    
    def update_backend_config(self):
        """Update backend configuration with current paths"""
        if not hasattr(self.gui, 'app') or not self.gui.app:
            return
        
        try:
            # Load noise registry data if not already loaded
            if not self.state.is_noise_registry_loaded():
                if self.state.state.selected_noise_registry:
                    if not self.state.load_noise_registry_data():
                        logger.warning("Failed to load noise registry data")
            
            # DEBUG: Log LF selections before creating config
            logger.info("=" * 60)
            logger.info("🔬 DEBUG: LF DATA BEFORE CONFIG CREATION")
            logger.info("=" * 60)
            logger.info(f"📊 selected_lf_test_numbers type: {type(self.state.state.selected_lf_test_numbers)}")
            logger.info(f"📊 selected_lf_test_numbers value: {self.state.state.selected_lf_test_numbers}")
            logger.info(f"📊 Is empty: {not self.state.state.selected_lf_test_numbers}")
            if self.state.state.selected_lf_test_numbers:
                for sap, tests in self.state.state.selected_lf_test_numbers.items():
                    logger.info(f"   📌 SAP {sap}: {len(tests)} test(s) - {tests}")
            logger.info("=" * 60)
            
            logo_path = None
            if LOGO_PATH:
                logo_candidate = Path(LOGO_PATH)
                if logo_candidate.exists():
                    logo_path = str(logo_candidate)

            new_config = AppConfig(
                tests_folder=self.state.state.selected_tests_folder,
                registry_path=self.state.state.selected_registry_file,
                output_path=".",
                logo_path=logo_path,
                noise_registry_path=self.state.state.selected_noise_registry,
                noise_dir=self.state.state.selected_noise_folder,
                test_lab_root=self.state.state.test_lab_directory or None,
                include_noise=self.state.state.include_noise,
                include_comparison=self.state.state.include_comparison,
                registry_sheet_name=self.state.state.registry_sheet_name,
                noise_registry_sheet_name=self.state.state.noise_registry_sheet_name,
                pressure_unit=self.state.state.pressure_unit,
                flow_unit=self.state.state.flow_unit,
                speed_unit=self.state.state.speed_unit,
                power_unit=self.state.state.power_unit,
                selected_lf_test_numbers=self.state.state.selected_lf_test_numbers,
            )
            
            # Create simplified noise handler with pre-loaded registry data
            noise_handler = None
            if self.state.is_noise_registry_loaded():
                # Simplified noise handler removed; skip custom initialization
                noise_handler = None
            
            # Reinitialize backend with new config and simplified noise handler
            from ...core.motor_report_engine import MotorReportApp
            self.gui.app = MotorReportApp(new_config, noise_handler=noise_handler)
            logger.info("Backend configuration updated successfully.")
            
        except Exception as e:
            logger.error(f"Failed to update backend configuration: {e}")
    
    def generate_report_with_path(self, tests_to_process: List[Test], noise_saps: List[str], comparison_saps: List[str], output_path: str, multiple_comparisons: Optional[List] = None):
        """Generate the motor report with specified file path and noise test pre-check"""
        
        def continue_report_generation(selected_noise_tests):
            """Continue report generation after noise test validation."""
            try:
                self._do_generate_report(tests_to_process, noise_saps, comparison_saps, output_path, selected_noise_tests, multiple_comparisons)
            except Exception as e:
                logger.error(f"Report generation failed: {e}")
                self.gui.status_manager.update_status(
                    f"❌ Report generation failed: {str(e)}", 
                    "red"
                )
                
        # Start with noise test validation
        self.validate_noise_tests_if_needed(continue_report_generation)
    
    def _do_generate_report(self, tests_to_process: List[Test], noise_saps: List[str], comparison_saps: List[str], output_path: str, selected_noise_tests: List, multiple_comparisons: Optional[List] = None):
        """Perform the actual report generation after validation."""
        try:
            if not hasattr(self.gui, 'app') or not self.gui.app:
                raise Exception("Backend not initialized")

            # Validate data flow first
            if not self.validate_data_flow():
                raise Exception("Data flow validation failed")

            logger.info(f"Starting report generation with separated data:")
            logger.info(f"  Output path: {output_path}")
            logger.info(f"  Performance tests: {len(tests_to_process)} tests")
            logger.info(f"  Tests: {[(t.test_lab_number, t.sap_code) for t in tests_to_process]}")
            logger.info(f"  Noise SAPs: {noise_saps}")
            logger.info(f"  Comparison SAPs: {comparison_saps}")
            logger.info(f"  Selected noise tests: {len(selected_noise_tests)}")

            # Validate data consistency
            all_test_saps = set(t.sap_code for t in tests_to_process if t.sap_code)
            logger.info(f"  All SAP codes in selected tests: {sorted(all_test_saps)}")

            # Validate noise SAPs are subset of test SAPs
            invalid_noise_saps = set(noise_saps) - all_test_saps if noise_saps else set()
            if invalid_noise_saps:
                logger.warning(f"  Warning: Noise SAPs not in test selection: {invalid_noise_saps}")

            # Validate comparison SAPs are subset of test SAPs  
            invalid_comparison_saps = set(comparison_saps) - all_test_saps if comparison_saps else set()
            if invalid_comparison_saps:
                logger.warning(f"  Warning: Comparison SAPs not in test selection: {invalid_comparison_saps}")
            
            # Set the output path in the backend config
            self.gui.app.config.output_path = output_path
            
            # Update backend configuration with current settings
            self.gui.app.config.include_noise = self.state.state.include_noise
            self.gui.app.config.include_comparison = self.state.state.include_comparison
            self.gui.app.config.registry_sheet_name = self.state.state.registry_sheet_name
            self.gui.app.config.noise_registry_sheet_name = self.state.state.noise_registry_sheet_name
            self.gui.app.config.pressure_unit = self.state.state.pressure_unit
            self.gui.app.config.flow_unit = self.state.state.flow_unit
            self.gui.app.config.speed_unit = self.state.state.speed_unit
            self.gui.app.config.power_unit = self.state.state.power_unit
            self.gui.app.config.test_lab_root = self.state.state.test_lab_directory or None
            
            # Update progress
            self.gui.status_manager.update_status(
                f"Processing {len(tests_to_process)} test(s) for performance report...", 
                "blue"
            )
            
            # Create status callback for backend updates
            def status_callback(message, color='blue'):
                self.gui.status_manager.update_status(message, color)
            
            # Prepare fine-grained comparison test lab selection
            comparison_test_labs = {}
            if comparison_saps:
                for sap in comparison_saps:
                    selected_labs = self.state.state.selected_comparison_test_labs.get(sap, set())
                    if selected_labs:
                        comparison_test_labs[sap] = list(selected_labs)
                        logger.info(f"SAP {sap}: Will use selected test labs {selected_labs} for comparison")
                    else:
                        logger.info(f"SAP {sap}: No specific test labs selected, will use all tests")
            
            # Use the backend's generate_report method that properly separates the data
            self.gui.app.generate_report(
                selected_tests=tests_to_process,           # All tests go to performance sheet
                performance_saps=None,                     # Use all SAPs from selected tests for performance
                noise_saps=noise_saps if noise_saps else None,     # Only selected noise SAPs
                comparison_saps=comparison_saps if comparison_saps else None,  # Only selected comparison SAPs
                comparison_test_labs=comparison_test_labs if comparison_test_labs else None,  # Fine-grained test lab selection
                selected_noise_tests=selected_noise_tests if selected_noise_tests else None,  # Pre-filtered noise tests
                multiple_comparisons=multiple_comparisons if multiple_comparisons else None,  # Multiple comparison groups
                include_noise_data=self.state.state.include_noise,
                registry_sheet_name=self.state.state.registry_sheet_name,
                noise_registry_sheet_name=self.state.state.noise_registry_sheet_name
            )
            
            # Check if the report was actually generated
            from pathlib import Path
            output_file = Path(output_path)
            if output_file.exists():
                self.gui.status_manager.update_status(
                    f"✅ Report generated successfully! File: {output_file.name}", 
                    "green"
                )
                logger.info(f"Report generation completed successfully: {output_path}")
                
                # Show file location in status  
                self.gui.status_manager.update_status(
                    f"📁 Report saved to: {output_file.parent}", 
                    "blue"
                )
                
                # Hide any long-running progress indicator now that work is complete
                if hasattr(self.gui, 'status_manager'):
                    self.gui.status_manager.hide_progress()

                # Auto-refresh the Generate tab to update the UI after report generation
                # This ensures the user sees any state changes without manually refreshing
                if hasattr(self.gui, 'generate_tab'):
                    logger.info("🔄 Auto-refreshing Generate tab after successful report generation")
                    self.gui.generate_tab.refresh_content()
                
                # Update the page to show all changes
                if hasattr(self.gui, 'page'):
                    self.gui.page.update()
                else:
                    self.gui._safe_page_update()
                
                # Open the file
                self._open_report_file(output_file)
            else:
                self.gui.status_manager.update_status(
                    "Report generation completed but output file not found. Check the logs.", 
                    "orange"
                )
                logger.warning("Report generation completed but output file not found")
                
        except Exception as e:
            logger.error(f"Report generation failed: {e}", exc_info=True)
            self.gui.status_manager.update_status(
                f"Report generation failed: {str(e)}", 
                "red"
            )
    
    def _open_report_file(self, output_path: Path):
        """Open the generated report file"""
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                subprocess.run(['start', '', str(output_path)], shell=True, check=True)
            elif sys.platform == "darwin":
                subprocess.run(['open', str(output_path)], check=True)
            else:
                subprocess.run(['xdg-open', str(output_path)], check=True)
                
            logger.info(f"Opened report file: {output_path}")
            self.gui.status_manager.update_status(f"Opened report: {output_path.name}", "green")
            
        except Exception as e:
            logger.error(f"Failed to open report file: {e}")
            self.gui.status_manager.update_status(
                f"Report generated but failed to open: {str(e)}", 
                "orange"
            )

    def search_tests(self, query: str) -> List[Test]:
        """Search for tests using the backend"""
        if not hasattr(self.gui, 'app') or not self.gui.app:
            raise Exception("Backend not initialized")
        
        return self.gui.app.search_tests(query)
    
    def check_noise_data_exists(self, sap_code: str) -> bool:
        """Check if noise data exists for a SAP code"""
        if not hasattr(self.gui, 'app') or not self.gui.app:
            return False
        
        try:
            # For now, just return the configured state
            # TODO: Implement actual check in backend if needed
            return self.gui.app.config.include_noise
        except Exception as e:
            logger.warning(f"Failed to check noise data for {sap_code}: {e}")
            return False
    
    def get_available_features(self) -> dict:
        """Get available features based on backend configuration"""
        features = {
            'noise': False,
            'comparison': False
        }
        
        if not hasattr(self.gui, 'app') or not self.gui.app:
            return features
        
        try:
            # Check if noise registry and directory are configured
            config = self.gui.app.config
            features['noise'] = bool(config.noise_registry_path and config.noise_dir)
            features['comparison'] = True  # Always available if we have tests
            
        except Exception as e:
            logger.warning(f"Failed to determine available features: {e}")
        
        return features
    
    def validate_data_flow(self):
        """Validate that the data flow from frontend to backend is correct"""
        try:
            logger.info("=== VALIDATING DATA FLOW ===")
            
            # Check state manager data
            tests_to_process = self.state.get_tests_to_process()
            noise_saps = list(self.state.state.selected_noise_saps)
            comparison_saps = list(self.state.state.selected_comparison_saps)
            
            logger.info(f"Frontend state:")
            logger.info(f"  Selected tests: {len(tests_to_process)}")
            logger.info(f"  Test details: {[(t.test_lab_number, t.sap_code) for t in tests_to_process]}")
            logger.info(f"  Noise SAPs: {noise_saps}")
            logger.info(f"  Comparison SAPs: {comparison_saps}")
            
            # Validate consistency
            all_test_saps = set(t.sap_code for t in tests_to_process if t.sap_code)
            logger.info(f"  All test SAPs: {sorted(all_test_saps)}")
            
            # Check for invalid selections
            invalid_noise = set(noise_saps) - all_test_saps if noise_saps else set()
            invalid_comparison = set(comparison_saps) - all_test_saps if comparison_saps else set()
            
            if invalid_noise:
                logger.warning(f"❌ Invalid noise SAPs (not in selected tests): {invalid_noise}")
            else:
                logger.info(f"✅ All noise SAPs are valid")
                
            if invalid_comparison:
                logger.warning(f"❌ Invalid comparison SAPs (not in selected tests): {invalid_comparison}")
            else:
                logger.info(f"✅ All comparison SAPs are valid")
            
            logger.info("=== DATA FLOW VALIDATION COMPLETE ===")
            return True
            
        except Exception as e:
            logger.error(f"❌ Data flow validation failed: {e}")
            return False
    
    def validate_noise_tests_if_needed(self, callback):
        """
        Validate noise tests if noise data is included, then call callback.
        
        Args:
            callback: Function to call after noise test validation is complete.
                     Will be called with selected_noise_tests as argument.
        """
        # Check if noise tests should be included
        if not self.state.state.include_noise:
            logger.info("Noise tests not included in report, skipping validation")
            callback([])
            return
            
        # Check if we have required paths
        if not self.state.state.selected_tests_folder:
            logger.warning("No test folder selected, skipping noise validation")
            callback([])
            return
            
        if not self.state.state.selected_registry_file:
            logger.warning("No registry file selected, skipping noise validation")
            callback([])
            return
        
        try:
            # Update status
            self.gui.status_manager.update_status(
                "Validating noise test files...", 
                "blue"
            )
            
            # Import validator and info container
            from ...validators.noise_test_validator import NoiseTestValidator, NoiseTestValidationInfo
            from ...config.directory_config import NOISE_TEST_DIR, NOISE_REGISTRY_FILE
            from ...utils.common import normalize_sap_code
            
            # Use the noise test directory from config, fallback to looking in performance test folder
            noise_folder = str(NOISE_TEST_DIR) if NOISE_TEST_DIR else self.state.state.selected_tests_folder
            
            # Use the noise registry file from config, fallback to selected registry
            noise_registry_path = str(NOISE_REGISTRY_FILE) if NOISE_REGISTRY_FILE else self.state.state.selected_registry_file
            
            logger.info(f"Using noise directory: {noise_folder}")
            logger.info(f"Using noise registry: {noise_registry_path}")
            
            # Create validator
            validator = NoiseTestValidator(
                noise_folder,
                self.state.state.noise_registry_sheet_name
            )
            
            # Validate noise tests using the correct registry file
            noise_tests = validator.validate_from_registry(noise_registry_path)

            # Normalisation helper for test/lab identifiers
            def _normalize_test_id(value) -> str:
                text = "" if value is None else str(value).strip()
                if not text:
                    return ""
                try:
                    # Handle values like "1234.0" or floats cleanly
                    return str(int(float(text)))
                except (ValueError, TypeError):
                    # Fall back to stripping leading zeros without changing alphanumerics
                    stripped = text.lstrip('0')
                    return stripped or "0"

            # Inject placeholder entries for SAP/test combos chosen in the UI but missing in the registry
            tests_to_process = self.state.get_tests_to_process()
            existing_noise_keys = {
                (normalize_sap_code(nt.sap_code), _normalize_test_id(getattr(nt, 'test_no', None)))
                for nt in noise_tests
                if getattr(nt, 'sap_code', None) and getattr(nt, 'test_no', None)
            }

            supplemental_noise_entries: List[NoiseTestValidationInfo] = []

            for sap_code, test_ids in self.state.state.selected_noise_test_labs.items():
                if not test_ids:
                    continue

                normalized_sap = normalize_sap_code(sap_code)

                for test_id in test_ids:
                    normalized_test_id = _normalize_test_id(test_id)
                    if not normalized_test_id:
                        continue

                    key = (normalized_sap, normalized_test_id)
                    if key in existing_noise_keys:
                        continue

                    # Try to enrich with metadata from the selected performance tests
                    matching_performance_test = next(
                        (
                            perf_test
                            for perf_test in tests_to_process
                            if normalize_sap_code(perf_test.sap_code) == normalized_sap
                            and _normalize_test_id(perf_test.test_lab_number) == normalized_test_id
                        ),
                        None,
                    )

                    supplemental_entry = NoiseTestValidationInfo(
                        sap_code=sap_code,
                        test_no=normalized_test_id,
                        file_path="Selected via GUI (no registry entry)",
                        exists=True,
                        is_valid=True,
                        error_message=None,
                        file_size=None,
                        date=getattr(matching_performance_test, 'date', None) if matching_performance_test else None,
                        registry_row=None,
                        voltage=getattr(matching_performance_test, 'voltage', None) if matching_performance_test else None,
                        client=None,
                        application=None,
                        notes=getattr(matching_performance_test, 'notes', None) if matching_performance_test else None,
                        responsible=None,
                        test_lab=getattr(matching_performance_test, 'test_lab_number', None) if matching_performance_test else None,
                    )

                    supplemental_noise_entries.append(supplemental_entry)
                    existing_noise_keys.add(key)

            if supplemental_noise_entries:
                logger.info(
                    "Added %d supplemental noise test entries from GUI selections", 
                    len(supplemental_noise_entries)
                )
                noise_tests.extend(supplemental_noise_entries)
            
            logger.info(f"Noise test validation complete: {len(noise_tests)} tests found")
            
            # Check if we have any valid tests
            valid_tests = [t for t in noise_tests if t.is_valid]
            
            if not valid_tests:
                # No valid tests found, proceed without noise data
                logger.info("No valid noise tests found, continuing without noise data")
                self.gui.status_manager.update_status(
                    "No valid noise tests found. Continuing without noise data.", 
                    "orange"
                )
                callback([])
                return
            
            # Get selected noise tests from the state (fine-grained selection from config tab)
            selected_noise_tests = []

            # Normalise SAP codes once so we can handle format differences (spaces, case, etc.)
            normalized_noise_saps = {
                normalize_sap_code(sap) for sap in self.state.state.selected_noise_saps
                if sap
            }

            # Ensure any SAP that still has explicit test selections is considered noise-enabled.
            additional_noise_saps = {
                normalize_sap_code(sap)
                for sap in self.state.state.selected_noise_test_labs.keys()
                if sap
            }

            if additional_noise_saps:
                normalized_noise_saps.update(additional_noise_saps)
                # Keep the raw state in sync so later stages know these SAPs should carry noise data.
                for sap_code in self.state.state.selected_noise_test_labs.keys():
                    if sap_code and sap_code not in self.state.state.selected_noise_saps:
                        self.state.state.selected_noise_saps.add(sap_code)

            # Track combinations we have already added to avoid duplicates when registry rows repeat
            seen_noise_keys = set()

            for sap_code, selected_test_nos in self.state.state.selected_noise_test_labs.items():
                normalized_sap = normalize_sap_code(sap_code)

                if normalized_sap not in normalized_noise_saps:
                    continue

                # Normalise the selected test ids so "1234", "1234.0" and " 1234 " match
                normalized_selected_ids = {
                    _normalize_test_id(test_id)
                    for test_id in selected_test_nos
                    if _normalize_test_id(test_id)
                }

                for test in valid_tests:
                    if not test.test_no:
                        continue

                    normalized_test_sap = normalize_sap_code(test.sap_code)
                    normalized_test_id = _normalize_test_id(test.test_no)

                    if normalized_test_sap != normalized_sap:
                        continue

                    if normalized_test_id not in normalized_selected_ids:
                        continue

                    noise_key = (normalized_test_sap, normalized_test_id)
                    if noise_key in seen_noise_keys:
                        continue

                    seen_noise_keys.add(noise_key)
                    selected_noise_tests.append(test)
            
            logger.info(f"Using fine-grained noise test selection: {len(selected_noise_tests)} tests selected from config")
            
            if selected_noise_tests:
                self.gui.status_manager.update_status(
                    f"Using {len(selected_noise_tests)} noise tests from your selection", 
                    "green"
                )
            else:
                self.gui.status_manager.update_status(
                    "No noise tests selected in configuration. Continuing without noise data.", 
                    "orange"
                )
            
            # Continue with report generation using the selected tests
            callback(selected_noise_tests)
            
        except Exception as e:
            logger.error(f"Error during noise test validation: {e}")
            self.gui.status_manager.update_status(
                f"Error validating noise tests: {str(e)}. Continuing without noise data.", 
                "red"
            )
            # Continue without noise data
            callback([])

