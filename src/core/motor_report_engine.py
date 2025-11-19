#!/usr/bin/env python3
"""
Refactored Motor Performance Excel Report Generator v3.0

This script generates professional Excel reports from motor test CSV/INF files,
grouped by SAP code with advanced formatting, charts, summary sheets, and
image integration. This version focuses on improved organization,
modularity, and robustness.

Author: Motor Test Analysis Team (Refactored by AI)
Version: 3.0
"""

import logging
import argparse
import sys
import os
from pathlib import Path
from typing import List, Dict, Optional, Any, cast
from collections import defaultdict
from datetime import datetime

from ..config.app_config import AppConfig
from ..data.models import InfData, NoiseTestInfo, MotorTestData, Test, LifeTestInfo
from ..utils.common import open_file_externally, normalize_sap_code
from ..data.parsers import InfParser, CsvParser
from ..analysis.noise_handler import NoiseDataHandler
from ..reports.excel_report import ExcelReport
from ..config.directory_config import (
    OUTPUT_DIR,
    LOGS_DIR,
    LOGO_PATH,
    ensure_directories_initialized,
    invalidate_directory_cache,
)
from ..services.directory_locator import DirectoryLocator
from ..services.registry_service import RegistryService
from ..services.test_lab_summary_loader import TestLabSummaryLoader
from .telemetry import log_duration
from ..services.noise_registry_reader import (
    load_registry_dataframe,
    N_PROVA_STD,
    CODICE_SAP_STD,
    VOLTAGE_STD,
    NOTE_STD,
    ANNO_ORIGINAL_STD,
    ANNO_DATETIME_STD,
)

# Logging is configured in directory_config.py, so we just get the logger
logger = logging.getLogger(__name__)

try:
    import pandas as pd
except ImportError:
    logger.critical("pandas is required. Please install with 'pip install pandas'.")
    sys.exit(1)

# --- Global Configuration & Utilities ---

# --- Main Application Class ---
class MotorReportApp:
    """Main application class."""
    def __init__(self, config: AppConfig, noise_handler=None,
                 directory_locator: Optional[DirectoryLocator] = None,
                 registry_service: Optional[RegistryService] = None):
        self.config = config
        self.logger = logging.getLogger(__class__.__name__)
        self.inf_parser = InfParser()
        self.directory_locator = directory_locator or DirectoryLocator(
            logger=self.logger,
        )
        self.csv_parser = CsvParser()
        self.registry_service = registry_service or RegistryService(logger=self.logger)
        self.test_lab_loader: Optional[TestLabSummaryLoader] = None

        # Initialize paths first to populate the config object
        self._initialize_paths()
        self._initialize_test_lab_loader()

        # Use provided noise handler or create default one
        if noise_handler is not None:
            self.noise_handler = noise_handler
            self.logger.info("Using provided noise handler")
        else:
            # Fallback to default noise handler
            self.noise_handler = NoiseDataHandler(app_config=self.config)
            self.logger.info("Created default noise handler")

    def _initialize_paths(self):
        """
        Initializes and auto-detects necessary directory paths for the report generation.
        Updates the config object with detected paths.
        """
        self.directory_locator.apply_defaults(self.config)

        self.logger.info(f"Using tests folder: {self.config.tests_folder}")
        self.logger.info(f"Using noise directory: {self.config.noise_dir}")
        self.logger.info(f"Output path: {self.config.output_path}")

    def _initialize_test_lab_loader(self) -> None:
        base_path = getattr(self.config, "test_lab_root", None)
        if not base_path:
            self.logger.info("Test lab workbook directory not configured; skipping summary loader setup.")
            return

        loader_logger = logging.getLogger(f"{self.__class__.__name__}.TestLab")
        self.test_lab_loader = TestLabSummaryLoader(base_path, logger_=loader_logger)

        if not self.test_lab_loader.available:
            self.logger.info("Test lab workbook directory '%s' is not accessible; summaries disabled.", base_path)
            self.test_lab_loader = None

    # --- Core Logic Methods ---

    def _find_inf_files(self) -> List[Path]:
        """Finds all .inf files in the configured tests folder."""
        if not self.config.tests_folder or not os.path.isdir(self.config.tests_folder):
            self.logger.error(f"Tests folder not found or not a directory: {self.config.tests_folder}")
            return []
        return list(Path(self.config.tests_folder).glob("*.inf"))

    def _group_inf_files_by_sap(self, inf_files: List[Path]) -> Dict[str, List[str]]:
        """Groups INF files by SAP code found within them."""
        grouped = defaultdict(list)
        for inf_file in inf_files:
            inf_data = self.inf_parser.parse(inf_file)
            if inf_data and inf_data.motor_type:
                grouped[inf_data.motor_type].append(inf_file.stem)
        return dict(sorted(grouped.items()))

    def _process_selected_saps(self, sap_codes: List[str], all_inf_tests_by_sap: Dict[str, List[str]]) -> List[MotorTestData]:
        """Processes all tests for the selected SAP codes."""
        processed_tests = []
        with log_duration(self.logger, "process_selected_saps"):
            for sap_code in sap_codes:
                test_numbers = all_inf_tests_by_sap.get(sap_code, [])
                self.logger.info(f"Processing {len(test_numbers)} tests for SAP: {sap_code}")
                for test_number in test_numbers:
                    test_data = self._process_single_test(test_number)
                    if test_data:
                        processed_tests.append(test_data)
        return processed_tests

    def _process_single_test(self, test_number: str) -> Optional[MotorTestData]:
        self.logger.debug("Processing test: %s", test_number)
        if not self.config.tests_folder:
            self.logger.error("Tests folder is not configured.")
            return None

        base_dir = Path(self.config.tests_folder)

        def _resolve_variant_file(number: str, extension: str) -> tuple[Path, Optional[str]]:
            """Return path to file for *number* (possibly reusing base variant)."""
            direct_path = base_dir / f"{number}{extension}"
            if direct_path.exists():
                return direct_path, number

            if number.upper().endswith("A"):
                base_number = number[:-1]
                if base_number:
                    base_path = base_dir / f"{base_number}{extension}"
                    if base_path.exists():
                        self.logger.info(
                            "Reusing %s for variant %s", base_path.name, number
                        )
                        return base_path, base_number

            return direct_path, None

        inf_path, inf_source_number = _resolve_variant_file(test_number, ".inf")
        csv_path, csv_source_number = _resolve_variant_file(test_number, ".csv")

        status_notes: List[str] = []

        inf_data = InfData()
        if inf_path.exists():
            inf_data = self.inf_parser.parse(inf_path)
            if inf_source_number and inf_source_number != test_number:
                status_notes.append(f"INF data reused from {inf_source_number}")
        else:
            status_notes.append("INF missing")

        if not inf_data.motor_type and inf_path.exists():
            status_notes.append("INF parsed with missing motor type")

        csv_data = None
        if csv_path.exists():
            csv_data = self.csv_parser.parse(csv_path)
            if csv_data is None:
                status_notes.append("CSV failed to parse")
            elif csv_source_number and csv_source_number != test_number:
                status_notes.append(f"CSV data reused from {csv_source_number}")
        else:
            status_notes.append("CSV missing")

        noise_info: Optional[NoiseTestInfo] = None
        noise_status_summary = "SKIPPED"

        if self.config.include_noise:
            noise_status_summary = "NOT FOUND"
            lookup_number_for_noise = inf_source_number or test_number
            if inf_data and inf_data.motor_type:
                current_sap_code = normalize_sap_code(inf_data.motor_type)
                test_date_for_noise = inf_data.date if inf_data and inf_data.date else None
                noise_info = self.noise_handler.get_noise_test_info(
                    sap_code=current_sap_code,
                    test_date_str=test_date_for_noise,
                    test_number_for_lab_matching=lookup_number_for_noise,
                )
                if noise_info:
                    noise_status_summary = (
                        "OK" if noise_info.image_paths else "Info OK, No Images"
                    )

        test_lab_summary = None
        if self.test_lab_loader and self.test_lab_loader.available:
            # Check for override in config
            override_path = self.config.selected_carichi_map.get(test_number)
            
            self.logger.debug(
                "Fetching test-lab summary for %s using base %s (override: %s)",
                test_number,
                getattr(self.test_lab_loader, "base_path", None),
                override_path
            )
            test_lab_summary = self.test_lab_loader.load_summary(test_number, override_path=override_path)
            if test_lab_summary:
                self.logger.info(
                    "Test %s linked to workbook %s (scheda=%s, collaudo=%s)",
                    test_number,
                    test_lab_summary.source_path,
                    "yes" if test_lab_summary.scheda else "no",
                    "yes" if test_lab_summary.collaudo_media else "no",
                )
            else:
                self.logger.info("No test-lab summary available for %s", test_number)

        status_message = "; ".join(status_notes) if status_notes else "OK"

        return MotorTestData(
            test_number=test_number,
            inf_data=inf_data,
            csv_path=csv_path,
            csv_data=csv_data,
            noise_info=noise_info,
            noise_status=noise_status_summary,
            status_message=status_message.strip(),
            test_lab_summary=test_lab_summary,
        )

    def _generate_report(self, status_callback=None):
        """Generates the final Excel report."""
        output_path = self.config.output_path
        if not output_path:
            self.logger.error("Output path is not configured. Cannot generate report.")
            if status_callback: status_callback("Error: Output path not configured.", color='red')
            return

        self.logger.info(f"Generating report with {len(self.all_motor_test_data)} test data objects")
        grouped_by_sap = self._group_tests_by_sap(self.all_motor_test_data)
        self.logger.info(f"Tests grouped by SAP: {list(grouped_by_sap.keys())}")
        for sap, tests in grouped_by_sap.items():
            self.logger.info(f"  SAP {sap}: {len(tests)} tests - {[t.test_number for t in tests]}")
        
        all_noise_tests_by_sap = defaultdict(list)
        if self.config.include_noise:
            with log_duration(self.logger, "collect_noise_tests"):
                # Check if we have pre-filtered noise tests from the GUI
                if self.config.selected_noise_tests is not None:
                    self.logger.info(f"Using pre-filtered noise tests from GUI: {len(self.config.selected_noise_tests)} tests")
                    
                    # Group the pre-filtered noise tests by SAP code and convert to NoiseTestInfo
                    for noise_test in self.config.selected_noise_tests:
                        # Get SAP code from the noise test
                        sap_code = getattr(noise_test, 'sap_code', None)
                        if sap_code:
                            # Convert NoiseTestValidationInfo to NoiseTestInfo for compatibility
                            from ..data.models import NoiseTestInfo
                            
                            # Extract year from date for file finding
                            date_value = getattr(noise_test, 'date', None)
                            year = self._extract_year_from_date(date_value)
                            
                            # Get actual file paths (images AND TXT) using the noise handler
                            noise_info_data = None
                            test_number = getattr(noise_test, 'test_no', None)
                            if test_number and year and self.noise_handler:
                                try:
                                    method = getattr(self.noise_handler, 'get_noise_test_info_by_test_year', None)
                                    if callable(method):
                                        # Using SimplifiedNoiseDataHandler - gets BOTH images AND TXT files
                                        noise_info_data = cast(Any, method(
                                            str(test_number),
                                            str(year),
                                            sap_code=sap_code,
                                        ))
                                        if noise_info_data:
                                            self.logger.info(f"Found data for test {test_number}: {noise_info_data.data_type}, "
                                                           f"{len(noise_info_data.image_paths)} images, "
                                                           f"{len(noise_info_data.txt_files)} TXT files")
                                    else:
                                        # Fallback: Using regular NoiseDataHandler
                                        comprehensive_data = self.noise_handler.get_noise_data_comprehensive(str(test_number), str(year))
                                        if comprehensive_data:
                                            from ..data.models import NoiseTestInfo
                                            noise_info_data = NoiseTestInfo(
                                                nprova=test_number,
                                                year=year,
                                                image_paths=[Path(img) for img in comprehensive_data.get('images', [])],
                                                txt_files=[Path(txt) for txt in comprehensive_data.get('txt_files', [])],
                                                data_type=comprehensive_data.get('type', 'none')
                                            )
                                except Exception as e:
                                    self.logger.warning(f"Failed to get data for test {test_number}, year {year}: {e}")
                            
                            # Create NoiseTestInfo with discovered file data
                            if noise_info_data:
                                # Use the discovered data
                                from ..data.models import NoiseTestInfo
                                noise_info = NoiseTestInfo(
                                    test_number=getattr(noise_test, 'test_no', None),
                                    date=getattr(noise_test, 'date', None),
                                    mic_position=None,
                                    background_noise=None,
                                    motor_noise=None,
                                    result=None,
                                    image_path=noise_info_data.image_paths[0] if noise_info_data.image_paths else None,
                                    nprova=getattr(noise_test, 'test_no', None),
                                    test_lab=getattr(noise_test, 'test_lab', None),
                                    year=year,
                                    image_paths=noise_info_data.image_paths,  # Images from discovery
                                    txt_files=noise_info_data.txt_files,  # TXT files from discovery
                                    data_type=noise_info_data.data_type,  # Data type from discovery
                                    sap_code=sap_code
                                )
                            else:
                                # No data found - create empty NoiseTestInfo
                                from ..data.models import NoiseTestInfo
                                noise_info = NoiseTestInfo(
                                    test_number=getattr(noise_test, 'test_no', None),
                                    date=getattr(noise_test, 'date', None),
                                    mic_position=None,
                                    background_noise=None,
                                    motor_noise=None,
                                    result=None,
                                    image_path=None,
                                    nprova=getattr(noise_test, 'test_no', None),
                                    test_lab=getattr(noise_test, 'test_lab', None),
                                    year=year,
                                    image_paths=[],
                                    txt_files=[],
                                    data_type="none",
                                    sap_code=sap_code
                                )
                            all_noise_tests_by_sap[sap_code].append(noise_info)
                            
                    self.logger.info(f"Pre-filtered noise data by SAP: {dict(all_noise_tests_by_sap)}")
                else:
                    # Original logic: collect noise tests from motor test data
                    self.logger.info("No pre-filtered noise tests provided, using original lookup logic")
                    allowed_noise_saps = set(self.config.noise_saps) if self.config.noise_saps else set()
                    for test in self.all_motor_test_data:
                        if test.sap_code and test.noise_info:
                            # Only include if noise SAPs are selected AND this SAP is in the selection
                            if allowed_noise_saps and test.sap_code in allowed_noise_saps:
                                all_noise_tests_by_sap[test.sap_code].append(test.noise_info)
                    self.logger.info(f"Noise data by SAP (original logic): {dict(all_noise_tests_by_sap)}")

        comparison_data = {}
        if self.config.include_comparison:
            self.logger.info(f"Comparison enabled. compare_saps: {self.config.compare_saps}")
            self.logger.info(f"Comparison test labs: {getattr(self.config, 'comparison_test_labs', {})}")
            
            # Filter comparison data by selected test lab numbers
            for sap in self.config.compare_saps:
                if sap in grouped_by_sap:
                    sap_tests = grouped_by_sap[sap]
                    
                    # If specific test lab numbers are selected for this SAP, filter them
                    selected_test_labs = getattr(self.config, 'comparison_test_labs', {}).get(sap, set())
                    if selected_test_labs:
                        # Convert to set for faster lookup
                        selected_test_labs_set = set(selected_test_labs) if isinstance(selected_test_labs, (list, tuple)) else selected_test_labs
                        filtered_tests = [test for test in sap_tests if test.test_number in selected_test_labs_set]
                        self.logger.info(f"  SAP {sap}: Filtered from {len(sap_tests)} to {len(filtered_tests)} tests based on selection: {selected_test_labs_set}")
                        comparison_data[sap] = filtered_tests
                    else:
                        # No specific selection, include all tests for this SAP
                        self.logger.info(f"  SAP {sap}: No specific test selection, including all {len(sap_tests)} tests")
                        comparison_data[sap] = sap_tests
            
            self.logger.info(f"Final comparison data: {list(comparison_data.keys())}")
            for sap, tests in comparison_data.items():
                self.logger.info(f"  Comparison SAP {sap}: {len(tests)} tests - {[t.test_number for t in tests]}")
        else:
            self.logger.info("Comparison disabled")

        report_generator = ExcelReport(self.config, noise_handler=self.noise_handler)

        # Log what's being passed to the Excel writer
        self.logger.info("=== Data being passed to Excel writer ===")
        self.logger.info(f"grouped_data keys: {list(grouped_by_sap.keys())}")
        self.logger.info(f"all_tests_summary count: {len(self.all_motor_test_data)}")
        self.logger.info(f"all_noise_tests_by_sap keys: {list(all_noise_tests_by_sap.keys())}")
        self.logger.info(f"comparison_data keys: {list(comparison_data.keys())}")
        
        # Use multiple comparisons if available, otherwise fall back to legacy single comparison
        multiple_comparisons = getattr(self.config, 'multiple_comparisons', [])
        if multiple_comparisons:
            self.logger.info(f"multiple_comparisons: {len(multiple_comparisons)} groups")
        
        # Collect Life Test (LF) data if selected
        lf_tests_by_sap = {}
        self.logger.info("=" * 60)
        self.logger.info("🔬 LIFE TEST (LF) DATA COLLECTION")
        self.logger.info("=" * 60)
        
        # Check if config has LF attribute
        if not hasattr(self.config, 'selected_lf_test_numbers'):
            self.logger.warning("⚠️ Config does not have 'selected_lf_test_numbers' attribute")
            self.logger.info("📋 Available config attributes: " + ", ".join(dir(self.config)))
        else:
            self.logger.info(f"✅ Config has 'selected_lf_test_numbers' attribute")
            self.logger.info(f"📊 Type: {type(self.config.selected_lf_test_numbers)}")
            self.logger.info(f"📊 Value: {self.config.selected_lf_test_numbers}")
            self.logger.info(f"📊 Is empty: {not self.config.selected_lf_test_numbers}")
            
            if self.config.selected_lf_test_numbers:
                self.logger.info(f"🎯 Processing Life Test (LF) selections for {len(self.config.selected_lf_test_numbers)} SAP code(s)...")
                from ..services.lf_registry_reader import LifeTestRegistryReader
                lf_reader = LifeTestRegistryReader()
                
                for sap_code, test_numbers in self.config.selected_lf_test_numbers.items():
                    self.logger.info(f"📌 SAP {sap_code}: {len(test_numbers) if test_numbers else 0} test(s) selected")
                    if not test_numbers:
                        self.logger.warning(f"   ⚠️ No test numbers for SAP {sap_code}")
                        continue
                    
                    lf_tests_for_sap = []
                    for test_number in test_numbers:
                        self.logger.info(f"   🔍 Looking up LF test: {test_number}")
                        lf_test = lf_reader.get_test_by_number(test_number)
                        if not lf_test:
                                # Try an index-based suggestion as a fallback (helps when registry
                                # entries are malformed or filenames changed). This will NOT
                                # modify the registry but will allow the report to include
                                # a hyperlink when a likely file is found.
                                try:
                                    test_id, year = lf_reader.parse_test_number(test_number)
                                    candidate = None
                                    indexer = getattr(lf_reader, 'indexer', None)
                                    if indexer is not None and test_id and hasattr(indexer, 'get_best_file'):
                                        candidate = indexer.get_best_file(test_id, year)
                                    if candidate:
                                        lf_test = LifeTestInfo(
                                            test_number=test_number,
                                            sap_code=sap_code,
                                            notes=None,
                                            year=year,
                                            test_id=test_id,
                                            file_path=candidate,
                                            file_exists=True,
                                            responsible=None
                                        )
                                        self.logger.info(f"   🔎 Suggested file for {test_number}: {candidate}")
                                    else:
                                        self.logger.warning(f"   ⚠️ Could not find LF test {test_number} in registry or via index")
                                except Exception as e:
                                    self.logger.debug(f"Index suggestion failed for {test_number}: {e}")

                        if lf_test:
                            lf_tests_for_sap.append(lf_test)
                            file_status = "✅ exists" if lf_test.file_exists else "❌ not found"
                            self.logger.info(f"   ✅ Added LF test {test_number} - File: {file_status}")
                    
                    if lf_tests_for_sap:
                        lf_tests_by_sap[sap_code] = lf_tests_for_sap
                        self.logger.info(f"✅ Collected {len(lf_tests_for_sap)} LF test(s) for SAP {sap_code}")
                    else:
                        self.logger.warning(f"⚠️ No valid LF tests found for SAP {sap_code}")
                
                total_tests = sum(len(tests) for tests in lf_tests_by_sap.values())
                if total_tests > 0:
                    self.logger.info("=" * 60)
                    self.logger.info(f"🎉 Total LF tests: {total_tests} across {len(lf_tests_by_sap)} SAP code(s)")
                    self.logger.info("=" * 60)
                else:
                    self.logger.warning("⚠️ No LF tests collected (all SAP codes had no valid tests)")
            else:
                self.logger.info("ℹ️ No Life Test (LF) selections found in config")
                self.logger.info("   This is normal if no LF tests were selected in the Config tab")
        
        self.logger.info("")  # Empty line for readability
        
        # Final check before passing to Excel generator
        self.logger.info("=" * 60)
        self.logger.info("📊 FINAL DATA CHECK BEFORE EXCEL GENERATION")
        self.logger.info("=" * 60)
        self.logger.info(f"📋 lf_tests_by_sap type: {type(lf_tests_by_sap)}")
        self.logger.info(f"📋 lf_tests_by_sap keys: {list(lf_tests_by_sap.keys()) if lf_tests_by_sap else 'EMPTY'}")
        if lf_tests_by_sap:
            for sap, tests in lf_tests_by_sap.items():
                self.logger.info(f"   📌 SAP {sap}: {len(tests)} test(s)")
                for test in tests:
                    self.logger.info(f"      • {test.test_number} - Path: {test.file_path}")
        else:
            self.logger.warning("⚠️ lf_tests_by_sap is EMPTY - no hyperlinks will be created!")
        self.logger.info("=" * 60)

        with log_duration(self.logger, "excel_report_generate"):
            generation_success = report_generator.generate(
                grouped_data=grouped_by_sap, 
                all_tests_summary=self.all_motor_test_data, 
                all_noise_tests_by_sap=all_noise_tests_by_sap,
                comparison_data=comparison_data,
                multiple_comparisons=multiple_comparisons,
                lf_tests_by_sap=lf_tests_by_sap
            )

        final_output_path = self.config.output_path or output_path

        if generation_success:
            self.logger.info(f"Excel report generated: {final_output_path}")
            if status_callback: 
                status_callback(f"Report successfully generated at: {final_output_path}", color='green')
            else: 
                self.logger.info(f"[SUCCESS] Report generated: {final_output_path}")
            
            # Only open file automatically in CLI mode, not in GUI mode
            if self.config.open_after_creation and not status_callback:
                open_file_externally(final_output_path)
        else:
            msg = "Report generation failed. Check log for details."
            self.logger.error("Report generation failed - invalidating directory cache for next startup")
            invalidate_directory_cache()
            if status_callback: 
                status_callback(msg, color='red')
            else: 
                self.logger.error(f"[FAILURE] {msg}")

    def _group_tests_by_sap(self, motor_tests: List[MotorTestData]) -> Dict[str, List[MotorTestData]]:
        """Groups motor test data by SAP code."""
        grouped: Dict[str, List[MotorTestData]] = defaultdict(list)
        for test_data in motor_tests:
            sap = test_data.sap_code
            if sap and sap != 'UNKNOWN_SAP':
                # Include tests even if CSV data is None for debugging
                grouped[sap].append(test_data)
                self.logger.debug(f"Added test {test_data.test_number} to SAP {sap} (CSV data present: {test_data.csv_data is not None})")
            else:
                self.logger.warning(f"Test {test_data.test_number} has no valid SAP code: {sap}")
        return grouped

    def load_registry(self) -> List[Test]:
        """Loads the test registry from the Excel file with caching for better performance."""
        import os
        from pathlib import Path
        
        try:
            return self.registry_service.load_tests(self.config)
        except Exception as exc:
            self.logger.error("Failed to load registry: %s", exc)
            return []

    def find_tests_by_sap(self, sap_code: str) -> List[Test]:
        """Finds tests in the registry by SAP code."""
        all_tests = self.load_registry()
        normalized_sap = normalize_sap_code(sap_code)
        return [test for test in all_tests if normalize_sap_code(test.sap_code) == normalized_sap]

    def find_test_by_number(self, test_number: str) -> Optional[Test]:
        """Finds a test in the registry by its lab number using exact matching."""
        all_tests = self.load_registry()
        # Use exact matching - no padding or manipulation
        for test in all_tests:
            if test.test_lab_number == test_number:
                return test
        return None

    def search_tests(self, query: str) -> List[Test]:
        """
        Enhanced search for tests by SAP code or test number with smart input recognition.
        Supports multiple comma-separated inputs.
        Returns a list of matching tests from the registry.
        """
        if not query:
            return []
        
        query = query.strip()
        all_tests = self.load_registry()
        matching_tests = []
        
        # Split by comma for multiple inputs
        input_items = [item.strip() for item in query.split(',') if item.strip()]
        
        for item in input_items:
            self.logger.info(f"Processing search item: '{item}'")
            
            # Try exact match as test number first (with zero-padding)
            test_number_matches = self._find_by_test_number(all_tests, item)
            if test_number_matches:
                self.logger.info(f"'{item}' matched as test number: found {len(test_number_matches)} test(s)")
                matching_tests.extend(test_number_matches)
                continue
            
            # Try exact match as SAP code
            sap_code_matches = self._find_by_sap_code(all_tests, item)
            if sap_code_matches:
                self.logger.info(f"'{item}' matched as SAP code: found {len(sap_code_matches)} test(s)")
                matching_tests.extend(sap_code_matches)
                continue
            
            self.logger.warning(f"No exact matches found for '{item}'")
        
        # Remove duplicates while preserving order
        seen = set()
        unique_tests = []
        for test in matching_tests:
            test_key = (test.test_lab_number, test.sap_code)
            if test_key not in seen:
                seen.add(test_key)
                unique_tests.append(test)
        
        self.logger.info(f"Search for '{query}' found {len(unique_tests)} unique test(s)")
        return unique_tests

    def _find_by_test_number(self, all_tests: List[Test], input_item: str) -> List[Test]:
        """Find tests by exact test number match only"""
        matches = []
        # Use exact matching only - no padding or manipulation
        for test in all_tests:
            if test.test_lab_number == input_item:
                matches.append(test)
        return matches

    def _find_by_sap_code(self, all_tests: List[Test], input_item: str) -> List[Test]:
        """Find tests by exact SAP code match"""
        matches = []
        normalized_input = normalize_sap_code(input_item)
        for test in all_tests:
            if normalize_sap_code(test.sap_code) == normalized_input:
                matches.append(test)
        
        return matches

    def analyze_search_input(self, query: str) -> dict:
        """
        Analyze the search input to categorize by type and provide search strategy.
        Returns information about what was found and how to present it to the user.
        """
        if not query:
            return {"status": "empty", "message": "No search query provided"}
        
        query = query.strip()
        all_tests = self.load_registry()
        input_items = [item.strip() for item in query.split(',') if item.strip()]
        
        analysis = {
            "status": "success",
            "total_inputs": len(input_items),
            "test_number_inputs": [],
            "sap_code_inputs": [],
            "unmatched_inputs": [],
            "found_tests": [],
            "search_strategy": "mixed"  # Can be 'test_numbers_only', 'sap_codes_only', or 'mixed'
        }
        
        for item in input_items:
            test_matches = self._find_by_test_number(all_tests, item)
            sap_matches = self._find_by_sap_code(all_tests, item)
            
            if test_matches:
                analysis["test_number_inputs"].append({
                    "input": item,
                    "matches": test_matches,
                    "type": "test_number"
                })
                analysis["found_tests"].extend(test_matches)
            elif sap_matches:
                analysis["sap_code_inputs"].append({
                    "input": item,
                    "matches": sap_matches,
                    "type": "sap_code"
                })
                analysis["found_tests"].extend(sap_matches)
            else:
                analysis["unmatched_inputs"].append(item)
        
        # Determine search strategy
        if analysis["test_number_inputs"] and not analysis["sap_code_inputs"]:
            analysis["search_strategy"] = "test_numbers_only"
        elif analysis["sap_code_inputs"] and not analysis["test_number_inputs"]:
            analysis["search_strategy"] = "sap_codes_only"
        else:
            analysis["search_strategy"] = "mixed"
        
        # Remove duplicates from found tests
        seen = set()
        unique_tests = []
        for test in analysis["found_tests"]:
            test_key = (test.test_lab_number, test.sap_code)
            if test_key not in seen:
                seen.add(test_key)
                unique_tests.append(test)
        
        analysis["found_tests"] = unique_tests
        analysis["total_found"] = len(unique_tests)
        
        return analysis

    def run_with_selected_tests(self, tests: List[Test], comparison_saps: Optional[List[str]], status_callback=None):
        """
        Runs the report generation process with a specific list of selected tests.
        """
        self.logger.info(f"Running report for {len(tests)} selected tests.")
        self.logger.info(f"Input tests: {[(t.test_lab_number, t.sap_code) for t in tests]}")
        self.logger.info(f"Input comparison SAPs: {comparison_saps}")
        
        if status_callback:
            status_callback(f"Processing {len(tests)} selected tests...")

        # Process only the selected tests and filter out None results immediately
        processed_tests = []
        for test in tests:
            self.logger.info(f"Processing test {test.test_lab_number} with SAP {test.sap_code}")
            # Use test_lab_number to find the corresponding INF/CSV files
            test_data = self._process_single_test(test.test_lab_number)
            if test_data:
                # Validate SAP code consistency between registry and INF file
                if test_data.sap_code != test.sap_code:
                    self.logger.warning(f"SAP code mismatch for test {test.test_lab_number}: "
                                      f"Registry={test.sap_code}, INF={test_data.sap_code}")
                    # Use the registry SAP code as authoritative
                    test_data.inf_data.motor_type = test.sap_code
                self.logger.info(f"Successfully processed test {test.test_lab_number} with final SAP: {test_data.sap_code}")
                processed_tests.append(test_data)
            else:
                self.logger.warning(f"Could not process test {test.test_lab_number} from registry")
        
        self.all_motor_test_data = processed_tests
        self.logger.info(f"Final processed tests: {[(t.test_number, t.sap_code) for t in self.all_motor_test_data]}")

        if not self.all_motor_test_data:
            msg = "No valid test data could be processed from the selection."
            self.logger.error(msg)
            if status_callback:
                status_callback(msg, color='red')
            return

        # Set comparison SAPs if provided
        if self.config.include_comparison and comparison_saps:
            self.config.compare_saps = comparison_saps
            self.logger.info(f"Comparison enabled. Set compare_saps to: {self.config.compare_saps}")
        else:
            self.config.compare_saps = []
            self.logger.info("Comparison disabled or no comparison SAPs provided")

        self._generate_report(status_callback=status_callback)

    def find_and_group_saps(self) -> Dict[str, List[str]]:
        """Finds all .inf files and groups them by SAP code, for GUI use."""
        self.logger.info("Searching for .inf files to group by SAP...")
        inf_files = self._find_inf_files()
        if not inf_files:
            self.logger.warning("No .inf files found in the specified directory.")
            return {}
        grouped_saps = self._group_inf_files_by_sap(inf_files)
        self.logger.info(f"Found {len(grouped_saps)} SAP codes.")
        return grouped_saps

    def run(self, status_callback=None):
        """
        Main execution method for the application, driven by selected SAP codes.
        This is the primary entry point for the CLI-based workflow.
        """
        self.logger.info("Starting motor report generation process...")
        if status_callback:
            status_callback("Finding all available tests...")

        all_inf_tests_by_sap = self.find_and_group_saps()
        if not all_inf_tests_by_sap:
            msg = "No SAP codes found. Exiting."
            self.logger.warning(msg)
            if status_callback:
                status_callback(msg, color='orange')
            return

        # In CLI mode, we might pre-select SAPs via args
        saps_to_process = self.config.sap_codes or list(all_inf_tests_by_sap.keys())
        
        self.logger.info(f"Processing data for SAP codes: {saps_to_process}")
        if status_callback:
            status_callback(f"Processing {len(saps_to_process)} SAP code(s)...")

        self.all_motor_test_data = self._process_selected_saps(saps_to_process, all_inf_tests_by_sap)

        if not self.all_motor_test_data:
            msg = "No valid test data could be processed."
            self.logger.error(msg)
            if status_callback:
                status_callback(msg, color='red')
            return

        self._generate_report(status_callback=status_callback)

    def generate_report(self, selected_tests: List[Test], performance_saps=None, noise_saps=None, comparison_saps=None, comparison_test_labs=None, selected_noise_tests=None, multiple_comparisons=None, include_noise_data=True, registry_sheet_name=None, noise_registry_sheet_name=None, carichi_matches=None):
        """
        Public method to generate a report with selected SAP codes for each feature.
        
        Args:
            selected_tests: List of Test objects to include in the report
            performance_saps: SAP codes for performance sheets (None = use all from selected_tests)
            noise_saps: SAP codes to include noise data for
            comparison_saps: SAP codes to include in comparison sheet  
            comparison_test_labs: Dict[str, Set[str]] mapping SAP codes to specific test lab numbers for comparison
            selected_noise_tests: Pre-filtered list of NoiseTestInfo objects to use (bypasses noise lookup)
            multiple_comparisons: List of multiple comparison group definitions
            include_noise_data: Whether to include noise data
            registry_sheet_name: Name of the registry sheet
            noise_registry_sheet_name: Name of the noise registry sheet
            carichi_matches: Dict[str, str] mapping test numbers to specific Carichi file paths
        """
        # Set config options
        self.config.include_noise = include_noise_data
        self.config.registry_sheet_name = registry_sheet_name or self.config.registry_sheet_name
        self.config.noise_registry_sheet_name = noise_registry_sheet_name or self.config.noise_registry_sheet_name
        self.config.compare_saps = comparison_saps or []
        
        # Store carichi matches if provided
        if carichi_matches:
            self.config.selected_carichi_map = carichi_matches
            logger.info(f"Using manual Carichi file selections for {len(carichi_matches)} tests")
        else:
            self.config.selected_carichi_map = {}
        
        # Store multiple comparisons if provided
        if multiple_comparisons:
            self.config.multiple_comparisons = multiple_comparisons
            logger.info(f"Using multiple comparisons: {len(multiple_comparisons)} groups")
        else:
            self.config.multiple_comparisons = []
        
        # Store fine-grained comparison test lab selection
        if comparison_test_labs:
            self.config.comparison_test_labs = comparison_test_labs
        else:
            self.config.comparison_test_labs = {}
        
        # Store pre-filtered noise tests if provided
        if selected_noise_tests is not None:
            self.config.selected_noise_tests = selected_noise_tests
            logger.info(f"Using pre-filtered noise tests: {len(selected_noise_tests)} tests")
        else:
            self.config.selected_noise_tests = None
        
        # Optionally, store noise_saps in config if you want to use it elsewhere
        if noise_saps is not None:
            self.config.noise_saps = noise_saps
        # Filter selected_tests by performance_saps if provided
        filtered_tests = selected_tests
        if performance_saps:
            filtered_tests = [t for t in filtered_tests if t.sap_code in performance_saps]
        # Optionally, filter noise tests by noise_saps if needed (not implemented here)
        # Call the main workflow
        self.run_with_selected_tests(filtered_tests, comparison_saps)

    def _extract_year_from_date(self, date_value):
        """Extract year from various date formats."""
        import re
        if pd.isna(date_value) or date_value == '' or date_value is None:
            return None
        
        try:
            # Convert to string for processing
            date_str = str(date_value).strip()
            
            # Try direct year conversion (4-digit number)
            if date_str.isdigit() and len(date_str) == 4:
                year = int(date_str)
                if 1900 <= year <= 2100:
                    return str(year)
            
            # Try parsing as datetime first
            try:
                if isinstance(date_value, datetime):
                    return str(date_value.year)
                # Try different date formats
                dt = pd.to_datetime(date_value, errors='raise', dayfirst=True)
                return str(dt.year)
            except:
                pass
            
            # Extract 4-digit year using regex
            year_match = re.search(r'\b(19|20)\d{2}\b', date_str)
            if year_match:
                return year_match.group()
            
            # Try Excel date numbers
            try:
                if date_str.replace('.', '').isdigit():
                    excel_date = pd.to_datetime(float(date_str), origin='1899-12-30', unit='D')
                    return str(excel_date.year)
            except:
                pass
                
        except Exception as e:
            self.logger.debug(f"Could not extract year from '{date_value}': {e}")
        
        return None

def main():
    """Command-line interface setup and execution."""
    parser = argparse.ArgumentParser(description="Motor Performance Excel Report Generator")
    parser.add_argument("-t", "--tests-folder", help="Path to the folder containing test data.")
    parser.add_argument("-r", "--registry-path", help="Path to the registry file.")
    parser.add_argument("-o", "--output-path", help="Path to save the generated Excel report.")
    parser.add_argument("-l", "--logo-path", help="Path to the company logo image.")
    parser.add_argument("-s", "--sap", nargs='+', help="Specific SAP codes to process.")
    parser.add_argument("-c", "--compare", nargs='+', help="Specific SAP codes to include in the comparison sheet.")
    parser.add_argument("--no-noise", action="store_true", help="Disable noise data processing.")
    parser.add_argument("--no-comparison", action="store_true", help="Disable the comparison sheet generation.")
    parser.add_argument("--no-open", action="store_true", help="Do not open the report automatically after creation.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose logging.")

    args = parser.parse_args()

    # Ensure auto-discovered directories (performance, noise, CARICHI, etc.) are populated
    try:
        ensure_directories_initialized()
    except Exception as exc:  # pragma: no cover - defensive guard for CLI mode
        logger.warning("Directory auto-detection failed during CLI startup: %s", exc)

    locator = DirectoryLocator()

    default_tests_dir = locator.performance_dir
    default_registry_file = locator.lab_registry_file
    default_test_lab_dir = locator.test_lab_dir

    default_output_dir = locator.output_dir
    if not default_output_dir and OUTPUT_DIR:
        default_output_dir = Path(OUTPUT_DIR)

    default_logs_dir = locator.logs_dir
    if not default_logs_dir and LOGS_DIR:
        default_logs_dir = Path(LOGS_DIR)

    default_logo_path = locator.logo_path or (Path(LOGO_PATH) if LOGO_PATH else None)

    tests_folder = args.tests_folder or (str(default_tests_dir) if default_tests_dir else None)
    registry_path = args.registry_path or (str(default_registry_file) if default_registry_file else None)

    output_base = Path(args.output_path) if args.output_path else None
    if not output_base:
        if default_output_dir:
            output_base = default_output_dir / DirectoryLocator.DEFAULT_OUTPUT_FILENAME
        else:
            output_base = Path.cwd() / DirectoryLocator.DEFAULT_OUTPUT_FILENAME

    logs_base_dir = default_logs_dir if default_logs_dir else (Path.cwd() / "logs")
    logs_base_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_base_dir / DirectoryLocator.DEFAULT_LOG_FILENAME

    # --- Configuration Setup ---
    config = AppConfig(
        tests_folder=tests_folder,
        registry_path=registry_path,
        output_path=str(output_base),
        log_path=str(log_path),
        logo_path=(
            args.logo_path
            if args.logo_path
            else (str(default_logo_path) if default_logo_path and default_logo_path.exists() else None)
        ),
        test_lab_root=str(default_test_lab_dir) if default_test_lab_dir else None,
        sap_codes=args.sap or [],
        compare_saps=args.compare or [],
        include_noise=not args.no_noise,
        include_comparison=not args.no_comparison,
        open_after_creation=not args.no_open,
        verbose=args.verbose,
        script_version="3.1.0"
    )

    app = MotorReportApp(config, directory_locator=locator)
    app.run()

if __name__ == "__main__":
    main()
