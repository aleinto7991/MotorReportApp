import logging
from typing import List, Dict, Optional, cast
from pathlib import Path

import pandas as pd
from xlsxwriter.workbook import Workbook as XlsxWorkbook

from ..config.app_config import AppConfig
from ..data.models import NoiseTestInfo, MotorTestData
from .builders.excel_formatter import ExcelFormatter
from .builders.sap_sheet_builder import SapSheetBuilder
from .builders.summary_sheet_builder import SummarySheetBuilder
from .builders.comparison_sheet_builder import ComparisonSheetBuilder
from .builders.carichi_sheet_builder import CarichiSheetBuilder
from ..analysis.image_utils import extract_dominant_colors
from ..utils.common import sanitize_sheet_name
from .excel_profiler import ExcelProfiler
from ..services.directory_locator import DirectoryLocator

logger = logging.getLogger(__name__)

class ExcelReport:
    """Main class to orchestrate Excel report generation."""

    def __init__(self, config: AppConfig, noise_handler=None):
        self.config = config
        self.logger = logging.getLogger(__class__.__name__)
        self.output_path = Path(config.output_path) if config.output_path else None
        if not self.output_path:
            raise ValueError("Output path must be set in the configuration.")
        
        resolved_logo_path = self._resolve_logo_path(config.logo_path)
        if resolved_logo_path and config.logo_path != resolved_logo_path:
            config.logo_path = resolved_logo_path
        self.logo_path = resolved_logo_path
        self.noise_handler = noise_handler  # Store the noise handler
        self.writer: Optional[pd.ExcelWriter] = None
        self.workbook: Optional[XlsxWorkbook] = None
        self.formatter: Optional[ExcelFormatter] = None
        self.logo_tab_colors: List[str] = []
        
        # Performance profiling
        self.profiler = ExcelProfiler("Excel Report Generation")
        self.enable_profiling = False  # Can be enabled via config or env var

    def generate(self, grouped_data: Dict[str, List[MotorTestData]], 
                 all_tests_summary: List[MotorTestData], 
                 all_noise_tests_by_sap: Dict[str, List[NoiseTestInfo]],
                 comparison_data: Dict[str, List[MotorTestData]],
                 multiple_comparisons: Optional[List[Dict]] = None,
                 lf_tests_by_sap: Optional[Dict[str, List]] = None) -> bool:
        """
        Generates the full Excel report with all its components.
        Returns True on success, False on failure.
        
        Args:
            grouped_data: SAP-grouped motor test data
            all_tests_summary: Summary of all tests
            all_noise_tests_by_sap: Noise tests grouped by SAP
            comparison_data: Legacy single comparison data
            multiple_comparisons: New multiple comparison groups
            lf_tests_by_sap: Life Test (LF) data grouped by SAP
        """
        if not self.output_path:
            self.logger.error("Cannot generate report without a valid output path.")
            return False

        # Start profiling session
        if self.enable_profiling:
            self.profiler.start_session()

        try:
            with self.profiler.time_operation("excel_writer_open"):
                writer = pd.ExcelWriter(self.output_path, engine='xlsxwriter')
            
            with writer:
                self.writer = writer
                self.workbook = cast(XlsxWorkbook, writer.book)
                
                with self.profiler.time_operation("extract_logo_colors"):
                    if self.logo_path and Path(self.logo_path).exists():
                        self.logo_tab_colors = extract_dominant_colors(str(self.logo_path))
                    else:
                        self.logger.warning("Logo path not found, using default colors.")
                        self.logo_tab_colors = ['#0070C0', '#C00000', '#00B050', '#FFC000']

                if not self.workbook:
                    self.logger.error("Workbook not created. Aborting.")
                    return False
                
                with self.profiler.time_operation("create_formatter"):
                    self.formatter = ExcelFormatter(self.workbook, self.logo_tab_colors)

                sap_sheet_name_map = {sap: sanitize_sheet_name(f"SAP_{sap}") for sap in grouped_data.keys()}

                # Log LF data status
                if lf_tests_by_sap:
                    total_lf = sum(len(tests) for tests in lf_tests_by_sap.values())
                    self.logger.info(f"📊 ExcelWriter: Received {total_lf} LF test(s) for {len(lf_tests_by_sap)} SAP code(s)")
                else:
                    self.logger.info("📊 ExcelWriter: No LF data received (lf_tests_by_sap is empty or None)")
                
                with self.profiler.time_operation("create_summary_sheet"):
                    self._create_summary_sheet(all_tests_summary, sap_sheet_name_map, multiple_comparisons, lf_tests_by_sap)
                
                with self.profiler.time_operation("create_sap_sheets"):
                    self._create_sap_sheets(grouped_data, all_noise_tests_by_sap)
                
                if self.config.include_comparison:
                    # Create legacy single comparison sheet if comparison_data exists
                    if comparison_data:
                        with self.profiler.time_operation("create_comparison_sheet"):
                            self._create_comparison_sheet(comparison_data, sap_sheet_name_map)
                    
                    # Create multiple comparison sheets if multiple_comparisons exist
                    if multiple_comparisons:
                        with self.profiler.time_operation("create_multiple_comparison_sheets"):
                            self._create_multiple_comparison_sheets(multiple_comparisons, grouped_data, sap_sheet_name_map)

            self.logger.info(f"Successfully generated Excel report at {self.output_path}")
            
            # End profiling and print report
            if self.enable_profiling:
                self.profiler.end_session()
                self.profiler.print_report()
            
            return True
        except Exception as e:
            self.logger.error(f"Failed to generate Excel report: {e}", exc_info=True)
            if self.enable_profiling:
                self.profiler.end_session()
            return False

    def _create_summary_sheet(self, all_tests_summary: List[MotorTestData], sap_sheet_name_map: Dict[str, str], multiple_comparisons: Optional[List[Dict]] = None, lf_tests_by_sap: Optional[Dict[str, List]] = None):
        """Creates the main summary sheet."""
        if not self.workbook or not self.formatter: return
        summary_builder = SummarySheetBuilder(
            workbook=self.workbook,
            all_motor_tests=all_tests_summary,
            formatter=self.formatter,
            logo_tab_colors=self.logo_tab_colors,
            sap_sheet_name_map=sap_sheet_name_map,
            multiple_comparisons=multiple_comparisons or [],
            lf_tests_by_sap=lf_tests_by_sap or {}
        )
        summary_builder.build()

    def _create_sap_sheets(self, grouped_data: Dict[str, List[MotorTestData]], all_noise_tests_by_sap: Dict[str, List[NoiseTestInfo]]):
        """Creates a sheet for each SAP code and a consolidated 'CARICHI NOMINALI' sheet."""
        if not self.workbook or not self.formatter:
            return

        # Create a single CarichiSheetBuilder that will accumulate data across SAPs
        carichi_builder = CarichiSheetBuilder(
            workbook=self.workbook,
            formatter=self.formatter,
            logo_tab_colors=self.logo_tab_colors,
            logo_path=self.logo_path,
        )

        for sap_code, tests in grouped_data.items():
            sap_builder = SapSheetBuilder(
                workbook=self.workbook,
                sap_code=sap_code,
                motor_tests=tests,
                formatter=self.formatter,
                config=self.config,
                logo_tab_colors=self.logo_tab_colors,
                all_noise_tests=all_noise_tests_by_sap.get(sap_code, []),
                noise_handler=self.noise_handler  # Pass the noise handler
            )
            sap_builder.build()

            # Add SAP data to the consolidated Carichi builder (it will skip tests without summaries)
            carichi_builder.add_sap_data(sap_code, tests)

        # After iterating all SAPs, build the single consolidated sheet
        carichi_builder.build()

    def _create_comparison_sheet(self, comparison_data: Dict[str, List[MotorTestData]], sap_sheet_name_map: Dict[str, str]):
        """Creates the comparison sheet if enabled."""
        if not self.workbook or not self.formatter: return
        comp_builder = ComparisonSheetBuilder(
            workbook=self.workbook,
            comparison_data=comparison_data,
            formatter=self.formatter,
            logo_colors=self.logo_tab_colors,
            sap_sheet_name_map=sap_sheet_name_map,
            config=self.config
        )
        comp_builder.build()

    def _create_multiple_comparison_sheets(self, multiple_comparisons: List[Dict], grouped_data: Dict[str, List[MotorTestData]], sap_sheet_name_map: Dict[str, str]):
        """Creates multiple comparison sheets based on user-defined comparison groups."""
        if not self.workbook or not self.formatter: 
            return
            
        self.logger.info(f"Creating {len(multiple_comparisons)} comparison sheets")
        
        for i, comparison_group in enumerate(multiple_comparisons):
            group_name = comparison_group.get('name', f'Comparison {i+1}')
            try:
                # Extract comparison group data
                test_labs = comparison_group.get('test_labs', [])
                description = comparison_group.get('description', '')
                
                self.logger.info(f"Creating comparison sheet '{group_name}' with test labs: {test_labs}")
                
                # Build comparison data for this group
                comparison_data = self._build_comparison_data_for_group(test_labs, grouped_data)
                
                if not comparison_data:
                    self.logger.warning(f"No data found for comparison group '{group_name}', skipping")
                    continue
                
                # Create the comparison sheet with a unique name
                sheet_name = f"Comparison_{i+1}"
                comp_builder = ComparisonSheetBuilder(
                    workbook=self.workbook,
                    comparison_data=comparison_data,
                    formatter=self.formatter,
                    logo_colors=self.logo_tab_colors,
                    sap_sheet_name_map=sap_sheet_name_map,
                    config=self.config,
                    custom_sheet_name=sheet_name,
                    custom_title=group_name,
                    custom_description=description
                )
                comp_builder.build()
                
            except Exception as e:
                self.logger.error(f"Error creating comparison sheet for group '{group_name}': {e}")
    
    def _build_comparison_data_for_group(self, test_labs: List[str], grouped_data: Dict[str, List[MotorTestData]]) -> Dict[str, List[MotorTestData]]:
        """Build comparison data for a specific group of test labs."""
        comparison_data = {}
        
        # Search through all SAP data to find the specified test labs
        for sap_code, tests in grouped_data.items():
            matching_tests = []
            for test in tests:
                if test.test_number in test_labs:
                    matching_tests.append(test)
            
            if matching_tests:
                comparison_data[sap_code] = matching_tests
        
        return comparison_data

    def _resolve_logo_path(self, candidate: Optional[str]) -> Optional[str]:
        """Find a usable logo asset path, falling back to auto-detected directories."""
        if candidate:
            candidate_path = Path(candidate)
            if candidate_path.exists():
                return str(candidate_path)

        locator = DirectoryLocator(logger=self.logger)
        fallback = locator.logo_path
        if fallback and fallback.exists():
            return str(fallback)

        # 3. Hardcoded fallback relative to this file
        try:
            # src/reports/excel_report.py -> src/reports -> src -> root -> assets/logo.png
            fallback_path = Path(__file__).resolve().parent.parent.parent / 'assets' / 'logo.png'
            if fallback_path.exists():
                return str(fallback_path)
        except Exception:
            pass

        return None
        # Copy merged cells (adjust coordinates)

