import logging
from typing import List, Dict, Optional
from collections import Counter, defaultdict
from datetime import datetime

from xlsxwriter.workbook import Workbook as XlsxWorkbook
from ...data.models import MotorTestData
from ...utils.common import sanitize_sheet_name
from .excel_formatter import ExcelFormatter
from .excel_sheet_helper import ExcelSheetHelper

logger = logging.getLogger(__name__)

class SummarySheetBuilder:
    """Builds the summary sheet."""
    def __init__(self, workbook: XlsxWorkbook, all_motor_tests: List[MotorTestData], 
                 formatter: ExcelFormatter, logo_tab_colors: Optional[List[str]] = None,
                 sap_sheet_name_map: Optional[Dict[str, str]] = None,
                 multiple_comparisons: Optional[List[Dict]] = None,
                 lf_tests_by_sap: Optional[Dict[str, List]] = None): # Added LF tests
        self.wb = workbook
        self.all_motor_tests = all_motor_tests
        self.fmt = formatter
        self.logger = logging.getLogger(__class__.__name__)
        self.sheet_name = sanitize_sheet_name("Summary_Report") # This will be the first sheet
        self.ws = self.wb.add_worksheet(self.sheet_name)
        self.logo_tab_colors = logo_tab_colors 
        self.sap_sheet_name_map = sap_sheet_name_map if sap_sheet_name_map else {}
        self.multiple_comparisons = multiple_comparisons or []
        self.lf_tests_by_sap = lf_tests_by_sap or {}
        self.current_row = 0

        if self.ws and self.logo_tab_colors:
            try:
                # Use first logo color for Summary sheet tab
                self.ws.set_tab_color(self.logo_tab_colors[0 % len(self.logo_tab_colors)])
            except Exception as e:
                self.logger.warning(f"Could not set tab color for sheet {self.sheet_name}: {e}")
        self.helper = ExcelSheetHelper(self.ws, self.wb, self.fmt)
    
    def build(self):
        # Enhanced header section for Summary
        self.ws.set_row(0, 50) # Increased title row height
        self.ws.set_row(1, 25) # Subtitle row height
        
        # Main title
        self.ws.merge_range(0, 0, 0, 7, "MOTOR TEST SUMMARY REPORT", self.fmt.get('report_title'))
        
        # Add subtitle with test count and generation date
        total_tests = len(self.all_motor_tests)
        unique_saps = len(set(test.sap_code for test in self.all_motor_tests if test.sap_code))
        
        subtitle = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | {total_tests} Tests | {unique_saps} SAP Codes"
        self.ws.merge_range(1, 0, 1, 7, subtitle, self.fmt.get('info_label'))
        
        self.current_row = 3 # Content starts after subtitle

        # Group tests by SAP code
        grouped_by_sap = defaultdict(list)
        for test in self.all_motor_tests:
            grouped_by_sap[test.sap_code or "No SAP Code"].append(test)

        # Sort SAP codes for consistent order
        sorted_sap_codes = sorted(grouped_by_sap.keys())

        # Define headers for the individual tests within each group
        test_headers = ["Test No.", "Date", "Voltage", "Frequency", "Noise", "Status"]
        
        # Set column widths for the new layout
        self.ws.set_column('A:A', 40)  # Merged SAP header
        self.ws.set_column('B:B', 12)  # Test No.
        self.ws.set_column('C:C', 12)  # Date
        self.ws.set_column('D:D', 10)  # Voltage
        self.ws.set_column('E:E', 10)  # Frequency
        self.ws.set_column('F:F', 10)  # Noise
        self.ws.set_column('G:G', 35)  # Status
        # Hide the first column as it's just for the merged range anchor to create an indented look
        self.ws.set_column('A:A', None, None, {'hidden': True})


        for sap_code in sorted_sap_codes:
            tests_for_sap = grouped_by_sap[sap_code]
            test_count = len(tests_for_sap)

            # --- SAP Group Header ---
            sap_header_text = f"SAP Code: {sap_code} ({test_count} test{'s' if test_count > 1 else ''})"
            
            # Merge cells for the group header, from col A to G
            self.ws.merge_range(self.current_row, 0, self.current_row, len(test_headers), sap_header_text, self.fmt.get('header'))
            
            # Add hyperlink to the corresponding SAP sheet if available
            if sap_code in self.sap_sheet_name_map:
                link_target_sheet = self.sap_sheet_name_map[sap_code]
                # The URL is written to the merged range, but the text is taken from the cell content
                self.ws.write_url(self.current_row, 0, f"internal:'{link_target_sheet}'!A1", string=sap_header_text)

            self.current_row += 1

            # --- Test Headers for the Group ---
            # Start headers from column B (index 1) to appear indented
            for col, header in enumerate(test_headers):
                self.ws.write(self.current_row, col + 1, header, self.fmt.get('header'))
            self.current_row += 1

            # --- Individual Test Rows ---
            for test_data in tests_for_sap:
                col_offset = 1  # Start data from column B
                
                # Test Number
                self.ws.write(self.current_row, col_offset, test_data.test_number, self.fmt.get('cell'))
                
                # Date
                self.ws.write(self.current_row, col_offset + 1, test_data.inf_data.date if test_data.inf_data else "N/A", self.fmt.get('cell'))
                
                # Voltage
                voltage_display = "N/A"
                if test_data.inf_data and test_data.inf_data.voltage is not None:
                    try:
                        voltage_display = float(test_data.inf_data.voltage)
                    except (ValueError, TypeError):
                        voltage_display = test_data.inf_data.voltage
                self.ws.write(self.current_row, col_offset + 2, voltage_display, self.fmt.get('cell'))

                # Frequency
                hz_display = "N/A"
                if test_data.inf_data and test_data.inf_data.hz is not None:
                    try:
                        hz_display = float(test_data.inf_data.hz)
                    except (ValueError, TypeError):
                        hz_display = test_data.inf_data.hz
                self.ws.write(self.current_row, col_offset + 3, hz_display, self.fmt.get('cell'))
                
                # Noise Status
                noise_status = "No"
                if test_data.noise_info and test_data.noise_info.image_paths:
                    noise_status = "Yes"
                self.ws.write(self.current_row, col_offset + 4, noise_status, self.fmt.get('cell'))

                # Overall Status
                overall_status = "Data Available"
                if test_data.csv_data is None or test_data.csv_data.empty:
                    overall_status = "Performance Data Missing"
                elif test_data.inf_data is None:
                    overall_status = "INF Data Missing"
                elif not test_data.inf_data.motor_type:
                    overall_status = "Motor Type Missing in INF"
                self.ws.write(self.current_row, col_offset + 5, overall_status, self.fmt.get('cell'))
                
                self.current_row += 1
            
            # Add a blank row for spacing between SAP groups
            self.current_row += 1
        
        # Add Summary Statistics Section
        self._add_multiple_comparisons_summary()
        self._add_noise_analysis_summary()
        self._add_lf_tests_summary()
        self._add_summary_statistics()
        
        self.logger.info(f"Finished building sheet: {self.sheet_name}")

    def _add_multiple_comparisons_summary(self):
        """Add a summary of multiple comparison sheets if any exist."""
        if not self.multiple_comparisons:
            return
        
        # Add spacing
        self.current_row += 2
        
        # Multiple Comparisons Section Title
        self.ws.merge_range(self.current_row, 0, self.current_row, 7, 
                           "COMPARISON SHEETS", self.fmt.get('report_title'))
        self.current_row += 2
        
        # Add header row for comparison details
        headers = ["Comparison Name", "Test Labs", "Description", "Go to Sheet"]
        for i, header in enumerate(headers):
            self.ws.write(self.current_row, i, header, self.fmt.get('info_label'))
        self.current_row += 1
        
        # Add each comparison group
        for i, comparison_group in enumerate(self.multiple_comparisons):
            name = comparison_group.get('name', f'Comparison {i+1}')
            test_labs = comparison_group.get('test_labs', [])
            description = comparison_group.get('description', 'No description')
            sheet_name = f"Comparison_{i+1}"
            
            # Format test labs as comma-separated string
            test_labs_str = ', '.join(test_labs) if test_labs else 'No test labs'
            
            # Write comparison details
            self.ws.write(self.current_row, 0, name, self.fmt.get('cell'))
            self.ws.write(self.current_row, 1, test_labs_str, self.fmt.get('cell'))
            self.ws.write(self.current_row, 2, description, self.fmt.get('cell'))
            
            # Add hyperlink to the comparison sheet in the "Go to Sheet" column
            self.ws.write_url(self.current_row, 3, f"internal:'{sheet_name}'!A1", string=f"→ {sheet_name}", cell_format=self.fmt.get('cell_link'))
            
            self.current_row += 1
        
        # Add summary line
        self.current_row += 1
        summary_text = f"Total: {len(self.multiple_comparisons)} comparison sheet{'s' if len(self.multiple_comparisons) != 1 else ''} created"
        self.ws.merge_range(self.current_row, 0, self.current_row, 3, 
                           summary_text, self.fmt.get('info_label'))
        self.current_row += 1

    def _add_noise_analysis_summary(self):
        """Add a summary of noise analysis with links to SAP sheets containing noise data."""
        # Find SAP codes that have noise data
        sap_codes_with_noise = set()
        noise_test_counts = {}
        
        for test in self.all_motor_tests:
            if test.noise_info and test.noise_info.image_paths:
                sap_code = test.sap_code
                if sap_code:
                    sap_codes_with_noise.add(sap_code)
                    noise_test_counts[sap_code] = noise_test_counts.get(sap_code, 0) + 1
        
        if not sap_codes_with_noise:
            return  # No noise data to show
        
        # Add spacing before section
        self.current_row += 2
        
        # Noise Analysis Section Title
        self.ws.merge_range(self.current_row, 0, self.current_row, 3, 
                           "NOISE ANALYSIS SHEETS", self.fmt.get('report_title'))
        self.current_row += 1
        
        # Add header row for noise analysis details
        headers = ["SAP Code", "Tests with Noise", "Sheet Name", "Go to Sheet"]
        for i, header in enumerate(headers):
            self.ws.write(self.current_row, i, header, self.fmt.get('info_label'))
        self.current_row += 1
        
        # Add each SAP with noise data
        for sap_code in sorted(sap_codes_with_noise):
            test_count = noise_test_counts.get(sap_code, 0)
            sheet_name = self.sap_sheet_name_map.get(sap_code, f"SAP_{sap_code}")
            
            # Write noise analysis details
            self.ws.write(self.current_row, 0, sap_code, self.fmt.get('cell'))
            self.ws.write(self.current_row, 1, f"{test_count} test{'s' if test_count != 1 else ''}", self.fmt.get('cell'))
            self.ws.write(self.current_row, 2, sheet_name, self.fmt.get('cell'))
            
            # Add hyperlink to the SAP sheet in the "Go to Sheet" column
            if sap_code in self.sap_sheet_name_map:
                self.ws.write_url(self.current_row, 3, f"internal:'{sheet_name}'!A1", string="→ View Noise Analysis", cell_format=self.fmt.get('cell_link'))
            else:
                self.ws.write(self.current_row, 3, "Sheet not found", self.fmt.get('cell'))
            
            self.current_row += 1
        
        # Add summary line
        self.current_row += 1
        total_noise_tests = sum(noise_test_counts.values())
        summary_text = f"Total: {total_noise_tests} test{'s' if total_noise_tests != 1 else ''} with noise data across {len(sap_codes_with_noise)} SAP code{'s' if len(sap_codes_with_noise) != 1 else ''}"
        self.ws.merge_range(self.current_row, 0, self.current_row, 3, 
                           summary_text, self.fmt.get('info_label'))
        self.current_row += 1

    def _add_lf_tests_summary(self):
        """Add a summary of Life Test (LF) data with hyperlinks to test files."""
        self.logger.info("=" * 60)
        self.logger.info("🔬 BUILDING LF TESTS SUMMARY SECTION")
        self.logger.info("=" * 60)
        self.logger.info(f"📊 lf_tests_by_sap type: {type(self.lf_tests_by_sap)}")
        self.logger.info(f"📊 lf_tests_by_sap value: {self.lf_tests_by_sap}")
        
        if not self.lf_tests_by_sap:
            self.logger.info("ℹ️ No LF data to display (lf_tests_by_sap is empty or None)")
            self.logger.info("   This is normal if no LF tests were selected")
            self.logger.warning("⚠️ IF YOU SELECTED LF TESTS IN THE GUI, THIS IS A BUG!")
            return  # No LF data to show
        
        # Add spacing before section
        self.current_row += 2
        
        self.logger.info(f"📋 LF tests by SAP: {list(self.lf_tests_by_sap.keys())}")
        
        # LF Tests Section Title
        self.ws.merge_range(self.current_row, 0, self.current_row, 4, 
                           "🔬 LIFE TEST (LF) DATA", self.fmt.get('report_title'))
        self.logger.info(f"✅ Added LF section title at row {self.current_row}")
        self.current_row += 1
        
        # Add header row for LF test details
        headers = ["SAP Code", "Test Number", "Notes", "File Path", "Open File"]
        for i, header in enumerate(headers):
            self.ws.write(self.current_row, i, header, self.fmt.get('info_label'))
        self.current_row += 1
        
        # Track totals
        total_lf_tests = 0
        sap_count = 0
        
        # Add each SAP with LF tests
        for sap_code in sorted(self.lf_tests_by_sap.keys()):
            lf_tests = self.lf_tests_by_sap[sap_code]
            if not lf_tests:
                continue
            
            sap_count += 1
            
            for lf_test in lf_tests:
                total_lf_tests += 1
                
                # Write LF test details
                self.ws.write(self.current_row, 0, sap_code, self.fmt.get('cell'))
                self.ws.write(self.current_row, 1, lf_test.test_number, self.fmt.get('cell'))
                
                self.logger.debug(f"  📝 Writing LF test {lf_test.test_number} at row {self.current_row}")
                
                # Notes (truncate if too long)
                notes = lf_test.notes or "No notes"
                if len(notes) > 80:
                    notes = notes[:77] + "..."
                self.ws.write(self.current_row, 2, notes, self.fmt.get('cell'))
                
                # File path
                if lf_test.file_exists and lf_test.file_path:
                    file_path_str = str(lf_test.file_path)
                    self.ws.write(self.current_row, 3, file_path_str, self.fmt.get('cell'))
                    
                    # Add hyperlink to open the file
                    try:
                        # Convert Windows path to file URL format
                        # Replace backslashes with forward slashes
                        file_url = file_path_str.replace('\\', '/')
                        
                        # Add file:/// protocol (three slashes for local files)
                        # Don't add extra slashes if the path already starts with a drive letter
                        if not file_url.startswith('file:'):
                            file_url = f"file:///{file_url}"
                        
                        self.logger.info(f"  🔗 Creating hyperlink for {lf_test.test_number}")
                        self.logger.info(f"     Original path: {file_path_str}")
                        self.logger.info(f"     Hyperlink URL: {file_url}")
                        
                        self.ws.write_url(self.current_row, 4, file_url, 
                                         string="📂 Open File", 
                                         cell_format=self.fmt.get('cell_link'))
                        self.logger.info(f"  ✅ Successfully added hyperlink")
                    except Exception as e:
                        self.logger.error(f"  ❌ Could not create hyperlink for {lf_test.test_number}: {e}")
                        import traceback
                        self.logger.error(traceback.format_exc())
                        self.ws.write(self.current_row, 4, "Link error", self.fmt.get('cell'))
                else:
                    self.ws.write(self.current_row, 3, "File not found", self.fmt.get('red_highlight'))
                    self.ws.write(self.current_row, 4, "❌ Not available", self.fmt.get('red_highlight'))
                    self.logger.warning(f"  ❌ File not found for test {lf_test.test_number}")
                
                self.current_row += 1
        
        # Add summary line
        if total_lf_tests > 0:
            self.current_row += 1
            summary_text = f"Total: {total_lf_tests} Life Test{'s' if total_lf_tests != 1 else ''} across {sap_count} SAP code{'s' if sap_count != 1 else ''}"
            self.ws.merge_range(self.current_row, 0, self.current_row, 4, 
                               summary_text, self.fmt.get('info_label'))
            self.current_row += 1
            self.logger.info(f"✅ Added LF tests summary: {total_lf_tests} tests for {sap_count} SAP codes")

    def _add_summary_statistics(self):
        """Add a comprehensive summary statistics section."""
        # Add spacing
        self.current_row += 2
        
        # Statistics Section Title
        self.ws.merge_range(self.current_row, 0, self.current_row, 7, 
                           "SUMMARY STATISTICS", self.fmt.get('report_title'))
        self.current_row += 2
        
        # Calculate statistics
        total_tests = len(self.all_motor_tests)
        unique_saps = len(set(test.sap_code for test in self.all_motor_tests if test.sap_code))
        
        # SAP code statistics
        sap_test_counts = {}
        for test_data in self.all_motor_tests:
            sap_code = test_data.sap_code or "No SAP Code"
            if sap_code not in sap_test_counts:
                sap_test_counts[sap_code] = []
            sap_test_counts[sap_code].append(test_data.test_number)
        
        multi_test_saps = sum(1 for tests in sap_test_counts.values() if len(tests) > 1)
        single_test_saps = len(sap_test_counts) - multi_test_saps
        
        # Data availability statistics
        tests_with_performance = sum(1 for test in self.all_motor_tests 
                                   if test.csv_data is not None and not test.csv_data.empty)
        tests_with_inf = sum(1 for test in self.all_motor_tests if test.inf_data is not None)
        tests_with_noise = sum(1 for test in self.all_motor_tests 
                             if test.noise_info and test.noise_info.image_paths)
        
        # Voltage/Frequency analysis
        voltages = []
        frequencies = []
        for test in self.all_motor_tests:
            if test.inf_data:
                if test.inf_data.voltage:
                    try:
                        voltages.append(float(test.inf_data.voltage))
                    except (ValueError, TypeError):
                        pass
                if test.inf_data.hz:
                    try:
                        frequencies.append(float(test.inf_data.hz))
                    except (ValueError, TypeError):
                        pass
        
        unique_voltages = len(set(voltages)) if voltages else 0
        unique_frequencies = len(set(frequencies)) if frequencies else 0
        
        # Create statistics table
        stats_data = [
            ["Total Tests", total_tests],
            ["Unique SAP Codes", unique_saps],
            ["SAP Codes with Multiple Tests", multi_test_saps],
            ["SAP Codes with Single Test", single_test_saps],
            ["", ""],  # Spacer
            ["Tests with Performance Data", f"{tests_with_performance}/{total_tests} ({tests_with_performance/total_tests*100:.1f}%)"],
            ["Tests with INF Data", f"{tests_with_inf}/{total_tests} ({tests_with_inf/total_tests*100:.1f}%)"],
            ["Tests with Noise Images", f"{tests_with_noise}/{total_tests} ({tests_with_noise/total_tests*100:.1f}%)"],
            ["", ""],  # Spacer
            ["Unique Voltage Levels", unique_voltages],
            ["Unique Frequency Values", unique_frequencies],
        ]
        
        # Write statistics table
        for i, (label, value) in enumerate(stats_data):
            if label == "":  # Spacer row
                self.current_row += 1
                continue
                
            # Label column
            self.ws.write(self.current_row, 0, label, self.fmt.get('info_label'))
            # Value column
            self.ws.write(self.current_row, 1, value, self.fmt.get('cell'))
            self.current_row += 1
        
        # Add voltage/frequency breakdown if available
        if voltages or frequencies:
            self.current_row += 1
            self.ws.merge_range(self.current_row, 0, self.current_row, 7, 
                               "TEST PARAMETERS BREAKDOWN", self.fmt.get('header'))
            self.current_row += 1
            
            if voltages:
                voltage_counts = Counter(voltages)
                self.ws.write(self.current_row, 0, "Voltage Distribution:", self.fmt.get('info_label'))
                self.current_row += 1
                for voltage, count in sorted(voltage_counts.items()):
                    self.ws.write(self.current_row, 1, f"{voltage}V", self.fmt.get('cell'))
                    self.ws.write(self.current_row, 2, f"{count} tests", self.fmt.get('cell'))
                    self.current_row += 1
            
            if frequencies:
                self.current_row += 1
                freq_counts = Counter(frequencies)
                self.ws.write(self.current_row, 0, "Frequency Distribution:", self.fmt.get('info_label'))
                self.current_row += 1
                for freq, count in sorted(freq_counts.items()):
                    self.ws.write(self.current_row, 1, f"{freq}Hz", self.fmt.get('cell'))
                    self.ws.write(self.current_row, 2, f"{count} tests", self.fmt.get('cell'))
                    self.current_row += 1
