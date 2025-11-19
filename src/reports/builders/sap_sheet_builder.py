import logging
from typing import List, Dict, Any, Optional, Tuple, Union, Set
from collections import defaultdict, Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np
from xlsxwriter.workbook import Workbook as XlsxWorkbook
from xlsxwriter.worksheet import Worksheet as XlsxWorksheet
from xlsxwriter.chart import Chart as XlsxChart
from xlsxwriter.format import Format as XlsxFormat
from xlsxwriter.exceptions import InvalidWorksheetName, FileCreateError, XlsxWriterException
from xlsxwriter.utility import xl_col_to_name

import xlsxwriter # Added for version access

from PIL import Image

from ...config.app_config import AppConfig
from ...data.models import (
    InfData,
    NoiseTestInfo,
    MotorTestData,
    TestLabSummary,
    SchedaSummary,
    CollaudoSummary,
)
from ...config.measurement_units import apply_unit_preferences, UNIT_CONFIG
from ...utils.common import sanitize_sheet_name, EXCEL_MAX_SHEET_NAME_LENGTH
from .excel_formatter import ExcelFormatter
from .excel_sheet_helper import ExcelSheetHelper

logger = logging.getLogger(__name__)

class SapSheetBuilder:
    """Builds a single SAP-specific sheet in the Excel report."""
    def __init__(self, workbook: XlsxWorkbook, sap_code: str, motor_tests: List[MotorTestData],
                 formatter: ExcelFormatter, config: AppConfig, logo_tab_colors: Optional[List[str]] = None,
                 all_noise_tests: Optional[List[NoiseTestInfo]] = None, noise_handler=None): # Added noise_handler
        self.wb = workbook
        self.sap_code = sap_code
        self.all_motor_tests = list(motor_tests)
        self.test_lab_summary_map = {
            mt.test_number: mt.test_lab_summary
            for mt in self.all_motor_tests
            if mt.test_lab_summary is not None
        }

        self.performance_tests = self._select_performance_tests()
        if not self.performance_tests:
            self.logger.warning(
                "No primary performance datasets detected for SAP %s; falling back to all CSV-bearing tests.",
                self.sap_code,
            )
            self.performance_tests = [
                mt
                for mt in self.all_motor_tests
                if mt.csv_data is not None and not mt.csv_data.empty
            ]
        self.motor_tests = self.performance_tests
        self.fmt = formatter
        self.config = config
        self.logger = logging.getLogger(f"{__class__.__name__}[{sap_code}]")
        self.logo_tab_colors = logo_tab_colors # Store logo colors
        self.all_noise_tests = all_noise_tests or [] # Store all noise tests for this SAP
        self.noise_handler = noise_handler  # Store provided noise handler

        self.sheet_name = sanitize_sheet_name(f"SAP_{self.sap_code}")
        try:
            self.ws = self.wb.add_worksheet(self.sheet_name)
        except InvalidWorksheetName:
            self.logger.warning(f"Invalid sheet name '{self.sheet_name}', trying fallback.")
            self.sheet_name = sanitize_sheet_name(f"SAP_{self.sap_code[:20]}_FB")
            self.ws = self.wb.add_worksheet(self.sheet_name)

        if self.ws and self.logo_tab_colors:
            try:
                # Use first logo color for SAP sheets
                self.ws.set_tab_color(self.logo_tab_colors[0 % len(self.logo_tab_colors)])
            except Exception as e:
                self.logger.warning(f"Could not set tab color for sheet {self.sheet_name}: {e}")

        self.helper = ExcelSheetHelper(self.ws, self.wb, self.fmt)
        self.current_row = 0
        self.data_start_col = 0  # Start all content sections from Column A (index 0)
        self.unit_metadata: Optional[Dict[str, Dict[str, str]]] = None

    def _select_performance_tests(self) -> List[MotorTestData]:
        candidates = [
            mt
            for mt in self.all_motor_tests
            if mt.csv_data is not None and not mt.csv_data.empty
        ]
        if not candidates:
            return []

        base_names = {mt.test_number.lower() for mt in candidates}
        primaries: List[MotorTestData] = []
        for mt in candidates:
            if self._is_alias_test(mt, base_names):
                continue
            primaries.append(mt)

        if primaries:
            return primaries
        return candidates

    def _is_alias_test(self, mt: MotorTestData, known_tests: Set[str]) -> bool:
        status_lower = (mt.status_message or "").lower()
        if "csv data reused from" in status_lower:
            return True

        csv_path = getattr(mt, "csv_path", None)
        if csv_path:
            try:
                csv_stem = Path(csv_path).stem.lower()
                if csv_stem != mt.test_number.lower() and csv_stem in known_tests:
                    return True
            except Exception:
                pass
        return False

    def build(self):
        self._add_header_and_logo()

        if not self.performance_tests:
            self.logger.warning(
                f"No valid motor test data with CSVs for SAP {self.sap_code}. Sheet will only show supplemental sections."
            )
            self.ws.write(
                self.current_row,
                self.data_start_col,
                f"No performance data found for SAP {self.sap_code}",
                self.fmt.get('info_label'),
            )
            self.current_row += 2
        else:
            self._add_motor_info_section()

            all_dfs_for_sap = self._prepare_performance_data()
            if not all_dfs_for_sap:
                self.ws.write(
                    self.current_row,
                    self.data_start_col,
                    "No CSV data available for this SAP code.",
                    self.fmt.get('info_label'),
                )
                self.current_row += 2
            else:
                representative_df = pd.concat(all_dfs_for_sap, ignore_index=True)
                data_table_start_row = self._add_data_table_with_dividers(all_dfs_for_sap)
                if data_table_start_row > -1:
                    self._add_charts_section(all_dfs_for_sap, data_table_start_row, len(representative_df))

        # Summaries from external test-lab workbooks precede the noise section
        self._add_test_lab_summary_section()

        # Add noise section RIGHT AFTER performance charts for better organization
        self._add_noise_section()
        
        self.logger.info(f"Finished building sheet: {self.sheet_name}")

        # --- Restore print settings: fit to 1 page wide by 1 page tall ---
        try:
            self.ws.set_paper(9)  # A4
            self.ws.set_landscape()
            self.ws.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
            self.ws.fit_to_pages(1, 1)
            self.ws.center_horizontally()
            self.ws.center_vertically()
            max_col = 15  # 0-based, so column 16 = P
            max_row = self.current_row if self.current_row > 0 else 60
            self.ws.print_area(0, 0, max_row, max_col)  # A1 to P{max_row}
        except Exception as e:
            self.logger.warning(f"Could not set print settings for sheet {self.sheet_name}: {e}")

    def _prepare_performance_data(self) -> List[pd.DataFrame]:
        """Convert motor test dataframes according to the configured unit preferences."""
        converted_dfs: List[pd.DataFrame] = []
        self.unit_metadata = None

        for mt in self.motor_tests:
            if mt.csv_data is None or mt.csv_data.empty:
                continue

            converted_df, metadata = apply_unit_preferences(mt.csv_data, self.config)
            converted_dfs.append(converted_df)

            if self.unit_metadata is None:
                self.unit_metadata = metadata

        return converted_dfs

    def _add_header_and_logo(self):
        # Set row heights for header area
        self.ws.set_row(0, 45) # 45 points
        self.ws.set_row(1, 45) # 45 points

        a4_cols = 16 # Columns A to P
        
        # --- New pixel-based column width for precision ---
        # Let's define a standard column width in pixels. 
        # A width of 110px is a good starting point, similar to the previous default.
        default_col_width_pixels = 110 
        for col in range(a4_cols): # A-P
            self.ws.set_column_pixels(col, col, default_col_width_pixels)

        # Merge for logo area A1:B2 - do this once regardless of logo presence
        self.ws.merge_range(0, 0, 1, 1, '', None)  # Logo A1:B2

        if self.config.logo_path:
            logo_path_obj = Path(self.config.logo_path)
            if logo_path_obj.exists():
                try:
                    # --- Direct Pixel-Based Centering ---
                    # 1. Get merged cell dimensions in pixels.
                    # Row heights are in points. 1 point = 1/72 inch. Excel uses 96 DPI.
                    # So, pixel_height = (points / 72) * 96.
                    merged_cell_height_px = ((45 + 45) / 72) * 96
                    merged_cell_width_px = 2 * default_col_width_pixels # Since we set it directly

                    with Image.open(logo_path_obj) as img:
                        img_width, img_height = img.size

                    # 2. Calculate scale to fit the image within the pixel dimensions.
                    if img_width == 0 or img_height == 0:
                        scale = 1.0
                    else:
                        x_scale = merged_cell_width_px / img_width
                        y_scale = merged_cell_height_px / img_height
                        scale = min(x_scale, y_scale, 1.0) # Fit and don't upscale

                    # 3. Calculate the display size of the logo in pixels.
                    logo_display_width_px = img_width * scale
                    logo_display_height_px = img_height * scale

                    # 4. Calculate the offsets in pixels to center the logo.
                    x_offset = (merged_cell_width_px - logo_display_width_px) / 2
                    y_offset = (merged_cell_height_px - logo_display_height_px) / 2

                    self.ws.insert_image('A1', self.config.logo_path, {
                        'x_offset': x_offset, 'y_offset': y_offset,
                        'x_scale': scale, 'y_scale': scale,
                        'object_position': 1
                    })
                except Exception as e:
                    self.logger.error(f"Could not insert or scale logo {self.config.logo_path}: {e}")
                    self.ws.write(0, 0, "Logo Error", self.fmt.get('red_highlight'))
            else:
                self.logger.warning(f"Logo file not found at: {self.config.logo_path}")
                self.ws.write(0, 0, "Logo Not Found", self.fmt.get('info_label'))
        else:
            # No logo path provided
            self.logger.warning("No logo path configured.")
            self.ws.write(0, 0, "Logo Here", self.fmt.get('info_label'))

        # Merge for title: C1 to P2 (columns 2 to 15)
        self.ws.merge_range(0, 2, 1, a4_cols-1, f"MOTOR PERFORMANCE REPORT - SAP: {self.sap_code}", self.fmt.get('report_title'))

        self.current_row = 2  # Content starts on row index 2 (Excel row 3)

    def _add_motor_info_section(self):
        start_col = self.data_start_col # Will be 0 (Column A)

        # Enhanced header showing test count
        test_count = len(self.motor_tests)
        if test_count > 1:
            header_text = f"Motor Information ({test_count} Tests for SAP {self.sap_code}):"
            header_format = self.fmt.get('info_label')  # Highlight multiple tests
        else:
            header_text = f"Motor Information (SAP {self.sap_code}):"
            header_format = self.fmt.get('header')

        self.ws.write(self.current_row, start_col, header_text, header_format)
        self.ws.set_row(self.current_row, 20)
        self.current_row +=1

        col_headers = ["Test No.", "Date", "Voltage (V)", "Frequency (Hz)", "Comments"]
        self.ws.set_column(start_col + 0, start_col + 0, 12)
        self.ws.set_column(start_col + 1, start_col + 1, 12)
        self.ws.set_column(start_col + 2, start_col + 2, 12)
        self.ws.set_column(start_col + 3, start_col + 3, 12)
        self.ws.set_column(start_col + 4, start_col + 4, 60)

        for i, header in enumerate(col_headers):
             self.ws.write(self.current_row, start_col + i, header, self.fmt.get('motor_info_header'))

        start_info_row = self.current_row + 1

        for mt_idx, mt in enumerate(self.motor_tests):
            row_to_write = start_info_row + mt_idx
            self.ws.write(row_to_write, start_col + 0, mt.test_number, self.fmt.get('motor_info_value'))
            self.ws.write(row_to_write, start_col + 1, mt.inf_data.date, self.fmt.get('motor_info_value'))

            voltage_val = mt.inf_data.voltage
            try: # Attempt to convert to float for numeric storage
                voltage_val = float(mt.inf_data.voltage)
            except (ValueError, TypeError):
                pass # Keep as string if conversion fails
            self.ws.write(row_to_write, start_col + 2, voltage_val, self.fmt.get('motor_info_value'))

            hz_val = mt.inf_data.hz
            try: # Attempt to convert to float
                hz_val = float(mt.inf_data.hz)
            except (ValueError, TypeError):
                pass # Keep as string if conversion fails
            self.ws.write(row_to_write, start_col + 3, hz_val, self.fmt.get('motor_info_value'))

            comment_cell_format = self.fmt.get('text_left_border')
            self.ws.write(row_to_write, start_col + 4, mt.inf_data.comment, comment_cell_format)
            chars_per_line_approx = 50
            num_lines = max(1, (len(mt.inf_data.comment or "") // chars_per_line_approx + 1))
            self.ws.set_row(row_to_write, max(15, num_lines * 15) )

        self.current_row = start_info_row + len(self.motor_tests) + 1 # Leave one blank row after section

    def _add_data_table_with_dividers(self, list_of_dfs: List[pd.DataFrame]) -> int:
        """Add performance data table with visual dividers between different tests."""
        col_offset = self.data_start_col # Will be 0 (Column A)

        self.ws.write(self.current_row, col_offset, "Performance Data:", self.fmt.get('header'))
        self.ws.set_row(self.current_row, 20) # Height for the section title
        self.current_row += 1 # Move to next row for table headers/data

        if not list_of_dfs or all(df.empty for df in list_of_dfs):
            self.ws.write(self.current_row, col_offset, "No performance data available.", self.fmt.get('info_label'))
            self.current_row += 1
            return -1
        # Use the first non-empty DataFrame for column headers
        representative_df = next((df for df in list_of_dfs if not df.empty), None)
        if representative_df is None:
            self.ws.write(self.current_row, col_offset, "No performance data available.", self.fmt.get('info_label'))
            self.current_row += 1
            return -1

        # Write column headers
        for c_idx, col_name in enumerate(representative_df.columns):
            if col_offset + c_idx < 16: # Max 16 columns (0-15)
                self.ws.write(self.current_row, col_offset + c_idx, col_name, self.fmt.get('header'))

        data_values_start_row = self.current_row + 1
        current_data_row = data_values_start_row

        # Process each test's data separately with dividers
        for test_idx, df in enumerate(list_of_dfs):
            if df.empty:
                continue

            # Add test identifier if multiple tests
            if len(list_of_dfs) > 1:
                test_label = f"Test {self.motor_tests[test_idx].test_number}"
                # Add test label in first column, merge across visible columns
                visible_cols = min(len(df.columns), 16 - col_offset)
                if visible_cols > 1:
                    self.ws.merge_range(current_data_row, col_offset, current_data_row, col_offset + visible_cols - 1,
                                       test_label, self.fmt.get('info_label'))
                else:
                    self.ws.write(current_data_row, col_offset, test_label, self.fmt.get('info_label'))
                current_data_row += 1
              # Write data rows for this test
            test_start_row = current_data_row
            for r_idx, row in enumerate(df.itertuples(index=False)):
                for c_idx, value in enumerate(row):
                    if col_offset + c_idx < 16: # Max 16 columns
                        fmt_name = 'decimal_2'
                        if isinstance(value, str):
                            fmt_name = 'cell'
                        elif isinstance(value, int):
                            fmt_name = 'integer'
                        col_name_lower = str(df.columns[c_idx]).lower()
                        write_value = value
                        if 'efficiency' in col_name_lower or '%' in col_name_lower:
                            # Check if value is already in percentage format (0-100 range) vs decimal format (0-1 range)
                            if isinstance(value, (int, float)) and not pd.isna(value):
                                if value > 1.0:
                                    # Value is in percentage format (e.g., 85.5), convert to decimal for Excel percent format
                                    write_value = value / 100.0  # Convert 85.5 to 0.855
                                    fmt_name = 'percent'
                                else:
                                    # Value is already in decimal format (e.g., 0.855)
                                    fmt_name = 'percent'
                            else:
                                fmt_name = 'decimal_2'

                        # Use divider format for last row of test if there are more tests
                        if test_idx < len(list_of_dfs) - 1 and r_idx == len(df) - 1:
                            if fmt_name == 'decimal_2':
                                fmt_name = 'test_divider_decimal'
                            elif fmt_name == 'integer':
                                fmt_name = 'test_divider_integer'
                            elif fmt_name == 'percent':
                                fmt_name = 'test_divider_percent'
                            else:  # 'cell' or string
                                fmt_name = 'test_divider_cell'

                        # Safely handle the value to prevent Excel corruption
                        safe_value = write_value
                        if isinstance(value, (int, float)):
                            if pd.isna(value) or not np.isfinite(value):
                                safe_value = "N/A"
                                fmt_name = 'cell'  # Use text format for N/A
                        elif value is None:
                            safe_value = "N/A"
                            fmt_name = 'cell'  # Use text format for N/A

                        self.ws.write(current_data_row, col_offset + c_idx, safe_value, self.fmt.get(fmt_name))
                current_data_row += 1
              # Add conditional formatting for this test's data
            if not df.empty and len(df.columns) > 0 and col_offset + len(df.columns) - 1 < 16:
                self.helper.apply_conditional_formatting(df, test_start_row, start_col=col_offset)

        self.current_row = current_data_row + 1 # Leave one blank row after section
        return data_values_start_row

    def _add_data_table(self, combined_df: pd.DataFrame) -> int:
        # self.current_row += 1 # Removed for compactness, title will use blank row from previous section
        col_offset = self.data_start_col # Will be 0 (Column A)
        # Removed: if col_offset < 2: col_offset = 2

        self.ws.write(self.current_row, col_offset, "Performance Data:", self.fmt.get('header'))
        self.ws.set_row(self.current_row, 20) # Height for the section title
        self.current_row += 1 # Move to next row for table headers/data

        if combined_df.empty:
            self.ws.write(self.current_row, col_offset, "No performance data available.", self.fmt.get('info_label'))
            self.current_row += 1
            return -1

        for c_idx, col_name in enumerate(combined_df.columns):
            # Ensure table does not exceed print area (P / 15th column)
            if col_offset + c_idx < 16: # Max 16 columns (0-15)
                self.ws.write(self.current_row, col_offset + c_idx, col_name, self.fmt.get('header'))

        data_values_start_row = self.current_row + 1

        # DEBUG: Log data table position for comparison charts
        data_end_row = data_values_start_row + len(combined_df) - 1
        self.logger.info(f"[{self.sap_code}] Data table: rows {data_values_start_row}-{data_end_row}, cols {col_offset}-{col_offset + len(combined_df.columns) - 1}")
        self.logger.info(f"[{self.sap_code}] Columns: {list(combined_df.columns)}")

        for r_idx, row in enumerate(combined_df.itertuples(index=False)):
            for c_idx, value in enumerate(row):
                if col_offset + c_idx < 16: # Max 16 columns
                    fmt_name = 'decimal_2'
                    if isinstance(value, str): fmt_name = 'cell'
                    elif isinstance(value, int): fmt_name = 'integer'
                    col_name_lower = str(combined_df.columns[c_idx]).lower()
                    write_value = value
                    if 'efficiency' in col_name_lower or '%' in col_name_lower:
                        # Check if value is already in percentage format (0-100 range) vs decimal format (0-1 range)
                        if isinstance(value, (int, float)) and not pd.isna(value):
                            if value > 1.0:
                                # Value is in percentage format (e.g., 85.5), convert to decimal for Excel percent format
                                write_value = value / 100.0  # Convert 85.5 to 0.855
                                fmt_name = 'percent'
                            else:
                                # Value is already in decimal format (e.g., 0.855)
                                fmt_name = 'percent'
                        else:
                            fmt_name = 'decimal_2'
                    # Safely handle the value to prevent Excel corruption
                    safe_value = write_value
                    if isinstance(value, (int, float)):
                        if pd.isna(value) or not np.isfinite(value):
                            safe_value = "N/A"
                            fmt_name = 'cell'  # Use text format for N/A
                    elif value is None:
                        safe_value = "N/A"
                        fmt_name = 'cell'  # Use text format for N/A
                        
                    self.ws.write(data_values_start_row + r_idx, col_offset + c_idx, safe_value, self.fmt.get(fmt_name))

        # Removed: self.helper.auto_fit_columns(combined_df, start_col=col_offset, max_width=25)
        # Columns will use default width (15) set in _add_header_and_logo
        if not combined_df.empty and len(combined_df.columns) > 0 and col_offset + len(combined_df.columns) -1 < 16:
             self.helper.apply_conditional_formatting(combined_df, data_values_start_row, start_col=col_offset)

        self.current_row = data_values_start_row + len(combined_df) + 1 # Leave one blank row after section
        return data_values_start_row

    def _add_charts_section(self, list_of_dfs: List[pd.DataFrame], data_table_start_row: int, total_data_rows: int):
        # self.current_row +=1 # Removed for compactness
        col_offset = self.data_start_col # Will be 0 (Column A)

        self.ws.write(self.current_row, col_offset, "Performance Charts:", self.fmt.get('header'))
        self.ws.set_row(self.current_row, 20) # Height for the section title
        charts_block_start_row = self.current_row + 1

        flow_info = self._get_unit_info("flow")
        pressure_info = self._get_unit_info("pressure")
        power_info = self._get_unit_info("power")
        speed_info = self._get_unit_info("speed")

        chart_configs = [
            {
                'title': f"{pressure_info['chart_name']} vs {flow_info['chart_name']}",
                'x_col': flow_info['column'],
                'y_col': pressure_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': pressure_info['axis_label'],
            },
            {
                'title': f"Efficiency (%) vs {flow_info['chart_name']}",
                'x_col': flow_info['column'],
                'y_col': 'Efficiency (%)',
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': 'Efficiency (%)',
            },
            {
                'title': f"{power_info['chart_name']} vs {flow_info['chart_name']}",
                'x_col': flow_info['column'],
                'y_col': power_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': power_info['axis_label'],
            },
            {
                'title': f"{speed_info['chart_name']} vs {flow_info['chart_name']}",
                'x_col': flow_info['column'],
                'y_col': speed_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': speed_info['axis_label'],
            },
        ]

        # Chart layout: 2 charts per row. Each chart needs ~6 columns of width 15 (90 units / 15 units/col)
        # Chart 1: Col A (0). Chart 2: Col G (6).
        # Max columns A-P (0-15).
        chart_pixel_width = 680
        chart_pixel_height = 400
        # Approx 6 columns for a chart of 680px width if col width is 15 (680 / (15 * 7.5) ~ 6)
        cols_per_chart = 6

        chart_positions = [
            (charts_block_start_row, col_offset + 0*cols_per_chart),             # Chart 1 in Col A
            (charts_block_start_row, col_offset + 1*cols_per_chart),             # Chart 2 in Col G
            (charts_block_start_row + 22, col_offset + 0*cols_per_chart),        # Chart 3 in Col A (next row)
            (charts_block_start_row + 22, col_offset + 1*cols_per_chart)         # Chart 4 in Col G (next row)
        ]

        max_chart_row_span = 0

        if not list_of_dfs:
            self.logger.warning(f"No DataFrames available for charting in SAP {self.sap_code}.")
            self.ws.write(charts_block_start_row, col_offset, "No chart data.", self.fmt.get('info_label'))
            self.current_row = charts_block_start_row + 1
            return

        for chart_idx, cfg in enumerate(chart_configs):
            chart = self.helper.create_chart(cfg['title'], cfg['x_axis_label'], cfg['y_axis_label'])
            if not chart:
                self.logger.warning(f"Skipping chart '{cfg['title']}' due to creation failure.")
                continue
            
            current_data_row = data_table_start_row
            has_series = False
            
            # Use the first non-empty DataFrame to find column indices
            representative_df = next((df for df in list_of_dfs if not df.empty), None)
            if representative_df is None: continue

            try:
                x_col_idx = list(representative_df.columns).index(cfg['x_col'])
                y_col_idx = list(representative_df.columns).index(cfg['y_col'])
            except ValueError as e:
                self.logger.warning(f"Could not find column for chart '{cfg['title']}': {e}. Skipping chart.")
                continue

            for test_idx, df in enumerate(list_of_dfs):
                if df.empty: continue
                
                num_data_points = len(df)
                if num_data_points == 0: continue

                # Adjust for test label row if multiple tests
                if len(list_of_dfs) > 1:
                    current_data_row += 1

                series_name = f"Test {self.motor_tests[test_idx].test_number}"
                
                # Define automatic data ranges that adjust to actual data size
                # This ensures charts only include valid data points, not empty cells
                # Use explicit range references to avoid off-by-one or quoting issues
                x_values = [
                    self.sheet_name,
                    current_data_row,
                    col_offset + x_col_idx,
                    current_data_row + num_data_points - 1,
                    col_offset + x_col_idx,
                ]
                y_values = [
                    self.sheet_name,
                    current_data_row,
                    col_offset + y_col_idx,
                    current_data_row + num_data_points - 1,
                    col_offset + y_col_idx,
                ]
                
                # Validate that the y-column contains numeric, finite data before adding series
                try:
                    col_series = df.iloc[:, y_col_idx]
                    # Drop NaNs and check for at least one finite numeric value
                    numeric_vals = col_series.dropna()
                    if numeric_vals.empty or not np.isfinite(numeric_vals.astype(float)).any():
                        self.logger.warning(f"Skipping series '{series_name}' for chart '{cfg['title']}': no numeric data in column '{cfg['y_col']}'")
                        # Move current_data_row forward by num_data_points and continue
                        current_data_row += num_data_points
                        continue
                except Exception as e:
                    self.logger.debug(f"Could not validate numeric data for series '{series_name}': {e}")

                self.helper.add_chart_series(
                    chart=chart,
                    series_name=series_name,
                    x_range=x_values,
                    y_range=y_values
                )
                has_series = True
                current_data_row += num_data_points

            if has_series:
                insert_row, insert_col = chart_positions[chart_idx]
                self.helper.insert_chart_with_size(
                    chart=chart,
                    row=insert_row,
                    col=insert_col,
                    width=chart_pixel_width,
                    height=chart_pixel_height
                )
                max_chart_row_span = max(max_chart_row_span, 22) # 400px height is ~20 rows
        
        self.current_row = charts_block_start_row + max_chart_row_span * 2 + 2 # Update current row after all charts

    def _get_unit_info(self, measurement: str) -> Dict[str, str]:
        if self.unit_metadata and measurement in self.unit_metadata:
            return self.unit_metadata[measurement]

        config_attr = f"{measurement}_unit"
        selected_unit = getattr(self.config, config_attr, None)
        settings = UNIT_CONFIG[measurement]["settings"]
        if not selected_unit or selected_unit not in settings:
            selected_unit = next(iter(settings))

        info = settings[selected_unit]
        return {
            "column": info["label"],
            "axis_label": info["axis_label"],
            "chart_name": info["chart_name"],
            "unit": selected_unit,
        }

    def _add_test_lab_summary_section(self) -> None:
        """Add a compact Test-Lab (Carichi nominali) summary to the SAP sheet.

        Historically the report included a small "Carichi nominali (Test-Lab)"
        summary in each SAP performance sheet. Tests expect that a short
        summary and a few key markers (e.g. "Scheda SR Summary" and status
        messages like "CSV data reused from ...") are present in the
        workbook shared strings. This method re-introduces a compact,
        non-invasive section that lists linked test-lab workbooks and
        associated notes.
        """
        try:
            if not self.test_lab_summary_map:
                return

            start_col = self.data_start_col
            # Section title
            self.ws.write(self.current_row, start_col, "Carichi nominali (Test-Lab):", self.fmt.get('header'))
            self.current_row += 1

            # Iterate through the motor tests in sheet order and show any linked summaries
            for mt in self.all_motor_tests:
                tls = self.test_lab_summary_map.get(mt.test_number)
                if not tls:
                    continue

                # Show workbook filename if available
                if tls.source_path:
                    try:
                        name = Path(tls.source_path).name
                    except Exception:
                        name = str(tls.source_path)
                    self.ws.write(self.current_row, start_col, f"Workbook: {name}", self.fmt.get('cell'))
                    self.current_row += 1

                # Include the status message from MotorTestData (contains reuse notes)
                if getattr(mt, 'status_message', None):
                    self.ws.write(self.current_row, start_col, mt.status_message, self.fmt.get('cell'))
                    self.current_row += 1

                # Mark that a Scheda summary exists so tests can find the marker text
                if tls.scheda:
                    self.ws.write(self.current_row, start_col, "Scheda SR Summary", self.fmt.get('cell'))
                    self.current_row += 1

            # Small spacing after section
            self.current_row += 1

        except Exception as e:
            self.logger.error(f"Error while adding test-lab summary section: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())

    def _add_noise_section(self):
        if not self.all_noise_tests:
            self.logger.info(f"No noise tests found for SAP {self.sap_code}. Skipping noise section.")
            return

        # Add minimal space to keep noise charts close to performance data
        self.current_row += 1 # Reduced from 2 to 1 for better proximity
        start_col = self.data_start_col # Will be 0 (Column A)
        
        self.ws.write(self.current_row, start_col, "🔊 Noise Analysis Charts:", self.fmt.get('header'))
        self.ws.set_row(self.current_row, 20)
        self.current_row += 1
        
        # Add summary box at top of noise section
        self._add_noise_summary_box()
        
        # Check if we have TXT data for chart generation
        has_txt_data = self._check_for_txt_data()
        
        # Check if we have images
        has_images = any(
            getattr(noise_test, 'image_paths', []) or [getattr(noise_test, 'image_path', None)]
            for noise_test in self.all_noise_tests
        )
        
        if has_txt_data:
            # Generate charts from TXT files
            self._add_noise_charts()
        
        if has_images:
            # Add images (either alone or after TXT charts)
            self._add_noise_images()
        
        if not has_txt_data and not has_images:
            # Fall back to traditional table approach (for old data or errors)
            self._add_noise_table()
    
    def _add_noise_summary_box(self):
        """Add a summary box showing what noise tests were selected and any registry notes/errors."""
        try:
            start_col = self.data_start_col
            summary_start_row = self.current_row
            
            # Create summary box format
            summary_header_fmt = self.wb.add_format({
                'bold': True,
                'font_size': 11,
                'bg_color': '#E8F4F8',
                'border': 1,
                'border_color': '#0066CC',
                'align': 'left',
                'valign': 'vcenter'
            })
            
            summary_cell_fmt = self.wb.add_format({
                'font_size': 10,
                'bg_color': '#F5F9FC',
                'border': 1,
                'border_color': '#CCCCCC',
                'align': 'left',
                'valign': 'top',
                'text_wrap': True
            })
            
            warning_fmt = self.wb.add_format({
                'font_size': 10,
                'bg_color': '#FFF3CD',
                'font_color': '#856404',
                'border': 1,
                'border_color': '#CCCCCC',
                'align': 'left',
                'valign': 'top',
                'text_wrap': True,
                'bold': True
            })
            
            # Summary box title
            self.ws.merge_range(summary_start_row, start_col, summary_start_row, start_col + 7,
                               "📋 NOISE TEST SELECTION SUMMARY", summary_header_fmt)
            self.ws.set_row(summary_start_row, 22)
            self.current_row += 1
            
            # SAP Code
            self.ws.write(self.current_row, start_col, "SAP Code:", summary_header_fmt)
            self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 3,
                               self.sap_code or "Unknown", summary_cell_fmt)
            self.current_row += 1
            
            # Total tests selected
            self.ws.write(self.current_row, start_col, "Tests Selected:", summary_header_fmt)
            self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 3,
                               f"{len(self.all_noise_tests)} noise test(s)", summary_cell_fmt)
            self.current_row += 1
            
            # Test numbers list
            test_numbers = []
            for noise_test in self.all_noise_tests:
                test_num = (getattr(noise_test, 'test_number', None) or 
                           getattr(noise_test, 'nprova', None) or "Unknown")
                date = getattr(noise_test, 'date', None) or "N/A"
                test_numbers.append(f"Test {test_num} ({date})")
            
            self.ws.write(self.current_row, start_col, "Test Details:", summary_header_fmt)
            test_list = "\n".join(test_numbers)
            self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 7,
                               test_list, summary_cell_fmt)
            self.ws.set_row(self.current_row, 15 * len(test_numbers))  # Adjust row height
            self.current_row += 1
            
            # Registry notes (if available)
            notes_found = []
            errors_found = []
            
            for noise_test in self.all_noise_tests:
                # Check for notes
                notes = getattr(noise_test, 'notes', None) or getattr(noise_test, 'result', None)
                if notes and notes not in ["OK", "N/A", "Pass", ""]:
                    test_num = (getattr(noise_test, 'test_number', None) or 
                               getattr(noise_test, 'nprova', None) or "Unknown")
                    notes_found.append(f"Test {test_num}: {notes}")
                
                # Check for data type issues
                data_type = getattr(noise_test, 'data_type', 'images')
                test_num = (getattr(noise_test, 'test_number', None) or 
                           getattr(noise_test, 'nprova', None) or "Unknown")
                
                if data_type == 'none':
                    # No data at all - this is a real problem
                    errors_found.append(f"Test {test_num}: No data found (no images or TXT files)")
                elif data_type == 'images' and not getattr(noise_test, 'image_paths', []):
                    # Expected images but got none - check if TXT files exist instead
                    txt_files = getattr(noise_test, 'txt_files', [])
                    if not txt_files:
                        # No images AND no TXT files
                        errors_found.append(f"Test {test_num}: Images expected but not found")
                    # If TXT files exist, don't show warning - that's valid data
            
            # Display registry notes if any
            if notes_found:
                self.ws.write(self.current_row, start_col, "Registry Notes:", summary_header_fmt)
                notes_text = "\n".join(notes_found)
                self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 7,
                                   notes_text, warning_fmt)
                self.ws.set_row(self.current_row, 15 * len(notes_found))
                self.current_row += 1
            
            # Display errors/warnings if any
            if errors_found:
                self.ws.write(self.current_row, start_col, "⚠️ Warnings:", summary_header_fmt)
                errors_text = "\n".join(errors_found)
                self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 7,
                                   errors_text, warning_fmt)
                self.ws.set_row(self.current_row, 15 * len(errors_found))
                self.current_row += 1
            else:
                # All OK
                self.ws.write(self.current_row, start_col, "Status:", summary_header_fmt)
                self.ws.merge_range(self.current_row, start_col + 1, self.current_row, start_col + 7,
                                   "✅ All tests loaded successfully", summary_cell_fmt)
                self.current_row += 1
            
            # Add spacing after summary box
            self.current_row += 2
            
            self.logger.info(f"✅ Added noise summary box with {len(self.all_noise_tests)} tests, {len(notes_found)} notes, {len(errors_found)} warnings")
            
        except Exception as e:
            self.logger.error(f"Error creating noise summary box: {e}")
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            # Continue without summary box
            self.current_row += 1

    def _check_for_txt_data(self) -> bool:
        """Check if any noise tests have TXT data available for chart generation."""
        try:
            # Check directly in the NoiseTestInfo objects for txt_files
            self.logger.info(f"🔍 Checking {len(self.all_noise_tests)} pre-selected noise tests for TXT data")
            
            for noise_test in self.all_noise_tests:
                # Use the txt_files attribute directly from NoiseTestInfo
                txt_files = getattr(noise_test, 'txt_files', []) or []
                nprova = getattr(noise_test, 'nprova', None) or getattr(noise_test, 'test_number', None)
                
                if txt_files:
                    self.logger.info(f"✅ Found {len(txt_files)} TXT file(s) for noise test {nprova}, will generate charts")
                    return True
                else:
                    self.logger.info(f"❌ No TXT data found for noise test {nprova}")
            
            self.logger.info("No TXT data found for noise tests, will use traditional approach")
            return False
            
        except Exception as e:
            self.logger.error(f"Error checking for TXT data: {e}")
            return False

    def _add_noise_charts(self):
        """Add noise charts generated from TXT files using the chart generator."""
        try:
            import openpyxl
            
            # Use provided noise handler or create fallback
            if self.noise_handler:
                noise_handler = self.noise_handler
            else:
                from ..simplified_noise_handler import SimplifiedNoiseDataHandler
                noise_handler = SimplifiedNoiseDataHandler(self.config)
            
            # Collect all TXT files from all pre-selected noise tests
            all_noise_test_data = []
            
            self.logger.info(f"🔍 Processing {len(self.all_noise_tests)} pre-selected noise tests for chart generation")
            
            for noise_test in self.all_noise_tests:
                # Use the txt_files already discovered and stored in NoiseTestInfo
                txt_files = getattr(noise_test, 'txt_files', []) or []
                nprova = getattr(noise_test, 'nprova', None) or getattr(noise_test, 'test_number', None)
                
                if txt_files:
                    self.logger.info(f"✅ Processing {len(txt_files)} TXT files for test {nprova}")
                    
                    # Parse TXT files using the chart generator
                    # Check if chart_generator exists, if not create it
                    if not hasattr(noise_handler, 'chart_generator'):
                        from ..noise_chart_generator import NoiseChartGenerator
                        noise_handler.chart_generator = NoiseChartGenerator()
                        
                    for txt_file in txt_files:
                        self.logger.debug(f"📄 Parsing TXT file: {txt_file}")
                        parsed_data = noise_handler.chart_generator.parse_txt_file(txt_file)
                        if parsed_data:
                            all_noise_test_data.append(parsed_data)
                            self.logger.debug(f"✅ Successfully parsed {txt_file}")
                        else:
                            self.logger.warning(f"❌ Failed to parse {txt_file}")
                else:
                    self.logger.info(f"ℹ️ No TXT files found for test {nprova}")
            
            if not all_noise_test_data:
                self.logger.warning("No valid noise test data found for chart generation, falling back to table")
                self._add_noise_table()
                return
            
            self.logger.info(f"🔊 Creating noise charts from {len(all_noise_test_data)} TXT files across {len(self.all_noise_tests)} noise tests")
            
            # Convert XlsxWriter worksheet to something openpyxl-compatible for chart generation
            # Since we're using XlsxWriter, we'll need to create the charts using XlsxWriter's chart API
            self._create_noise_charts_xlsxwriter(all_noise_test_data)
            
        except Exception as e:
            self.logger.error(f"Error creating noise charts: {e}")
            # Fall back to traditional approach
            self._add_noise_table()

    def _create_noise_charts_xlsxwriter(self, noise_test_data):
        """Create noise charts using XlsxWriter's chart API - CLEAN VERTICAL LAYOUT WITH PROPER SPACING."""
        try:
            # Header for charts section
            self.ws.write(self.current_row, 0, f"🔊 Noise Analysis Charts ({len(noise_test_data)} measurements)", self.fmt.get('info_label'))
            self.current_row += 3
            
            charts_created = 0
            chart_spacing = 22  # Rows between each chart group (chart height 18 + spacing)
            
            self.logger.info(f"📊 Creating {len(noise_test_data)} noise charts with clean vertical layout")
            
            # Create border format for visual separation
            border_format = self.wb.add_format({
                'border': 1,
                'border_color': '#CCCCCC',
                'bg_color': '#F8F9FA'
            })
            
            for test_idx, test_data in enumerate(noise_test_data):
                # Calculate row position for this chart
                chart_start_row = self.current_row + (test_idx * chart_spacing)
                
                # Add filename title with background
                title = f"📄 {test_data.filename}"
                self.ws.write(chart_start_row, 0, title, self.fmt.get('header'))
                
                # Add separator line below title
                for col in range(8):
                    self.ws.write(chart_start_row + 1, col, "", border_format)
                
                # Position chart below title
                chart_row = chart_start_row + 2
                
                # Write chart data to hidden columns (far right)
                data_col = 35 + (test_idx * 10)  # Columns AI, AS, BC, etc.
                sampled_data = test_data.get_sampled_data(sample_rate=15)
                data_points = min(len(sampled_data), 40)  # Reasonable number of points
                
                # Write frequency and microphone data
                self.ws.write(chart_row - 1, data_col, "Frequency", self.fmt.get('header'))
                # Process ALL microphones (up to 5)
                available_mics = len(sampled_data[0][1]) if sampled_data else 0
                max_mics = min(5, available_mics)  # Use all available mics up to 5
                
                for mic_idx in range(max_mics):
                    self.ws.write(chart_row - 1, data_col + 1 + mic_idx, f"Mic {mic_idx + 1}", self.fmt.get('header'))
                
                for i in range(data_points):
                    freq, mic_data = sampled_data[i]
                    self.ws.write(chart_row + i, data_col, freq)
                    for mic_idx in range(max_mics):
                        value = mic_data[mic_idx] if mic_idx < len(mic_data) else 0.0
                        self.ws.write(chart_row + i, data_col + 1 + mic_idx, value)
                
                # Create frequency response chart
                try:
                    freq_chart = self.wb.add_chart({'type': 'line'})
                    if freq_chart:
                        freq_chart.set_title({
                            'name': f'Frequency Response - {test_data.filename[:40]}',
                            'name_font': {'size': 12, 'bold': True}
                        })
                        freq_chart.set_x_axis({
                            'name': 'Frequency (Hz)',
                            'name_font': {'size': 10},
                            'num_font': {'size': 9},
                            'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}},
                            'major_tick_mark': 'outside',
                            'line': {'color': '#595959', 'width': 1.0}
                        })
                        freq_chart.set_y_axis({
                            'name': 'Sound Level (dB)',
                            'name_font': {'size': 10},
                            'num_font': {'size': 9},
                            'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}},
                            'major_tick_mark': 'outside',
                            'line': {'color': '#595959', 'width': 1.0}
                        })
                        freq_chart.set_size({'width': 480, 'height': 320})
                        freq_chart.set_legend({'position': 'bottom', 'font': {'size': 9}})
                        
                        # Add ALL microphone series (up to 5 microphones)
                        mic_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
                        for mic_idx in range(max_mics):
                            freq_chart.add_series({
                                'name': f'Microphone {mic_idx + 1}',
                                'categories': [self.sheet_name, chart_row, data_col, chart_row + data_points - 1, data_col],
                                'values': [self.sheet_name, chart_row, data_col + 1 + mic_idx, chart_row + data_points - 1, data_col + 1 + mic_idx],
                                'line': {'color': mic_colors[mic_idx], 'width': 2}
                            })
                    
                    # Create bar chart for overall levels
                    bar_chart = self.wb.add_chart({'type': 'column'})
                    if bar_chart:
                        bar_chart.set_title({
                            'name': f'Overall Sound Levels - {test_data.filename[:30]}',
                            'name_font': {'size': 11, 'bold': True}
                        })
                        bar_chart.set_x_axis({
                            'name': 'Measurement Type',
                            'num_font': {'size': 9},
                            'major_tick_mark': 'outside',
                            'line': {'color': '#595959', 'width': 1.0}
                        })
                        bar_chart.set_y_axis({
                            'name': 'Level (dB)',
                            'num_font': {'size': 9},
                            'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}},
                            'major_tick_mark': 'outside',
                            'line': {'color': '#595959', 'width': 1.0}
                        })
                        bar_chart.set_size({'width': 300, 'height': 320})
                        bar_chart.set_legend({'position': 'top', 'font': {'size': 9}})  # Enable legend with proper positioning
                        
                        # Write bar chart data (position after all microphone columns)
                        bar_data_col = data_col + max_mics + 1  # After frequency + all microphones
                        self.ws.write(chart_row - 1, bar_data_col, "Sound Pressure")
                        self.ws.write(chart_row - 1, bar_data_col + 1, "Sound Power")
                        self.ws.write(chart_row, bar_data_col, test_data.overall_sound_pressure)
                        self.ws.write(chart_row, bar_data_col + 1, test_data.overall_sound_power)
                        
                        bar_chart.add_series({
                            'name': 'Sound Pressure',
                            'categories': [self.sheet_name, chart_row - 1, bar_data_col, chart_row - 1, bar_data_col],
                            'values': [self.sheet_name, chart_row, bar_data_col, chart_row, bar_data_col],
                            'fill': {'color': '#2ca02c'}
                        })
                        bar_chart.add_series({
                            'name': 'Sound Power',
                            'categories': [self.sheet_name, chart_row - 1, bar_data_col + 1, chart_row - 1, bar_data_col + 1],
                            'values': [self.sheet_name, chart_row, bar_data_col + 1, chart_row, bar_data_col + 1],
                            'fill': {'color': '#d62728'}
                        })
                    
                    # Insert charts side by side if both were created successfully
                    if freq_chart and bar_chart:
                        self.ws.insert_chart(chart_row, 0, freq_chart)
                        self.ws.insert_chart(chart_row, 8, bar_chart)
                        charts_created += 2
                        self.logger.info(f"✅ Created charts for {test_data.filename} at row {chart_row}")
                    else:
                        self.logger.warning(f"⚠️ Failed to create one or both charts for {test_data.filename}")
                    
                except Exception as chart_error:
                    self.logger.error(f"❌ Error creating charts for {test_data.filename}: {chart_error}")
            
            # Update current row to after all charts
            total_chart_height = len(noise_test_data) * chart_spacing
            self.current_row = self.current_row + total_chart_height + 5
            
            # Add summary
            self.ws.write(self.current_row, 0, f"📊 Successfully created {charts_created//2} chart pairs for {len(noise_test_data)} noise measurements", self.fmt.get('cell'))
            self.current_row += 2
            
            self.logger.info(f"✅ Completed noise chart generation: {charts_created//2} chart pairs created without overlaps")
            
        except Exception as e:
            self.logger.error(f"❌ Error in noise chart creation: {e}")
            import traceback
            self.logger.error(f"Traceback: {traceback.format_exc()}")
            self.current_row += 5

    def _add_noise_table(self):
        if not self.all_noise_tests:
            self.logger.info(f"No noise tests found for SAP {self.sap_code}. Skipping noise section.")
            return

        self.current_row += 2 # Add space before noise section
        start_col = self.data_start_col # Will be 0 (Column A)
        
        self.ws.write(self.current_row, start_col, "Noise Test Results:", self.fmt.get('header'))
        self.ws.set_row(self.current_row, 20)
        self.current_row += 1
        
        # --- Noise Info Table ---
        noise_info_start_row = self.current_row
        noise_headers = ["Test No.", "Date", "Mic Position", "Background Noise (dBA)", "Motor Noise (dBA)", "Result"]
        for i, header in enumerate(noise_headers):
            self.ws.write(noise_info_start_row, start_col + i, header, self.fmt.get('motor_info_header'))
        
        for idx, noise_test in enumerate(self.all_noise_tests):
            row = noise_info_start_row + 1 + idx
            
            # Safely handle potentially None or invalid values
            # Debug: Log the type of noise_test object and available attributes
            self.logger.debug(f"Processing noise test object: {type(noise_test).__name__}")
            if hasattr(noise_test, '__dict__'):
                self.logger.debug(f"Noise test attributes: {list(noise_test.__dict__.keys())}")
            
            # Handle different attribute names (old vs new)
            test_number = (getattr(noise_test, 'test_number', None) or 
                          getattr(noise_test, 'nprova', None) or "N/A")
            date = getattr(noise_test, 'date', None) or "N/A"
            mic_position = getattr(noise_test, 'mic_position', None) or "N/A"
            result = getattr(noise_test, 'result', None) or "N/A"
            
            self.logger.debug(f"Noise test {idx}: test_number={test_number}, date={date}")
            
            self.ws.write(row, start_col + 0, test_number, self.fmt.get('cell'))
            self.ws.write(row, start_col + 1, date, self.fmt.get('cell'))
            self.ws.write(row, start_col + 2, mic_position, self.fmt.get('cell'))
            
            # Handle background noise with proper validation
            background_noise = getattr(noise_test, 'background_noise', None)
            if background_noise is None:
                self.ws.write(row, start_col + 3, "N/A", self.fmt.get('cell'))
            elif isinstance(background_noise, str):
                self.ws.write(row, start_col + 3, background_noise, self.fmt.get('cell'))
            elif isinstance(background_noise, (int, float)):
                if pd.isna(background_noise) or not np.isfinite(background_noise):
                    self.ws.write(row, start_col + 3, "N/A", self.fmt.get('cell'))
                else:
                    self.ws.write(row, start_col + 3, background_noise, self.fmt.get('decimal_2'))
            else:
                self.ws.write(row, start_col + 3, str(background_noise), self.fmt.get('cell'))
            
            # Handle motor noise with proper validation
            motor_noise = getattr(noise_test, 'motor_noise', None)
            if motor_noise is None:
                self.ws.write(row, start_col + 4, "N/A", self.fmt.get('cell'))
            elif isinstance(motor_noise, str):
                self.ws.write(row, start_col + 4, motor_noise, self.fmt.get('cell'))
            elif isinstance(motor_noise, (int, float)):
                if pd.isna(motor_noise) or not np.isfinite(motor_noise):
                    self.ws.write(row, start_col + 4, "N/A", self.fmt.get('cell'))
                else:
                    self.ws.write(row, start_col + 4, motor_noise, self.fmt.get('decimal_2'))
            else:
                self.ws.write(row, start_col + 4, str(motor_noise), self.fmt.get('cell'))
            
            self.ws.write(row, start_col + 5, result, self.fmt.get('cell'))
        
        self.current_row = noise_info_start_row + len(self.all_noise_tests) + 2

        # --- Noise Images ---
        self._add_noise_images()

    def _add_noise_images(self):
        """Dynamically add all noise test images with section title and dynamic layout."""
        if not self.all_noise_tests:
            return
            
        start_col = self.data_start_col  # Will be 0 (Column A)
        
        # Add section title for the images
        self.ws.write(self.current_row, start_col, "Noise Test Images", self.fmt.get('section_header'))
        self.current_row += 2  # Add spacing after title
        
        # Collect all available images first
        all_images = []
        for noise_test in self.all_noise_tests:
            # Handle different attribute names
            test_number = (getattr(noise_test, 'test_number', None) or 
                          getattr(noise_test, 'nprova', None) or "Unknown")
            
            # Check for image path or image paths
            image_path = getattr(noise_test, 'image_path', None)
            image_paths = getattr(noise_test, 'image_paths', [])
            
            # Collect all available images for this test
            available_images = []
            if image_path and Path(image_path).exists():
                available_images.append(str(image_path))
            if image_paths:
                for img_path in image_paths:
                    if Path(img_path).exists() and str(img_path) not in available_images:
                        available_images.append(str(img_path))
            
            # Add to master list with metadata - use actual filename as label
            for actual_image_path in available_images:
                # Extract clean filename from path
                filename = Path(actual_image_path).stem  # Gets filename without extension
                # Create a clean label from filename
                label = filename.replace('_', ' ').replace('GRAFICO ', '').strip()
                if not label:  # Fallback if filename is empty
                    label = f"Test {test_number} Image"

                all_images.append({
                    'path': actual_image_path,
                    'label': label,
                    'test_number': test_number
                })
        
        if not all_images:
            self.logger.warning("No noise test images found to add")
            return
        
        # Dynamic layout configuration
        images_per_row = 2  # Can be adjusted as needed
        cols_per_image = 6  # Excel columns per image
        rows_per_image = 22  # Excel rows per image (including label space)
        label_offset = 1  # Rows above image for label
        
        images_start_row = self.current_row
        max_row_used = images_start_row
        
        # Calculate and place images dynamically
        for img_idx, image_info in enumerate(all_images):
            # Calculate grid position
            row_group = img_idx // images_per_row
            col_position = img_idx % images_per_row
            
            # Calculate actual Excel position
            insert_row = images_start_row + (row_group * rows_per_image) + label_offset
            insert_col = start_col + (col_position * cols_per_image)
            
            # Add label above the image
            label_row = insert_row - label_offset
            self.ws.write(label_row, insert_col, image_info['label'], self.fmt.get('noise_label'))
            
            # Insert the image
            try:
                img_path_str = str(image_info['path'])
                self.ws.insert_image(insert_row, insert_col, img_path_str, {
                    'x_scale': 0.5, 'y_scale': 0.5,  # Adjust scaling as needed
                    'object_position': 1
                })
                max_row_used = max(max_row_used, insert_row + rows_per_image - label_offset)
                self.logger.debug(f"Inserted noise image at row {insert_row}, col {insert_col}: {image_info['label']}")
            except Exception as e:
                self.logger.error(f"Could not insert noise image {image_info['path']}: {e}")
                self.ws.write(insert_row, insert_col, "Image Error", self.fmt.get('red_highlight'))
                max_row_used = max(max_row_used, insert_row + 2)
        
        # Update current row to be after all images
        self.current_row = max_row_used + 2  # Add some spacing after images
        
        self.logger.info(f"Added {len(all_images)} noise test images in dynamic layout")
