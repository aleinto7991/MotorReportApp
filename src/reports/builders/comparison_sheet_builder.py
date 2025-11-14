import logging
from typing import List, Dict, Optional, Tuple
import pandas as pd
from xlsxwriter.workbook import Workbook as XlsxWorkbook

from ...config.app_config import AppConfig
from ...data.models import MotorTestData
from ...config.measurement_units import apply_unit_preferences, UNIT_CONFIG
from .excel_formatter import ExcelFormatter
from .excel_sheet_helper import ExcelSheetHelper

logger = logging.getLogger(__name__)

class ComparisonSheetBuilder:
    """Builds comparison sheet with performance charts referencing data from individual SAP sheets."""
    
    def __init__(self, workbook: XlsxWorkbook, comparison_data: Dict[str, List[MotorTestData]], 
                 formatter: ExcelFormatter, logo_colors: List[str], sap_sheet_name_map: Dict[str, str],
                 config: Optional[AppConfig] = None,
                 custom_sheet_name: Optional[str] = None, custom_title: Optional[str] = None, 
                 custom_description: Optional[str] = None):
        self.workbook = workbook
        self.comparison_data = comparison_data
        self.formatter = formatter
        self.logo_colors = logo_colors
        self.sap_sheet_name_map = sap_sheet_name_map
        self.config = config
        self.sap_codes = list(comparison_data.keys())
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # Custom naming support for multiple comparisons
        self.custom_title = custom_title
        self.custom_description = custom_description
        
        # Track data locations for chart generation
        self.data_ranges = []  # Will store info about where data is in the worksheet
        
        # Create the worksheet
        self.sheet_name = custom_sheet_name or "SAP Comparison"
        self.worksheet = workbook.add_worksheet(self.sheet_name)
        self.helper = ExcelSheetHelper(self.worksheet, self.workbook, self.formatter)
        self.unit_metadata: Optional[Dict[str, Dict[str, str]]] = None
        
        # Set tab color
        if logo_colors:
            try:
                self.worksheet.set_tab_color(logo_colors[0])
            except Exception as e:
                self.logger.debug(f"Could not set tab color: {e}")

    def build(self):
        """Build the comparison sheet content with performance charts and data tables."""
        try:
            self.logger.info(f"Building comparison sheet for {len(self.sap_codes)} SAP codes")
            self.logger.info(f"SAP codes: {self.sap_codes}")
            self.logger.info(f"Comparison data keys: {list(self.comparison_data.keys())}")
            
            # Check if we have any data to compare
            if not self.comparison_data or all(not tests for tests in self.comparison_data.values()):
                self.logger.warning("No comparison data available")
                self.worksheet.write(0, 0, "No comparison data available", self.formatter.get('info_label'))
                return
            
            # Set up the worksheet with same column widths as performance sheet
            # Use 15 units for all columns to match performance sheet layout
            self.helper.set_column_widths({'A:P': 15})
            
            # Title section
            self.logger.info("Writing title section...")
            row = self._write_title()
            self.logger.info(f"Title written, next row: {row}")
            
            # Add motor information comparison section
            self.logger.info("Adding motor information comparison...")
            row = self._add_motor_info_comparison(row)
            self.logger.info(f"Motor info comparison complete, next row: {row}")
            
            # Add performance data comparison tables
            self.logger.info("Adding performance data comparison...")
            row = self._add_performance_data_comparison(row)
            self.logger.info(f"Performance data comparison complete, next row: {row}")
            
            # Create comparison charts
            self.logger.info("Creating comparison charts...")
            row = self._create_comparison_charts(row)
            self.logger.info("Comparison charts created")
            
            self.logger.info("Comparison sheet completed successfully")
            
        except Exception as e:
            self.logger.error(f"Error building comparison sheet: {e}", exc_info=True)

    def _write_title(self) -> int:
        """Write the title section and return next available row."""
        try:
            self.logger.info("Writing comparison sheet title...")
            row = 0
            
            # Main title - use custom title if provided
            title = self.custom_title or "SAP Code Performance Comparison"
            self.worksheet.merge_range(row, 0, row, 10, title, 
                                       self.formatter.get('report_title'))
            row += 2
            self.logger.info(f"Main title written at row {row-2}")
            
            # Add custom description if provided
            if self.custom_description:
                self.worksheet.merge_range(row, 0, row, 10, f"Description: {self.custom_description}", 
                                         self.formatter.get('info_label'))
                row += 2
            
            # Comparison summary with test lab details
            comparison_summary = []
            for sap_code in self.sap_codes:
                sap_tests = self.comparison_data.get(sap_code, [])
                test_numbers = [test.test_number for test in sap_tests if test.csv_data is not None]
                if test_numbers:
                    comparison_summary.append(f"SAP {sap_code} ({len(test_numbers)} tests: {', '.join(test_numbers)})")
            
            summary_text = "Comparing: " + " | ".join(comparison_summary)
            self.worksheet.merge_range(row, 0, row, 10, summary_text, 
                                       self.formatter.get('info_label'))
            row += 2
            
            # Instructions
            self.worksheet.merge_range(row, 0, row, 10, 
                                       "Side-by-side comparison of selected test labs with motor information, performance data, and charts", 
                                       self.formatter.get('cell'))
            row += 2
            
            self.logger.info(f"Title section complete, returning row {row}")
            return row
            
        except Exception as e:
            self.logger.error(f"Error writing comparison title: {e}", exc_info=True)
            return 5  # fallback row
    
    def _create_comparison_charts(self, start_row: int):
        """Create comparison charts for motor performance metrics matching the performance sheet structure."""
        current_row = start_row
        
        # Add visual separation before charts
        self.worksheet.write(current_row, 0, "", self.formatter.get('cell'))
        current_row += 1
        
        self.worksheet.write(current_row, 0, "Performance Charts Comparison:", self.formatter.get('header'))
        self.worksheet.set_row(current_row, 20)
        charts_start_row = current_row + 1
        
        flow_info = self._get_unit_info('flow')
        pressure_info = self._get_unit_info('pressure')
        power_info = self._get_unit_info('power')
        speed_info = self._get_unit_info('speed')

        chart_configs = [
            {
                'title': f"{pressure_info['chart_name']} vs {flow_info['chart_name']} - Comparison",
                'x_col': flow_info['column'],
                'y_col': pressure_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': pressure_info['axis_label'],
            },
            {
                'title': f"Efficiency (%) vs {flow_info['chart_name']} - Comparison",
                'x_col': flow_info['column'],
                'y_col': 'Efficiency (%)',
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': 'Efficiency (%)',
            },
            {
                'title': f"{power_info['chart_name']} vs {flow_info['chart_name']} - Comparison",
                'x_col': flow_info['column'],
                'y_col': power_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': power_info['axis_label'],
            },
            {
                'title': f"{speed_info['chart_name']} vs {flow_info['chart_name']} - Comparison",
                'x_col': flow_info['column'],
                'y_col': speed_info['column'],
                'x_axis_label': flow_info['axis_label'],
                'y_axis_label': speed_info['axis_label'],
            },
        ]
        
        # Chart layout: 2 charts per row, same as performance sheet
        chart_pixel_width = 680
        chart_pixel_height = 400
        cols_per_chart = 6
        col_offset = 0  # Start from column A, same as performance sheet
        
        chart_positions = [
            (charts_start_row, col_offset + 0*cols_per_chart),             # Chart 1: Vacuum vs Air Flow
            (charts_start_row, col_offset + 1*cols_per_chart),             # Chart 2: Efficiency vs Air Flow  
            (charts_start_row + 22, col_offset + 0*cols_per_chart),        # Chart 3: Power vs Air Flow
            (charts_start_row + 22, col_offset + 1*cols_per_chart)         # Chart 4: Speed vs Air Flow
        ]
        
        for chart_idx, config in enumerate(chart_configs):
            try:
                chart_row, chart_col = chart_positions[chart_idx]
                self._create_comparison_chart(
                    chart_row, chart_col,
                    config['title'],
                    config['x_col'],
                    config['x_axis_label'],
                    config['y_axis_label'],
                    config['y_col']
                )
            except Exception as e:
                self.logger.error(f"Error creating comparison chart {config['title']}: {e}")
        
        # Return position after all charts (charts are about 22 rows tall, 2 rows of charts)
        # Same calculation as performance sheet: 22 * 2 + 2 = 46
        return charts_start_row + 46

    def _create_comparison_chart(self, start_row: int, start_col: int, title: str,
                                 x_col_name: str, x_axis_title: str, y_axis_title: str,
                                 y_col_name: str) -> int:
        """Create a comparison chart using worksheet data ranges for selected test labs."""
        try:
            # Create chart using helper
            chart = self.helper.create_chart(title, x_axis_title, y_axis_title)
            if not chart:
                self.logger.error(f"Failed to create chart: {title}")
                self.worksheet.write(start_row, start_col, f"{title} chart error: Could not create chart object.", self.formatter.get('cell'))
                return start_row + 2

            colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
            series_added = 0

            # Add data series using worksheet ranges for each test lab
            for data_range in self.data_ranges:
                try:
                    columns = data_range['columns']
                    
                    # Debug: Log available columns for this data range
                    self.logger.debug(f"Data range columns for {data_range['sap_code']}-{data_range['test_number']}: {columns}")
                    
                    # Find column indices for x and y data
                    try:
                        x_col_idx = columns.index(x_col_name)
                        y_col_idx = columns.index(y_col_name)
                        self.logger.debug(f"Chart {title}: X column '{x_col_name}' at index {x_col_idx}, Y column '{y_col_name}' at index {y_col_idx}")
                    except ValueError as ve:
                        self.logger.warning(f"Missing columns for chart {title}, test {data_range['test_number']}: {ve}")
                        self.logger.warning(f"Looking for X='{x_col_name}', Y='{y_col_name}' in columns: {columns}")
                        continue
                    
                    # Create series name that shows both SAP and test number for clarity
                    series_name = f"SAP {data_range['sap_code']} - Test {data_range['test_number']}"
                    color = colors[series_added % len(colors)]
                    
                    # Define data ranges for this test using Excel references (zero-based indices)
                    # Use automatic range that adjusts to actual data size
                    start_data_row = data_range['start_row']
                    end_data_row = start_data_row + data_range['num_rows'] - 1

                    # xlsxwriter expects list-format ranges to be zero-based indices:
                    # [sheet_name, first_row, first_col, last_row, last_col]
                    x_values = [self.sheet_name, start_data_row, x_col_idx, end_data_row, x_col_idx]
                    y_values = [self.sheet_name, start_data_row, y_col_idx, end_data_row, y_col_idx]
                    # Log both the internal list-format ranges and a human-friendly A1-style range
                    try:
                        a1_x = self._range_list_to_a1(x_values)
                        a1_y = self._range_list_to_a1(y_values)
                    except Exception:
                        a1_x = None
                        a1_y = None

                    self.logger.info(f"Adding chart series '{series_name}' -> X(list)={x_values}, X(A1)={a1_x}; Y(list)={y_values}, Y(A1)={a1_y}")
                    self.logger.debug(f"Chart series '{series_name}': X={x_values}, Y={y_values}")
                    
                    # Add series to chart using helper
                    self.helper.add_chart_series(
                        chart=chart,
                        series_name=series_name,
                        x_range=x_values,
                        y_range=y_values,
                        color=color
                    )
                    series_added += 1
                    self.logger.debug(f"Successfully added series {series_added} for {series_name}")
                    
                except Exception as e:
                    self.logger.warning(f"Error adding series for {data_range['sap_code']} - {data_range['test_number']}: {e}")
                    continue

            if series_added > 0:
                # Insert chart using helper
                self.helper.insert_chart_with_size(chart, start_row, start_col)
                self.logger.info(f"Created comparison chart '{title}' with {series_added} data series")
            else:
                self.logger.warning(f"No data series added to {title} chart")
                self.worksheet.write(start_row, start_col, f"{title}: No data available", self.formatter.get('info_label'))

            return start_row + 22
            
        except Exception as e:
            self.logger.error(f"Error creating {title} chart: {e}")
            self.worksheet.write(start_row, start_col, f"{title} chart error: {str(e)}", self.formatter.get('cell'))
            return start_row + 2

    def _add_motor_info_comparison(self, start_row: int) -> int:
        """Add motor information comparison section showing all selected test labs side-by-side."""
        try:
            current_row = start_row
            self.worksheet.write(current_row, 0, "Motor Information Comparison", self.formatter.get('header'))
            self.worksheet.set_row(current_row, 20)
            current_row += 2
            
            # Collect all test data for comparison
            all_test_data = []
            for sap_code in self.sap_codes:
                sap_tests = self.comparison_data.get(sap_code, [])
                for test in sap_tests:
                    if test.csv_data is not None:  # Only include tests with performance data
                        all_test_data.append({
                            'sap_code': sap_code,
                            'test': test,
                            'test_number': test.test_number,
                            'date': test.inf_data.date if test.inf_data else 'N/A',
                            'voltage': test.inf_data.voltage if test.inf_data else 'N/A',
                            'frequency': test.inf_data.hz if test.inf_data else 'N/A',
                            'comments': test.inf_data.comment if test.inf_data else 'N/A'
                        })
            
            if not all_test_data:
                self.worksheet.write(current_row, 0, "No test data available for comparison", self.formatter.get('info_label'))
                return current_row + 2
            
            # Create comparison table headers
            headers = ["SAP Code", "Test No.", "Date", "Voltage (V)", "Frequency (Hz)", "Comments"]
            for i, header in enumerate(headers):
                self.worksheet.write(current_row, i, header, self.formatter.get('motor_info_header'))
            
            current_row += 1
            
            # Write test data rows
            for test_data in all_test_data:
                self.worksheet.write(current_row, 0, test_data['sap_code'], self.formatter.get('motor_info_value'))
                self.worksheet.write(current_row, 1, test_data['test_number'], self.formatter.get('motor_info_value'))
                self.worksheet.write(current_row, 2, test_data['date'], self.formatter.get('motor_info_value'))
                
                # Handle voltage formatting
                voltage_val = test_data['voltage']
                try:
                    voltage_val = float(voltage_val)
                except (ValueError, TypeError):
                    pass
                self.worksheet.write(current_row, 3, voltage_val, self.formatter.get('motor_info_value'))
                
                # Handle frequency formatting
                hz_val = test_data['frequency']
                try:
                    hz_val = float(hz_val)
                except (ValueError, TypeError):
                    pass
                self.worksheet.write(current_row, 4, hz_val, self.formatter.get('motor_info_value'))
                
                # Comments (truncate if too long)
                comments = str(test_data['comments']) if test_data['comments'] else 'N/A'
                if len(comments) > 60:
                    comments = comments[:57] + "..."
                self.worksheet.write(current_row, 5, comments, self.formatter.get('motor_info_value'))
                
                current_row += 1
            
            # Add visual separation after motor info table
            current_row += 2  # Add some space
            
            # Add a colored separator row for visual separation
            for col in range(6):
                self.worksheet.write(current_row, col, "", self.formatter.get('motor_info_header'))
            current_row += 2
            
            return current_row
            
        except Exception as e:
            self.logger.error(f"Error adding motor info comparison: {e}")
            return start_row + 2

    def _add_performance_data_comparison(self, start_row: int) -> int:
        """Add performance data comparison tables showing all selected test labs side-by-side."""
        try:
            current_row = start_row
            self.worksheet.write(current_row, 0, "Performance Data Comparison", self.formatter.get('header'))
            self.worksheet.set_row(current_row, 20)
            current_row += 2
            
            # Collect all performance data and track ranges
            all_performance_data = []
            self.data_ranges = []  # Reset data ranges
            
            for sap_code in self.sap_codes:
                sap_tests = self.comparison_data.get(sap_code, [])
                for test in sap_tests:
                    if test.csv_data is not None and not test.csv_data.empty:
                        self.logger.debug(f"Processing test {test.test_number} from SAP {sap_code}")
                        self.logger.debug(f"Original columns: {list(test.csv_data.columns)}")
                        
                        # Apply configured unit preferences to align with SAP sheets
                        test_df, metadata = self._convert_performance_data(test.csv_data.copy())

                        if self.unit_metadata is None:
                            self.unit_metadata = metadata
                        
                        self.logger.debug(f"Columns after conversion: {list(test_df.columns)}")
                        
                        # Add SAP code and test number columns to the data
                        test_df.insert(0, 'Test No.', test.test_number)
                        test_df.insert(0, 'SAP Code', sap_code)
                        
                        # Track this data range for chart generation
                        data_range_info = {
                            'sap_code': sap_code,
                            'test_number': test.test_number,
                            'start_row': current_row + 1 + len(all_performance_data),  # Will be updated below
                            'num_rows': len(test_df),
                            'columns': list(test_df.columns)
                        }
                        self.logger.debug(f"Data range info: {data_range_info}")
                        self.data_ranges.append(data_range_info)
                        all_performance_data.append(test_df)
            
            if not all_performance_data:
                self.worksheet.write(current_row, 0, "No performance data available for comparison", self.formatter.get('info_label'))
                return current_row + 2
            
            flow_info = self._get_unit_info('flow')
            pressure_info = self._get_unit_info('pressure')
            power_info = self._get_unit_info('power')
            speed_info = self._get_unit_info('speed')

            # Combine all performance data (already converted with apply_unit_preferences)
            # Use SAME column order as SAP performance sheets - no reordering
            # SAP Code and Test No. were already inserted at position 0 and 1
            combined_df = pd.concat(all_performance_data, ignore_index=True)
            converted_df = combined_df
            
            self.logger.info(f"Comparison sheet using EXACT same column order as SAP sheets")
            self.logger.info(f"Columns: {list(converted_df.columns)[:10]}")
            
            # Write table headers - NO COLUMN LIMIT (removed the arbitrary 16-column restriction)
            self.logger.info(f"Writing ALL {len(converted_df.columns)} column headers to Excel")
            headers_written = []
            for c_idx, col_name in enumerate(converted_df.columns):
                self.worksheet.write(current_row, c_idx, col_name, self.formatter.get('header'))
                headers_written.append(f"{c_idx}:{col_name}")
            
            self.logger.info(f"Headers written to Excel: {headers_written}")
            
            # Verify the converted columns are being written
            chart_columns = [flow_info['column'], pressure_info['column'], 'Efficiency (%)', power_info['column'], speed_info['column']]
            for col in chart_columns:
                if col in converted_df.columns:
                    col_idx = list(converted_df.columns).index(col)
                    self.logger.info(f"✅ Chart column '{col}' will be written to Excel at column {col_idx}")
                else:
                    self.logger.error(f"❌ Chart column '{col}' missing from combined DataFrame!")
            
            
            data_start_row = current_row + 1
            
            # Update data ranges with correct start rows AND reordered columns
            current_data_row = data_start_row
            for i, data_range in enumerate(self.data_ranges):
                data_range['start_row'] = current_data_row
                # CRITICAL FIX: Update columns to match the reordered DataFrame
                data_range['columns'] = list(converted_df.columns)
                self.logger.debug(f"Updated data range {i} columns to match reordered DataFrame: {data_range['columns']}")
                current_data_row += data_range['num_rows']
            
            # Write data values with proper formatting and visual grouping by test lab
            self.logger.info(f"Writing {len(converted_df)} rows of ALL {len(converted_df.columns)} columns to Excel")
            rows_written = 0
            current_sap_test = None
            use_alternating = False
            
            for r_idx, row in enumerate(converted_df.itertuples(index=False)):
                # Track when we change test labs for visual grouping
                row_sap_test = f"{row[0]}_{row[1]}"  # Combine SAP Code and Test No.
                if current_sap_test != row_sap_test:
                    current_sap_test = row_sap_test
                    use_alternating = not use_alternating  # Toggle shading for each new test
                
                for c_idx, value in enumerate(row):
                    # Write ALL columns (removed the 16-column limit)
                    # Determine format based on data type and column name
                    fmt_name = 'cell_alt' if use_alternating else 'cell'
                    
                    if isinstance(value, (int, float)) and not pd.isna(value):
                        if isinstance(value, int):
                            fmt_name = 'integer_alt' if use_alternating else 'integer'
                        else:
                            fmt_name = 'decimal_2_alt' if use_alternating else 'decimal_2'
                    
                    # Special formatting for efficiency/percentage columns
                    col_name_lower = str(converted_df.columns[c_idx]).lower()
                    if 'efficiency' in col_name_lower or '%' in col_name_lower:
                        # Check if value is already in percentage format (0-100 range) vs decimal format (0-1 range)
                        if isinstance(value, (int, float)) and not pd.isna(value):
                            if value <= 1.0:
                                # Value is in decimal format (e.g., 0.855), use percent format to convert to %
                                fmt_name = 'percent_alt' if use_alternating else 'percent'
                            else:
                                # Value is already in percentage format (e.g., 85.5), use decimal format
                                fmt_name = 'decimal_2_alt' if use_alternating else 'decimal_2'
                        else:
                            fmt_name = 'decimal_2_alt' if use_alternating else 'decimal_2'
                    
                    # Use base format if _alt version doesn't exist
                    try:
                        cell_format = self.formatter.get(fmt_name)
                    except:
                        fmt_name = fmt_name.replace('_alt', '')
                        cell_format = self.formatter.get(fmt_name)
                    
                    self.worksheet.write(data_start_row + r_idx, c_idx, value, cell_format)
                
                rows_written += 1
                
                # Debug: Log the first few rows with converted column values
                if r_idx < 3:
                    chart_col_values = []
                    for col in ['Air Flow (m³/h)', 'Vacuum Corrected (kPa)', 'Efficiency (%)', 'Power Corrected Watts in', 'Speed RPM']:
                        if col in converted_df.columns:
                            col_idx = list(converted_df.columns).index(col)
                            if col_idx < len(row):
                                chart_col_values.append(f"{col}={row[col_idx]}")
                    if chart_col_values:
                        self.logger.info(f"Row {r_idx+1} chart values: {', '.join(chart_col_values)}")
            
            self.logger.info(f"Successfully wrote {rows_written} rows to Excel")
            
            # Apply conditional formatting if available
            if hasattr(self.helper, 'apply_conditional_formatting'):
                try:
                    self.helper.apply_conditional_formatting(converted_df, data_start_row, start_col=0)
                except Exception as e:
                    self.logger.debug(f"Could not apply conditional formatting: {e}")
            
            # Add visual separation after the performance data table
            table_end_row = data_start_row + len(converted_df)
            
            # Insert a few empty rows for visual separation
            for i in range(3):
                self.worksheet.write(table_end_row + i, 0, "", self.formatter.get('cell'))
            
            # Add a colored separator row
            separator_row = table_end_row + 3
            for col in range(len(converted_df.columns)):  # Use actual column count instead of fixed 16
                self.worksheet.write(separator_row, col, "", self.formatter.get('motor_info_header'))
            
            return separator_row + 2
            
        except Exception as e:
            self.logger.error(f"Error adding performance data comparison: {e}")
            return start_row + 2

    def _convert_performance_data(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
        """Apply unit preferences when configuration is available, otherwise fall back to defaults."""
        if self.config:
            converted_df, metadata = apply_unit_preferences(df, self.config)
            return converted_df, metadata

        return self._perform_unit_conversions(df)

    def _perform_unit_conversions(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
        """
        Apply robust unit conversions with flexible column matching to ensure chart compatibility.
        Convert vacuum from mmH2O to kPa and air flow from l/sec to m³/h.
        Always adds converted columns even if source columns are not found (filled with NaN).
        Returns the converted DataFrame and unit metadata aligned with default units.
        """
        try:
            # Make a copy to avoid modifying the original
            converted_df = df.copy()
            
            # Debug: Log all available columns
            self.logger.info(f"Input columns for unit conversion: {list(converted_df.columns)}")
            
            # Initialize converted columns with NaN values first
            converted_df['Vacuum Corrected (kPa)'] = pd.NA
            converted_df['Air Flow (m³/h)'] = pd.NA
            
            # Vacuum conversion: mmH2O to kPa (1 mmH2O = 0.00980665 kPa)
            vacuum_col_mmh2o = None
            vacuum_variations = [
                'Vacuum Corrected mmH2O',
                'Vacuum mmH2O', 
                'Pressione Finale vuoto (mmH2O)',
                'Vacuum Corrected mmH2O',  # Exact match
                'Vacuum mmH2O',            # Without "Corrected"
                'Vacuum Corrected (mmH2O)', # With parentheses
                'Vacuum (mmH2O)',          # Simplified
            ]
            
            # Try exact matches first, then flexible matching
            for col in converted_df.columns:
                if col in vacuum_variations:
                    vacuum_col_mmh2o = col
                    break
                # Flexible matching: contains key terms
                col_lower = col.lower()
                if ('vacuum' in col_lower and 'mmh2o' in col_lower) or \
                   ('pressione' in col_lower and 'vuoto' in col_lower and 'mmh2o' in col_lower):
                    vacuum_col_mmh2o = col
                    break

            if vacuum_col_mmh2o:
                if pd.api.types.is_numeric_dtype(converted_df[vacuum_col_mmh2o]):
                    converted_df['Vacuum Corrected (kPa)'] = converted_df[vacuum_col_mmh2o] * 0.00980665
                    self.logger.info(f"✅ Converted '{vacuum_col_mmh2o}' to 'Vacuum Corrected (kPa)'")
                    self.logger.debug(f"Sample values: {converted_df[vacuum_col_mmh2o].head(3).tolist()} → {converted_df['Vacuum Corrected (kPa)'].head(3).tolist()}")
                else:
                    self.logger.warning(f"❌ Vacuum column '{vacuum_col_mmh2o}' found but is not numeric, keeping NaN values")
            else:
                self.logger.warning(f"❌ No suitable vacuum column (mmH2O) found for kPa conversion, keeping NaN values")
            
            # Air Flow conversion: l/sec to m³/h (1 l/sec = 3.6 m³/h)
            airflow_col_lsec = None
            airflow_variations = [
                'Air Flow l/sec.',
                'Air Flow l/sec',  # Without period
                'Portata Finale Aria (l/s)',
                'Air Flow (l/sec)',
                'Air Flow (l/s)',
                'Airflow l/sec',
                'Airflow (l/sec)',
            ]
            
            # Try exact matches first, then flexible matching
            for col in converted_df.columns:
                if col in airflow_variations:
                    airflow_col_lsec = col
                    break
                # Flexible matching: contains key terms
                col_lower = col.lower()
                if ('air' in col_lower or 'portata' in col_lower) and \
                   ('l/sec' in col_lower or 'l/s' in col_lower or 'liter' in col_lower):
                    airflow_col_lsec = col
                    break
            
            if airflow_col_lsec:
                if pd.api.types.is_numeric_dtype(converted_df[airflow_col_lsec]):
                    converted_df['Air Flow (m³/h)'] = converted_df[airflow_col_lsec] * 3.6
                    self.logger.info(f"✅ Converted '{airflow_col_lsec}' to 'Air Flow (m³/h)'")
                    self.logger.debug(f"Sample values: {converted_df[airflow_col_lsec].head(3).tolist()} → {converted_df['Air Flow (m³/h)'].head(3).tolist()}")
                else:
                    self.logger.warning(f"❌ Air flow column '{airflow_col_lsec}' found but is not numeric, keeping NaN values")
            else:
                self.logger.warning(f"❌ No suitable air flow column (l/sec) found for m³/h conversion, keeping NaN values")
            
            # Debug: Log final columns after conversion
            self.logger.info(f"Output columns after unit conversion: {list(converted_df.columns)}")
            
            # Verify expected chart columns exist
            chart_columns = ['Air Flow (m³/h)', 'Vacuum Corrected (kPa)', 'Efficiency (%)', 'Power Corrected Watts in', 'Speed RPM']
            missing_columns = [col for col in chart_columns if col not in converted_df.columns]
            available_columns = [col for col in chart_columns if col in converted_df.columns]
            
            if available_columns:
                self.logger.info(f"✅ Chart columns available: {available_columns}")
            if missing_columns:
                self.logger.warning(f"❌ Chart columns missing: {missing_columns}")
                
            # Log the count of non-null values in converted columns
            vacuum_count = converted_df['Vacuum Corrected (kPa)'].notna().sum()
            airflow_count = converted_df['Air Flow (m³/h)'].notna().sum()
            self.logger.info(f"Converted column data counts: Vacuum={vacuum_count}/{len(converted_df)}, Airflow={airflow_count}/{len(converted_df)}")
            
            metadata = self._build_default_metadata()

            return converted_df, metadata
            
        except Exception as e:
            self.logger.error(f"Error performing unit conversions: {e}", exc_info=True)
            return df, self._build_default_metadata()  # Return original if conversion fails

    def _get_unit_info(self, measurement: str) -> Dict[str, str]:
        """Return unit metadata for the requested measurement, honoring selected preferences."""
        if self.unit_metadata and measurement in self.unit_metadata:
            return self.unit_metadata[measurement]

        settings = UNIT_CONFIG[measurement]['settings']

        if self.config:
            config_attr = f"{measurement}_unit"
            selected_unit = getattr(self.config, config_attr, None)
            if not selected_unit or selected_unit not in settings:
                selected_unit = next(iter(settings))
        else:
            selected_unit = next(iter(settings))

        info = settings[selected_unit]
        return {
            'column': info['label'],
            'axis_label': info['axis_label'],
            'chart_name': info['chart_name'],
            'unit': selected_unit,
        }

    def _build_default_metadata(self) -> Dict[str, Dict[str, str]]:
        """Metadata aligned with base units used when configuration data is unavailable."""
        return {
            'pressure': {
                'column': UNIT_CONFIG['pressure']['settings']['kPa']['label'],
                'axis_label': UNIT_CONFIG['pressure']['settings']['kPa']['axis_label'],
                'chart_name': UNIT_CONFIG['pressure']['settings']['kPa']['chart_name'],
                'unit': 'kPa',
            },
            'flow': {
                'column': UNIT_CONFIG['flow']['settings']['m³/h']['label'],
                'axis_label': UNIT_CONFIG['flow']['settings']['m³/h']['axis_label'],
                'chart_name': UNIT_CONFIG['flow']['settings']['m³/h']['chart_name'],
                'unit': 'm³/h',
            },
            'power': {
                'column': UNIT_CONFIG['power']['settings']['W']['label'],
                'axis_label': UNIT_CONFIG['power']['settings']['W']['axis_label'],
                'chart_name': UNIT_CONFIG['power']['settings']['W']['chart_name'],
                'unit': 'W',
            },
            'speed': {
                'column': UNIT_CONFIG['speed']['settings']['rpm']['label'],
                'axis_label': UNIT_CONFIG['speed']['settings']['rpm']['axis_label'],
                'chart_name': UNIT_CONFIG['speed']['settings']['rpm']['chart_name'],
                'unit': 'rpm',
            },
        }

    def _log_action(self, action, level="info", extra_info=None):
        """
        Standardized logging method to replace excessive emoji logging.
        
        Args:
            action: The action being performed
            level: Log level ("info", "debug", "warning", "error")
            extra_info: Optional additional information
        """
        message = f"ComparisonSheet - {action}"
        if extra_info:
            message += f": {extra_info}"
        
        log_method = getattr(self.logger, level, self.logger.info)
        log_method(message)

    def _col_idx_to_excel(self, idx: int) -> str:
        """Convert zero-based column index to Excel column letters (e.g. 0 -> 'A', 27 -> 'AB')."""
        # Simpler robust approach: convert using 1-based math
        n = idx + 1
        letters = ''
        while n > 0:
            n, remainder = divmod(n-1, 26)
            letters = chr(65 + remainder) + letters
        return letters

    def _range_list_to_a1(self, range_list: List) -> Optional[str]:
        """Convert xlsxwriter list-format range to A1 string with sheet name.

        Expects: [sheet_name, first_row, first_col, last_row, last_col] with zero-based indices.
        Returns string like "Sheet1!$A$1:$C$10" or None on error.
        """
        try:
            sheet, r1, c1, r2, c2 = range_list
            # Convert to 1-based rows
            r1_1 = int(r1) + 1
            r2_1 = int(r2) + 1
            c1_a = self._col_idx_to_excel(int(c1))
            c2_a = self._col_idx_to_excel(int(c2))
            # Escape sheet name if it has spaces or special chars
            sheet_str = str(sheet)
            if ' ' in sheet_str or any(ch in sheet_str for ch in "'![]:/\\"):
                sheet_name = f"'{sheet_str}'"
            else:
                sheet_name = sheet_str
            return f"{sheet_name}!${c1_a}${r1_1}:${c2_a}${r2_1}"
        except Exception:
            return None
