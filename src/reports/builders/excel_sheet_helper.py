import logging
from typing import Dict, Optional, List, Tuple

import pandas as pd
from xlsxwriter.worksheet import Worksheet as XlsxWorksheet
from xlsxwriter.workbook import Workbook as XlsxWorkbook
from xlsxwriter.chart import Chart as XlsxChart

from .excel_formatter import ExcelFormatter

logger = logging.getLogger(__name__)

class ExcelSheetHelper:
    """Helper class for common worksheet operations."""
    def __init__(self, worksheet: XlsxWorksheet, workbook: XlsxWorkbook, formatter: ExcelFormatter):
        self.ws = worksheet
        self.wb = workbook
        self.formatter = formatter
        self.logger = logging.getLogger(__class__.__name__)


    def set_column_widths(self, widths: Dict[str, float]):
        """Sets column widths from a dictionary of column ranges to widths."""
        for col_range, width in widths.items():
            self.ws.set_column(col_range, width)

    def write_data_frame(self, df: pd.DataFrame, start_row: int, start_col: int, include_header: bool = True):
        """Writes a pandas DataFrame to the worksheet."""
        if include_header:
            for c_idx, col_name in enumerate(df.columns):
                self.ws.write(start_row, start_col + c_idx, col_name, self.formatter.get('header'))
        
        for r_idx, row_data in enumerate(df.itertuples(index=False), start=1 if include_header else 0):
            for c_idx, value in enumerate(row_data):
                self.ws.write(start_row + r_idx, start_col + c_idx, value, self.formatter.get('cell'))

    def insert_image(self, row: int, col: int, image_path: str, options: Optional[Dict] = None):
        """Inserts an image into the worksheet."""
        if not options:
            options = {'x_scale': 0.5, 'y_scale': 0.5}
        self.ws.insert_image(row, col, image_path, options)

    def auto_fit_columns(self, df: pd.DataFrame, start_col: int = 0, max_width: int = 25) -> None:
        """Auto-fit column widths based on DataFrame content and header, with a max width."""
        if df is None or df.empty: return
        for i, col_name in enumerate(df.columns):
            try:
                header_len = len(str(col_name))
                max_data_len = df[col_name].astype(str).map(len).max() if not df[col_name].empty else 0
                width = max(header_len, int(max_data_len)) + 2 # Add a little padding
                final_width = min(max(width, 10), max_width) # Clamp between 10 and max_width
                self.ws.set_column(start_col + i, start_col + i, final_width)
            except Exception as e:
                self.logger.warning(f"Error auto-fitting column {col_name}: {e}. Setting default width.")
                self.ws.set_column(start_col + i, start_col + i, 15)

    def apply_conditional_formatting(self, df: pd.DataFrame, data_start_row: int, start_col: int) -> None:
        """Apply conditional formatting for performance columns."""
        if df is None or df.empty or len(df) <=1: return
        
        performance_keywords = ['efficiency', 'power', 'vacuum', 'speed', 'air flow', 'pressione', 'portata', 'current', 'torque', 'rpm', 'watts']
        exclude_keywords = ['orifice', 'costante', 'constant', 'setting', 'target', 'diam']

        for i, col_name in enumerate(df.columns):
            col_lower = col_name.lower()
            is_performance = (
                any(keyword in col_lower for keyword in performance_keywords) and
                not any(keyword in col_lower for keyword in exclude_keywords) and
                pd.api.types.is_numeric_dtype(df[col_name])
            )
            if is_performance:
                first_row_idx = data_start_row
                last_row_idx = data_start_row + len(df) - 1
                col_idx = start_col + i
                
                try:
                    self.ws.conditional_format(first_row_idx, col_idx, last_row_idx, col_idx,
                                            {'type': 'top', 'value': 1, 'format': self.formatter.get('green_highlight')})
                    self.ws.conditional_format(first_row_idx, col_idx, last_row_idx, col_idx,
                                            {'type': 'bottom', 'value': 1, 'format': self.formatter.get('red_highlight')})
                except Exception as e:
                     self.logger.warning(f"Error applying conditional format to col {col_name}: {e}")

    def create_chart(self, title: str, x_axis_label: str, y_axis_label: str, 
                     chart_type: str = 'scatter', subtype: str = 'smooth') -> Optional[XlsxChart]:
        """
        Create a chart with standard styling.
        
        Args:
            title: Chart title
            x_axis_label: Label for X-axis
            y_axis_label: Label for Y-axis
            chart_type: Chart type (default: 'scatter')
            subtype: Chart subtype (default: 'smooth')
            
        Returns:
            Chart object or None if creation failed
        """
        chart = self.wb.add_chart({'type': chart_type, 'subtype': subtype})
        if not chart:
            self.logger.error(f"Failed to create chart: {title}")
            return None
        
        # Set title
        chart.set_title({'name': title, 'name_font': {'size': 12, 'bold': True}})
        
        # Set X-axis with standard styling
        chart.set_x_axis({
            'name': x_axis_label,
            'name_font': {'size': 10, 'bold': True},
            'num_font': {'size': 9, 'bold': False},
            'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}},
            'major_tick_mark': 'outside',
            'line': {'color': '#595959', 'width': 1.0}
        })
        
        # Set Y-axis with standard styling
        chart.set_y_axis({
            'name': y_axis_label,
            'name_font': {'size': 10, 'bold': True},
            'num_font': {'size': 9, 'bold': False},
            'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}},
            'major_tick_mark': 'outside',
            'line': {'color': '#595959', 'width': 1.0}
        })
        
        # Set legend position
        chart.set_legend({'position': 'bottom'})
        
        return chart

    def add_chart_series(self, chart: XlsxChart, series_name: str, 
                        x_range: List, y_range: List, 
                        color: Optional[str] = None,
                        marker_size: int = 5, line_width: float = 2.0):
        """
        Add a data series to a chart.
        
        Args:
            chart: Chart object to add series to
            series_name: Name of the series for legend
            x_range: X-axis data range [sheet_name, row_start, col_start, row_end, col_end]
            y_range: Y-axis data range [sheet_name, row_start, col_start, row_end, col_end]
            color: Optional color hex code (e.g., '#4472C4')
            marker_size: Size of marker points (default: 5)
            line_width: Width of line (default: 2.0)
        """
        # Validate ranges
        if not x_range or not y_range:
            self.logger.warning(f"Skipping series '{series_name}': missing x_range or y_range")
            return

        # Determine chart type if possible and set series keys accordingly.
        # For scatter charts XlsxWriter expects 'x_values' and 'y_values'.
        # For line/column charts it expects 'categories' and 'values'.
        chart_type = None
        try:
            chart_type = getattr(chart, 'type', None) or getattr(chart, 'chart_type', None)
        except Exception:
            chart_type = None

        is_scatter = False
        if isinstance(chart_type, str) and chart_type.lower() == 'scatter':
            is_scatter = True

        series_config = {'name': series_name, 'marker': {'type': 'circle', 'size': marker_size}, 'line': {'width': line_width}}

        if is_scatter:
            series_config['x_values'] = x_range
            series_config['y_values'] = y_range
        else:
            series_config['categories'] = x_range
            series_config['values'] = y_range
        
        # Add color if specified
        if color:
            series_config['marker']['fill'] = {'color': color}
            series_config['marker']['border'] = {'color': color}
            series_config['line']['color'] = color
        
        try:
            chart.add_series(series_config)
        except Exception as e:
            self.logger.error(f"Failed to add series '{series_name}' to chart: {e}")
            raise

    def insert_chart_with_size(self, chart: XlsxChart, row: int, col: int, 
                               width: int = 680, height: int = 400):
        """
        Insert a chart into the worksheet with specified size.
        
        Args:
            chart: Chart object to insert
            row: Row index where chart should be inserted
            col: Column index where chart should be inserted
            width: Chart width in pixels (default: 680)
            height: Chart height in pixels (default: 400)
        """
        chart.set_size({'width': width, 'height': height})
        self.ws.insert_chart(row, col, chart)
