import logging
from typing import Dict, Optional

import pandas as pd
from xlsxwriter.worksheet import Worksheet as XlsxWorksheet
from xlsxwriter.workbook import Workbook as XlsxWorkbook

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
