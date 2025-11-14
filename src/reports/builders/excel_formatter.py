import logging
from typing import List, Dict, Any, Optional

from xlsxwriter.workbook import Workbook as XlsxWorkbook
from xlsxwriter.format import Format as XlsxFormat

logger = logging.getLogger(__name__)

class ExcelFormatter:
    """Manages Excel formats for the workbook."""
    def __init__(self, workbook: XlsxWorkbook, logo_colors: Optional[List[str]] = None): # Added logo_colors
        self.wb = workbook
        self.logo_colors = logo_colors if logo_colors else []
        self.formats: Dict[str, XlsxFormat] = {}
        self._create_formats()

    def _get_contrasting_font_color(self, hex_bg_color: str) -> str:
        """Determines if black or white font has better contrast against hex_bg_color."""
        if not hex_bg_color.startswith('#') or len(hex_bg_color) != 7:
            return '#000000' # Default to black for invalid colors
        try:
            r = int(hex_bg_color[1:3], 16)
            g = int(hex_bg_color[3:5], 16)
            b = int(hex_bg_color[5:7], 16)
            # Luminance formula (simplified)
            luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
            return '#000000' if luminance > 0.5 else '#FFFFFF' # Black on light, White on dark
        except ValueError:
            return '#000000'

    def _create_formats(self):
        # Standard formats
        self.formats['bold'] = self.wb.add_format({'bold': True})
        self.formats['cell'] = self.wb.add_format({'border': 1})
        self.formats['percent'] = self.wb.add_format({'num_format': '0.00%', 'border': 1})
        self.formats['integer'] = self.wb.add_format({'num_format': '0', 'border': 1})
        self.formats['decimal_2'] = self.wb.add_format({'num_format': '0.00', 'border': 1})
        self.formats['info_label'] = self.wb.add_format({'italic': True, 'font_color': 'gray'})
        self.formats['noise_label'] = self.wb.add_format({'bold': True, 'align': 'center'})
        self.formats['red_highlight'] = self.wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        self.formats['green_highlight'] = self.wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        self.formats['text_left_border'] = self.wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})
          # Link format
        self.formats['cell_link'] = self.wb.add_format({'font_color': 'blue', 'underline': 1, 'border': 1})        # Test divider formats - thick bottom border to separate tests with proper data formatting
        self.formats['test_divider'] = self.wb.add_format({
            'border': 1, 
            'bottom': 6, 
            'bottom_color': '#4F81BD',  # Blue color for the thick border
            'bg_color': '#F0F8FF'       # Light blue background to make it more visible
        })
        
        self.formats['test_divider_decimal'] = self.wb.add_format({
            'num_format': '0.00', 'border': 1,
            'bottom': 6, 'bottom_color': '#4F81BD', 'bg_color': '#F0F8FF'
        })
        
        self.formats['test_divider_integer'] = self.wb.add_format({
            'num_format': '0', 'border': 1,
            'bottom': 6, 'bottom_color': '#4F81BD', 'bg_color': '#F0F8FF'
        })
        
        self.formats['test_divider_percent'] = self.wb.add_format({
            'num_format': '0.00%', 'border': 1,
            'bottom': 6, 'bottom_color': '#4F81BD', 'bg_color': '#F0F8FF'
        })
        
        self.formats['test_divider_cell'] = self.wb.add_format({
            'border': 1,
            'bottom': 6, 'bottom_color': '#4F81BD', 'bg_color': '#F0F8FF'
        })

        # Logo-color based formats
        primary_logo_color = self.logo_colors[0] if self.logo_colors else '#4F81BD' # Default to a medium blue
        secondary_logo_color = self.logo_colors[1] if len(self.logo_colors) > 1 else '#D9D9D9' # Default to light grey

        title_font_color = self._get_contrasting_font_color(primary_logo_color)
        self.formats['report_title'] = self.wb.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
            'fg_color': primary_logo_color, 'font_color': title_font_color, 'border': 1
        })

        header_font_color = self._get_contrasting_font_color(secondary_logo_color)
        self.formats['header'] = self.wb.add_format({ # Overwrites previous 'header'
            'bold': True, 'align': 'center', 'valign': 'vcenter', 
            'fg_color': secondary_logo_color, 'font_color': header_font_color, 'border': 1
        })
        
        self.formats['motor_info_header'] = self.wb.add_format({ # Use secondary color for this too
            'bold': True, 'bg_color': secondary_logo_color, 'font_color': header_font_color, 
            'border':1, 'align':'left'
        })
        self.formats['motor_info_value'] = self.wb.add_format({'border':1, 'align':'left'})

        # Alternating row formats for better visual grouping in comparison sheets
        alt_bg_color = '#F5F5F5'  # Light gray for alternating rows
        self.formats['cell_alt'] = self.wb.add_format({'border': 1, 'bg_color': alt_bg_color})
        self.formats['percent_alt'] = self.wb.add_format({'num_format': '0.00%', 'border': 1, 'bg_color': alt_bg_color})
        self.formats['integer_alt'] = self.wb.add_format({'num_format': '0', 'border': 1, 'bg_color': alt_bg_color})
        self.formats['decimal_2_alt'] = self.wb.add_format({'num_format': '0.00', 'border': 1, 'bg_color': alt_bg_color})


    def get(self, name: str) -> Optional[XlsxFormat]:
        return self.formats.get(name)
