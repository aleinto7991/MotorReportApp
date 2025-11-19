"""Build Carichi Nominali sheets that mirror test-lab summaries per SAP (consolidated)."""
from __future__ import annotations

import logging
from typing import List, Tuple, Optional

from xlsxwriter.workbook import Workbook as XlsxWorkbook
from xlsxwriter.worksheet import Worksheet as XlsxWorksheet
from xlsxwriter.utility import xl_cell_to_rowcol

from ...data.models import MotorTestData, TestLabSummary
from .excel_formatter import ExcelFormatter

logger = logging.getLogger(__name__)


class CarichiSheetBuilder:
    """Render consolidated 'CARICHI NOMINALI' sheet inside the primary workbook."""

    def __init__(
        self,
        workbook: XlsxWorkbook,
        formatter: ExcelFormatter,
        logo_tab_colors: Optional[List[str]] = None,
        logo_path: Optional[str] = None,
    ) -> None:
        self.workbook = workbook
        self.formatter = formatter
        self.logo_tab_colors = logo_tab_colors or []
        self.logo_path = logo_path
        self.sheet: Optional[XlsxWorksheet] = None
        self.current_row = 0
        self._data: List[Tuple[str, List[MotorTestData]]] = []
        self._columns_set = False

        # Local formats
        self.section_title_fmt = self.workbook.add_format(
            {
                "bold": True,
                "font_size": 12,
                "bg_color": "#E8EAF6",
                "border": 1,
                "align": "left",
            }
        )
        self.table_cell_fmt = self.formatter.get("cell") or self.workbook.add_format({"border": 1})
        self.info_fmt = self.formatter.get("info_label") or self.workbook.add_format({"italic": True})

    def add_sap_data(self, sap_code: str, motor_tests: List[MotorTestData]) -> None:
        """Queue motor tests for a given SAP to be written in the consolidated sheet."""
        self._data.append((sap_code, motor_tests))

    def build(self) -> bool:
        """Create (or skip) the single 'CARICHI NOMINALI' worksheet with all queued data."""
        # Find any tests with test_lab_summary
        has_any = False
        for _, tests in self._data:
            if any(getattr(mt, "test_lab_summary", None) for mt in tests):
                has_any = True
                break
        if not has_any:
            logger.debug("No test-lab summaries for Carichi; skipping consolidated sheet.")
            return False

        # Create sheet
        try:
            self.sheet = self.workbook.add_worksheet("CARICHI NOMINALI")
        except Exception:
            # Fallback to sanitized name if necessary
            self.sheet = self.workbook.add_worksheet("CARICHI_NOMINALI")

        # Apply tab color if available
        if self.logo_tab_colors:
            try:
                self.sheet.set_tab_color(self.logo_tab_colors[min(1, len(self.logo_tab_colors) - 1)])
            except Exception:
                pass

        self._insert_logo()
        self._configure_print_settings()

        # Start writing after logo area
        self.current_row = 5

        for sap_code, tests in self._data:
            tests_with_summary = [mt for mt in tests if getattr(mt, "test_lab_summary", None)]
            if not tests_with_summary:
                continue

            # SAP header
            self.sheet.write(self.current_row, 0, f"SAP: {sap_code}", self.section_title_fmt)
            self.current_row += 1

            for mt in tests_with_summary:
                tls: TestLabSummary = mt.test_lab_summary  # type: ignore
                if not tls or not tls.raw_sheets:
                    continue

                # Test header
                self.sheet.write(self.current_row, 0, f"Test: {mt.test_number}", self.section_title_fmt)
                self.current_row += 1

                for raw_sheet in tls.raw_sheets:
                    # Source header
                    self.sheet.write(self.current_row, 0, f"Source: {raw_sheet.get('name', '')}", self.info_fmt)
                    self.current_row += 1

                    # Write raw sheet contents and get number of rows written
                    rows_written = self._write_raw_sheet(raw_sheet, self.current_row)
                    self.current_row += rows_written + 2

        logger.info("Created consolidated 'CARICHI NOMINALI' sheet")
        return True

    def _write_raw_sheet(self, raw_data: dict, start_row: int) -> int:
        """Write a raw extracted sheet block into the consolidated sheet starting at start_row.

        Returns the number of rows written.
        """
        assert self.sheet is not None

        # Set column widths once (avoid changing later)
        if not self._columns_set:
            for col_letter, width in raw_data.get("col_widths", {}).items():
                try:
                    self.sheet.set_column(f"{col_letter}:{col_letter}", width)
                except Exception:
                    # Ignore invalid column specification
                    pass
            self._columns_set = True

        # Set row heights (offsetting as described in source)
        for row_idx, height in raw_data.get("row_heights", {}).items():
            try:
                self.sheet.set_row(start_row + int(row_idx) - 1, height)
            except Exception:
                pass

        # Write values
        max_row_idx = -1
        for r_idx, row_data in enumerate(raw_data.get("values", [])):
            current_r = start_row + r_idx
            max_row_idx = max(max_row_idx, r_idx)
            for c_idx, value in enumerate(row_data):
                if value is not None:
                    try:
                        self.sheet.write(current_r, c_idx, value, self.table_cell_fmt)
                    except Exception:
                        # Best-effort writing - ignore cell-level write errors
                        pass

        # Apply merges (offset row coordinates)
        for merge_range in raw_data.get("merges", []):
            try:
                first_cell, last_cell = merge_range.split(":")
                r1, c1 = xl_cell_to_rowcol(first_cell)
                r2, c2 = xl_cell_to_rowcol(last_cell)
                new_r1 = start_row + r1
                new_r2 = start_row + r2

                # Extract a value for the merged area if present
                val = None
                if r1 < len(raw_data.get("values", [])):
                    row_vals = raw_data["values"][r1]
                    if c1 < len(row_vals):
                        val = row_vals[c1]

                if val is not None:
                    self.sheet.merge_range(new_r1, c1, new_r2, c2, val, self.table_cell_fmt)
                else:
                    self.sheet.merge_range(new_r1, c1, new_r2, c2, "", self.table_cell_fmt)
            except Exception as e:
                logger.debug("Merge apply failed for %s: %s", merge_range, e)

        return max_row_idx + 1 if max_row_idx >= 0 else 0

    def _insert_logo(self) -> None:
        """Insert the company logo at the top-left if present."""
        if not self.sheet or not self.logo_path:
            return

        try:
            self.sheet.insert_image(
                0, 0,
                self.logo_path,
                {
                    "x_scale": 0.18,
                    "y_scale": 0.18,
                    "object_position": 1,
                },
            )
        except Exception:
            logger.debug("Failed to insert logo into CARICHI NOMINALI sheet", exc_info=True)

    def _configure_print_settings(self) -> None:
        if not self.sheet:
            return
        try:
            self.sheet.set_paper(9)
            self.sheet.set_landscape()
            self.sheet.set_margins(left=0.3, right=0.3, top=0.5, bottom=0.5)
            self.sheet.fit_to_pages(1, 0)
        except Exception:
            logger.debug("Unable to set Carichi print settings", exc_info=True)
