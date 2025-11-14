"""
Noise Test Validator - integrates with existing Motor Report Generator architecture.

This module validates noise test files referenced in the registry and provides
a pre-check system to let users select which tests to include in reports.
"""

import os
import logging
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

import pandas as pd

from ..services.noise_registry_reader import (
    load_registry_dataframe,
    N_PROVA_STD,
    CODICE_SAP_STD,
    ANNO_ORIGINAL_STD,
    TEST_LAB_STD,
    NOTE_STD,
    VOLTAGE_STD,
    CLIENT_STD,
    APPLICATION_STD,
    RESPONSIBLE_STD,
)

logger = logging.getLogger(__name__)

@dataclass
class NoiseTestValidationInfo:
    """Information about a noise test file validation."""
    sap_code: str
    test_no: str
    file_path: str
    exists: bool
    is_valid: bool
    error_message: Optional[str] = None
    file_size: Optional[int] = None
    date: Optional[str] = None
    registry_row: Optional[int] = None  # Track which registry row this came from
    voltage: Optional[str] = None  # TENSIONE from registry
    client: Optional[str] = None  # CLIENTE from registry
    application: Optional[str] = None  # APPARECCHIO from registry  
    notes: Optional[str] = None  # NOTE from registry
    responsible: Optional[str] = None  # RESP. from registry
    test_lab: Optional[str] = None  # TEST LAB from registry

class NoiseTestValidator:
    """Validates noise test files from the registro rumore."""
    
    def __init__(self, test_folder: str, noise_registry_sheet_name: str = "Registro"):
        self.test_folder = test_folder
        self.noise_registry_sheet_name = noise_registry_sheet_name
        
        # First, check if the provided folder is already a noise folder (contains noise files directly)
        self.active_noise_folder = None
        
        # Check if this folder is already a noise folder by looking for common noise file patterns
        if os.path.exists(test_folder):
            # Look for CSV files that might be noise test files
            potential_files = []
            try:
                for file in os.listdir(test_folder):
                    if file.lower().endswith('.csv'):
                        potential_files.append(file)
            except (OSError, PermissionError):
                pass
            
            # If we find CSV files, assume this is already a noise folder
            if potential_files:
                self.active_noise_folder = test_folder
                logger.info(f"Using provided folder as noise folder: {test_folder}")
            else:
                # Look for noise folder in standard locations within the provided folder
                self.noise_folders = [
                    os.path.join(test_folder, "RUMORE"),
                    os.path.join(test_folder, "Tests Rumore"),
                    os.path.join(test_folder, "NOISE"),
                    os.path.join(test_folder, "noise")
                ]
                
                # Find the actual noise folder
                for folder in self.noise_folders:
                    if os.path.exists(folder):
                        self.active_noise_folder = folder
                        logger.info(f"Found noise folder: {folder}")
                        break
        
        if not self.active_noise_folder:
            logger.warning(f"No noise folder found in or under {test_folder}")
        
    def _load_registry(
        self,
        registry_path: str,
        *,
        nrows: Optional[int] = None,
    ) -> Tuple[pd.DataFrame, Dict[str, str]]:
        """Load the noise registry using shared reader with sheet fallbacks."""

        sheet_candidates = [
            candidate
            for candidate in [
                self.noise_registry_sheet_name,
                "Registro",
                "REGISTRO RUMORE",
                "Registro Rumore",
            ]
            if candidate
        ]
        ordered_candidates = list(dict.fromkeys(sheet_candidates))

        last_error: Optional[Exception] = None
        for sheet in ordered_candidates:
            try:
                df, mapping = load_registry_dataframe(
                    registry_path,
                    sheet_name=sheet,
                    log=logger,
                    nrows=nrows,
                )
                logger.info("Successfully read noise registry sheet: %s", sheet)
                return df, mapping
            except ValueError as exc:
                logger.debug("Sheet '%s' not found in noise registry: %s", sheet, exc)
                last_error = exc
            except FileNotFoundError:
                raise

        if last_error:
            logger.debug("Falling back to first sheet due to: %s", last_error)

        df, mapping = load_registry_dataframe(
            registry_path,
            sheet_name=0,
            log=logger,
            nrows=nrows,
        )
        logger.info("Using first sheet from noise registry workbook")
        return df, mapping

    def validate_from_registry(self, registry_path: str, max_rows: int = 5000) -> List[NoiseTestValidationInfo]:
        """
        Read noise tests from registry entries - optimized for performance with row limits.
        File existence will be checked later during actual report generation.
        
        Args:
            registry_path: Path to the registry Excel file
            max_rows: Maximum number of data rows to process (default: 5000 for performance)
        """
        noise_tests: List[NoiseTestValidationInfo] = []

        try:
            registry_df, column_mapping = self._load_registry(
                registry_path,
                nrows=max_rows + 50,
            )
        except FileNotFoundError as exc:
            logger.error("Noise registry file not found: %s", exc)
            raise
        except Exception as exc:
            logger.error("Error loading noise registry: %s", exc)
            raise

        if registry_df.empty:
            logger.warning("Noise registry returned no data after cleaning")
            return noise_tests

        if len(registry_df) > max_rows:
            logger.info(
                "Large registry detected (%s rows), limiting to first %s entries for performance",
                len(registry_df),
                max_rows,
            )
            registry_df = registry_df.head(max_rows)

        registry_df = registry_df.reset_index(drop=True)

        if column_mapping:
            logger.info("Registry column mapping: %s", column_mapping)
        logger.debug("Registry columns after normalisation: %s", list(registry_df.columns))
        if not registry_df.empty:
            logger.debug("Sample row data: %s", dict(registry_df.iloc[0]))

        processed_count = 0
        for row_number, (_, row) in enumerate(registry_df.iterrows()):
            if processed_count and processed_count % 1000 == 0:
                logger.info("Processed %s rows...", processed_count)
            processed_count += 1

            sap_code = self._safe_get_value_with_fallback(row, CODICE_SAP_STD, row_number, registry_df)
            test_no = self._safe_get_value_with_fallback(row, N_PROVA_STD, row_number, registry_df)
            date = self._safe_get_value_with_fallback(row, ANNO_ORIGINAL_STD, row_number, registry_df)
            voltage = self._safe_get_value_with_fallback(row, VOLTAGE_STD, row_number, registry_df)
            client = self._safe_get_value_with_fallback(row, CLIENT_STD, row_number, registry_df)
            application = self._safe_get_value_with_fallback(row, APPLICATION_STD, row_number, registry_df)
            notes = self._safe_get_value_with_fallback(row, NOTE_STD, row_number, registry_df)
            responsible = self._safe_get_value_with_fallback(row, RESPONSIBLE_STD, row_number, registry_df)
            test_lab = self._safe_get_value_with_fallback(row, TEST_LAB_STD, row_number, registry_df)

            if not sap_code and not test_no:
                continue

            logger.debug(
                "Row %s: SAP=%s, Test=%s, Date=%s, Voltage=%s, Client=%s, Notes=%s",
                row_number,
                sap_code,
                test_no,
                date,
                voltage,
                client,
                notes,
            )

            noise_tests.append(
                NoiseTestValidationInfo(
                    sap_code=sap_code or "Unknown",
                    test_no=test_no or "Unknown",
                    file_path="Will be determined during report generation",
                    exists=True,
                    is_valid=True,
                    error_message=None,
                    file_size=None,
                    date=date,
                    registry_row=row_number,
                    voltage=voltage or None,
                    client=client or None,
                    application=application or None,
                    notes=notes or None,
                    responsible=responsible or None,
                    test_lab=test_lab or None,
                )
            )

        logger.info("Found %s noise test entries in registry", len(noise_tests))
        return noise_tests
    
    def get_sap_codes_from_registry(self, registry_path: str, max_rows: int = 1000) -> List[str]:
        """
        Lightweight method to quickly extract just the SAP codes from the noise registry.
        This is used for pre-check validation in the UI - much faster than full validation.
        
        Args:
            registry_path: Path to the registry Excel file
            max_rows: Maximum number of data rows to process (default: 1000 for performance)
            
        Returns:
            List of unique SAP codes found in the registry
        """
        try:
            registry_df, _ = self._load_registry(
                registry_path,
                nrows=max_rows + 50,
            )
        except FileNotFoundError as exc:
            logger.error("Noise registry file not found: %s", exc)
            return []
        except Exception as exc:
            logger.error("Error reading SAP codes from noise registry: %s", exc)
            return []

        if registry_df.empty:
            return []

        if len(registry_df) > max_rows:
            logger.debug(
                "Large registry detected (%s rows), limiting to first %s entries for SAP code extraction",
                len(registry_df),
                max_rows,
            )
            registry_df = registry_df.head(max_rows)

        registry_df = registry_df.reset_index(drop=True)

        sap_codes: List[str] = []
        seen: set[str] = set()
        for _, row in registry_df.iterrows():
            sap_code = self._safe_get_value(row, CODICE_SAP_STD)
            if sap_code and sap_code.strip() and sap_code not in seen:
                seen.add(sap_code)
                sap_codes.append(sap_code.strip())

        logger.debug("Extracted %s unique SAP codes from noise registry", len(sap_codes))
        return sap_codes

    def _safe_get_value(self, row: pd.Series, column_name: str) -> str:
        """Safely get value from row, handling missing columns and NaN values."""
        if column_name in row.index:
            value = row[column_name]
            if pd.isna(value):
                return ''
            # Convert to string and handle float values properly
            str_value = str(value).strip()
            # Remove .0 from float values that are actually integers
            if str_value.endswith('.0') and str_value.replace('.0', '').replace('-', '').isdigit():
                str_value = str_value[:-2]
            return str_value
        return ''
    
    def _safe_get_value_with_fallback(self, row: pd.Series, column_name: str, row_idx: int, df: pd.DataFrame) -> str:
        """
        Safely get value from row, handling merged cells by checking adjacent rows.
        This handles cases where Excel merged cells cause data to appear in the row above/below.
        """
        # First try the normal approach
        value = self._safe_get_value(row, column_name)
        if value and value.strip() and value.strip().lower() != 'nan':
            return value
        
        # If empty, check the previous rows (up to 3 rows back for heavily merged areas)
        for back_offset in range(1, 4):
            if row_idx >= back_offset:
                prev_row = df.iloc[row_idx - back_offset]
                prev_value = self._safe_get_value(prev_row, column_name)
                if prev_value and prev_value.strip() and prev_value.strip().lower() != 'nan':
                    return prev_value
        
        # If still empty, check the next rows (up to 2 rows forward)
        for forward_offset in range(1, 3):
            if row_idx + forward_offset < len(df):
                next_row = df.iloc[row_idx + forward_offset]
                next_value = self._safe_get_value(next_row, column_name)
                if next_value and next_value.strip() and next_value.strip().lower() != 'nan':
                    return next_value
        
        return ''
    
    def _generate_possible_file_paths(self, sap_code: str, test_no: str, date: str) -> List[str]:
        """Generate possible file paths based on common naming conventions."""
        paths = []
        
        # Clean inputs
        sap_clean = sap_code.replace(' ', '_') if sap_code else 'unknown'
        test_clean = test_no.replace(' ', '_') if test_no else 'unknown'
        
        # Common naming patterns observed in the project
        patterns = [
            f"{test_clean}.csv",
            f"{sap_clean}_{test_clean}.csv",
            f"{test_clean}_{sap_clean}.csv",
            f"{sap_clean}.csv",
            f"noise_{test_clean}.csv",
            f"rumore_{test_clean}.csv",
            f"{test_clean}_noise.csv",
            f"{test_clean}_rumore.csv"
        ]
        
        # Add date-based patterns if date is available
        if date:
            date_clean = date.replace(' ', '_').replace('/', '_').replace('-', '_')
            patterns.extend([
                f"{test_clean}_{date_clean}.csv",
                f"{date_clean}_{test_clean}.csv"
            ])
        
        # Generate full paths
        if self.active_noise_folder:
            for pattern in patterns:
                full_path = os.path.join(self.active_noise_folder, pattern)
                paths.append(full_path)
            
        return paths
    
    def get_validation_summary(self, noise_tests: List[NoiseTestValidationInfo]) -> Dict[str, int]:
        """Get summary statistics of validation results."""
        summary = {
            'total': len(noise_tests),
            'valid': sum(1 for t in noise_tests if t.is_valid),
            'missing': sum(1 for t in noise_tests if not t.exists),
            'invalid': sum(1 for t in noise_tests if t.exists and not t.is_valid),
            'has_noise_folder': self.active_noise_folder is not None
        }
        return summary
