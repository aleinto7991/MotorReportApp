"""
Life Test (LF) Registry Reader

This module handles reading and processing the Life Test registry Excel file.
It provides functionality to search for LF tests by SAP code and locate the
corresponding test files.

Author: Motor Test Analysis Team
Version: 1.0
"""

import logging
import re
from pathlib import Path
from typing import List, Optional, Dict
import pandas as pd

# Relative imports (correct pattern for subpackage)
from ..data.models import LifeTestInfo

# Import directory configuration with fallback
try:
    from ..directory_config import LF_REGISTRY_FILE as DEFAULT_LF_REGISTRY
    from ..directory_config import LF_BASE_DIR as DEFAULT_LF_DIR
except ImportError:
    # Fallback to hardcoded paths if import fails
    DEFAULT_LF_REGISTRY = r"C:\Users\aintorbida\OneDrive - AMETEK Inc\ENG & Quality\UTE_wrk\LF\RELIABIL\REGISTRO LF .xlsx"
    DEFAULT_LF_DIR = r"C:\Users\aintorbida\OneDrive - AMETEK Inc\ENG & Quality\UTE_wrk\LF\RELIABIL"


class LifeTestRegistryReader:
    """Reader for Life Test registry and file locator."""
    
    def __init__(self, registry_path: Optional[str] = None, lf_base_dir: Optional[str] = None):
        """
        Initialize the LF registry reader.
        
        Args:
            registry_path: Path to the LF registry Excel file
            lf_base_dir: Base directory containing year folders with LF test files
        """
        self.logger = logging.getLogger(__name__)
        
        # Use provided paths, fallback to directory_config, then hardcoded defaults
        if registry_path:
            self.registry_path = Path(registry_path)
        elif DEFAULT_LF_REGISTRY:
            self.registry_path = Path(DEFAULT_LF_REGISTRY)
            self.logger.info(f"Using LF registry from directory_config: {self.registry_path}")
        else:
            # Final fallback
            self.registry_path = Path(r"C:\Users\aintorbida\OneDrive - AMETEK Inc\ENG & Quality\UTE_wrk\LF\RELIABIL\REGISTRO LF .xlsx")
            self.logger.warning(f"Using hardcoded fallback LF registry path: {self.registry_path}")
        
        if lf_base_dir:
            self.lf_base_dir = Path(lf_base_dir)
        elif DEFAULT_LF_DIR:
            self.lf_base_dir = Path(DEFAULT_LF_DIR)
            self.logger.info(f"Using LF base dir from directory_config: {self.lf_base_dir}")
        else:
            # Final fallback
            self.lf_base_dir = Path(r"C:\Users\aintorbida\OneDrive - AMETEK Inc\ENG & Quality\UTE_wrk\LF\RELIABIL")
            self.logger.warning(f"Using hardcoded fallback LF base dir: {self.lf_base_dir}")
        
        self.registry_df: Optional[pd.DataFrame] = None
        # Create file indexer for quick LF file lookups (build in background)
        try:
            from .lf_indexer import LFIndex
            # Start background indexing to keep lookups fast as data grows
            self.indexer = LFIndex(self.lf_base_dir, background=True, logger=self.logger)
        except Exception as e:
            self.indexer = None
            self.logger.debug(f"LF indexer not available: {e}")
        
    def load_registry(self) -> bool:
        """
        Load the LF registry from Excel file.
        
        Returns:
            True if loaded successfully, False otherwise
        """
        import time
        
        max_retries = 3
        retry_delay = 0.5  # seconds
        
        for attempt in range(max_retries):
            try:
                if not self.registry_path.exists():
                    self.logger.error(f"LF registry not found at: {self.registry_path}")
                    return False
                
                # Load with header at row 2 (0-indexed)
                self.registry_df = pd.read_excel(
                    self.registry_path,
                    sheet_name='Registro',
                    header=2
                )
                
                self.logger.info(f"✅ Loaded LF registry with {len(self.registry_df)} records")
                return True
                
            except PermissionError as e:
                if attempt < max_retries - 1:
                    self.logger.warning(f"Permission denied reading LF registry (attempt {attempt + 1}/{max_retries}). Retrying...")
                    time.sleep(retry_delay)
                else:
                    self.logger.error(f"Error loading LF registry after {max_retries} attempts: {e}")
                    self.logger.error("Please close the file if it's open in Excel and try again.")
                    return False
            except Exception as e:
                self.logger.error(f"Error loading LF registry: {e}")
                return False
        
        return False
    
    def parse_test_number(self, test_number: str) -> tuple:
        """
        Robust parsing of LF test numbers to extract test ID and year.

        Handles common variants encountered in the registry such as:
        - LF053/07, LF001-25
        - LF003_BIS/19 (suffixes between id and year)
        - LFO56-23 (unexpected letters between 'LF' and digits)

        Returns (test_id, year) or (None, None) on failure.
        """
        if not test_number or not isinstance(test_number, str):
            return (None, None)

        s = test_number.strip()
        try:
            # Primary (most common) pattern: LF<digits>[/-]<2-digit-year>
            m = re.search(r'LF\s*(\d{1,4})[/-](\d{2})', s, flags=re.IGNORECASE)
            if not m:
                # Fallback 1: any digits + separator + 2-digit year
                m = re.search(r'(\d{1,4})[\s\-_A-Za-z]*[/-][\s]*?(\d{2})', s)
            if not m:
                # Fallback 2: last occurrence of pattern digits/yy anywhere
                m = re.search(r'(\d{1,4})[/-](\d{2})', s)

            if m:
                test_id = m.group(1).zfill(3)
                year_short = m.group(2)
                year_int = int(year_short)
                if year_int >= 97:
                    year = f"19{year_short}"
                else:
                    year = f"20{year_short}"
                return (test_id, year)
            else:
                self.logger.debug(f"Could not parse LF test number: '{s}'")
        except Exception as e:
            self.logger.debug(f"Error parsing test number {test_number}: {e}")

        return (None, None)
    
    def find_test_file(self, test_id: str, year: str) -> Optional[Path]:
        """
        Find the LF test file in the year folder. If the file is not found using
        common filename patterns, perform a limited recursive search with
        heuristics to locate likely candidates.

        Args:
            test_id: Test ID (e.g., "053")
            year: Year as 4-digit string (e.g., "2007")
            
        Returns:
            Path to the test file if found, None otherwise
        """
        try:
            if not test_id:
                return None

            year_folder = self.lf_base_dir / year
            if year_folder.exists():
                # Common exact filename patterns
                possible_names = [
                    f"LF {test_id}.xlsx",
                    f"LF{test_id}.xlsx",
                    f"LF {test_id}.xls",
                    f"LF{test_id}.xls",
                    f"LF {test_id}-*.xls",
                    f"LF{test_id}-*.xls",
                    f"*{test_id}*.xls*",
                ]

                for pattern in possible_names:
                    # Use glob to allow patterns with wildcards
                    for p in year_folder.glob(pattern):
                        if p.is_file():
                            self.logger.debug(f"Found LF file (pattern match): {p}")
                            return p

                # If still not found, do a recursive search within the year folder
                try:
                    pattern = f"**/*{test_id}*.xls*"
                    candidates = list(year_folder.glob(pattern))
                    if candidates:
                        # Prefer filenames that contain the letters 'LF' and the test id
                        for p in candidates:
                            name_upper = p.name.upper()
                            if 'LF' in name_upper and test_id in name_upper:
                                self.logger.debug(f"Found LF file (year recursive): {p}")
                                return p
                        # Fallback to first candidate
                        self.logger.debug(f"Found LF file (year recursive fallback): {candidates[0]}")
                        return candidates[0]
                except Exception:
                    self.logger.debug("Year-folder recursive search failed, continuing to base-dir search")
            else:
                self.logger.debug(f"Year folder not found: {year_folder}")

            # Final fallback: search the entire LF base directory for likely matches
            try:
                pattern = f"**/*{test_id}*.xls*"
                for p in self.lf_base_dir.glob(pattern):
                    if not p.is_file():
                        continue
                    name_upper = p.name.upper()
                    # Prefer files that clearly include 'LF' and the test id
                    if 'LF' in name_upper and test_id in name_upper:
                        self.logger.debug(f"Found LF file (base recursive): {p}")
                        return p
                    # Otherwise accept files where the stem ends/starts with the id
                    stem = p.stem
                    if stem.endswith(test_id) or stem.startswith(test_id) or re.search(r'\b' + re.escape(test_id) + r'\b', stem):
                        self.logger.debug(f"Found LF file (base recursive fallback): {p}")
                        return p
            except Exception as e:
                self.logger.error(f"Error during fallback search for LF {test_id}: {e}")

            self.logger.debug(f"No file found for LF {test_id} in {year} after fallback search")
            return None

        except Exception as e:
            self.logger.error(f"Error finding test file: {e}")
            return None
    
    def get_tests_for_sap(self, sap_code: str) -> List[LifeTestInfo]:
        """
        Get all LF tests for a given SAP code.
        
        Args:
            sap_code: SAP code to search for
            
        Returns:
            List of LifeTestInfo objects
        """
        if self.registry_df is None:
            if not self.load_registry():
                return []
        
        try:
            # Filter by SAP code (handle both string and numeric SAP codes)
            sap_str = str(sap_code).strip()
            filtered_df = self.registry_df[
                self.registry_df['Cod. SAP'].astype(str).str.strip() == sap_str
            ]
            
            if filtered_df.empty:
                self.logger.info(f"No LF tests found for SAP {sap_code}")
                return []
            
            # Convert to LifeTestInfo objects
            lf_tests = []
            for _, row in filtered_df.iterrows():
                test_number = str(row.get('N. test', '')).strip()
                if not test_number or test_number == 'nan':
                    continue
                
                # Parse test number to get year and ID
                test_id, year = self.parse_test_number(test_number)
                
                # Find the actual file
                file_path = None
                file_exists = False
                if test_id and year:
                    file_path = self.find_test_file(test_id, year)
                    file_exists = file_path is not None
                
                lf_test = LifeTestInfo(
                    test_number=test_number,
                    sap_code=sap_str,
                    notes=str(row.get('NOTE', '')).strip() if pd.notna(row.get('NOTE')) else None,
                    year=year,
                    test_id=test_id,
                    file_path=file_path,
                    file_exists=file_exists,
                    responsible=str(row.get('Resp.', '')).strip() if pd.notna(row.get('Resp.')) else None
                )
                lf_tests.append(lf_test)
            
            self.logger.info(f"✅ Found {len(lf_tests)} LF test(s) for SAP {sap_code}")
            return lf_tests
            
        except Exception as e:
            self.logger.error(f"Error getting LF tests for SAP {sap_code}: {e}")
            return []
    
    def get_test_by_number(self, test_number: str) -> Optional[LifeTestInfo]:
        """
        Get a specific LF test by its test number.
        
        Args:
            test_number: Test number (e.g., "LF053/07")
            
        Returns:
            LifeTestInfo object or None if not found
        """
        if self.registry_df is None:
            if not self.load_registry():
                return None
        
        try:
            test_number_clean = test_number.strip()
            row = self.registry_df[
                self.registry_df['N. test'].astype(str).str.strip() == test_number_clean
            ]
            
            if row.empty:
                self.logger.debug(f"Test {test_number} not found in registry")
                return None
            
            row = row.iloc[0]
            test_id, year = self.parse_test_number(test_number_clean)
            
            file_path = None
            file_exists = False
            if test_id and year:
                file_path = self.find_test_file(test_id, year)
                file_exists = file_path is not None
            
            return LifeTestInfo(
                test_number=test_number_clean,
                sap_code=str(row.get('Cod. SAP', '')).strip() if pd.notna(row.get('Cod. SAP')) else None,
                notes=str(row.get('NOTE', '')).strip() if pd.notna(row.get('NOTE')) else None,
                year=year,
                test_id=test_id,
                file_path=file_path,
                file_exists=file_exists,
                responsible=str(row.get('Resp.', '')).strip() if pd.notna(row.get('Resp.')) else None
            )
            
        except Exception as e:
            self.logger.error(f"Error getting test {test_number}: {e}")
            return None

    def reconcile_registry(self, sap_codes: Optional[List[str]] = None, sample_per_sap: Optional[int] = None) -> Dict[str, List[Dict]]:
        """Attempt to reconcile registry entries with files on disk.

        Returns a mapping of SAP -> list of result dicts with keys:
          - test_number, test_id, year, file_path (if found), file_exists (bool)
          - suggestions: list of suggested file paths when exact match missing

        This method is useful for auditing and for auto-suggestion when the
        registry and file names drift over time.
        """
        import difflib

        if self.registry_df is None:
            if not self.load_registry():
                return {}

        results_by_sap: Dict[str, List[Dict]] = {}
        df = self.registry_df

        # Determine SAPs to inspect
        unique_saps = df['Cod. SAP'].dropna().astype(str).str.strip().unique().tolist()
        if sap_codes:
            target_saps = [s for s in unique_saps if s in sap_codes]
        else:
            target_saps = unique_saps

        for sap in target_saps:
            sap_df = df[df['Cod. SAP'].astype(str).str.strip() == sap]
            tests = sap_df['N. test'].tolist()
            if sample_per_sap:
                tests = tests[:sample_per_sap]

            sap_results = []
            for t in tests:
                t_str = str(t).strip()
                test_id, year = self.parse_test_number(t_str)
                entry = {'test_number': t_str, 'test_id': test_id, 'year': year, 'file_path': None, 'file_exists': False, 'suggestions': []}

                # Exact lookup
                if test_id and year:
                    p = self.find_test_file(test_id, year)
                    if p:
                        entry['file_path'] = str(p)
                        entry['file_exists'] = True
                        sap_results.append(entry)
                        continue

                # Try index-based id-only suggestions
                suggestions = []
                if hasattr(self, 'indexer') and self.indexer is not None:
                    cand = self.indexer.get_best_file(test_id or t_str, None)
                    if cand:
                        suggestions.append({'reason': 'id_only', 'path': str(cand)})

                    # Fuzzy filename matching against indexed filenames
                    # Build a small lookup of candidate names -> path
                    all_entries = []
                    for id_key, lst in self.indexer._index.items():
                        for ce in lst:
                            all_entries.append((ce.get('name', ''), ce.get('path')))
                    names = [n for n, _ in all_entries]
                    if names:
                        close = difflib.get_close_matches(t_str, names, n=3, cutoff=0.6)
                        for cname in close:
                            for nm, pth in all_entries:
                                if nm == cname:
                                    suggestions.append({'reason': 'fuzzy_name', 'path': pth})
                                    break

                # Fallback: limited recursive search by transformed test string
                if not suggestions:
                    try:
                        pattern = f"**/*{t_str.replace('/', '_')}*.*"
                        found = list(self.lf_base_dir.glob(pattern))
                        for f in found[:3]:
                            suggestions.append({'reason': 'recursive_name', 'path': str(f)})
                    except Exception:
                        pass

                entry['suggestions'] = suggestions
                sap_results.append(entry)

            results_by_sap[sap] = sap_results

        return results_by_sap
