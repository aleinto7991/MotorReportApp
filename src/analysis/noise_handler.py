"""
Noise Data Handler - Enhanced Version

This module handles noise test data from Excel registry files and locates
associated noise test images. It provides robust Excel parsing, data cleaning,
and comprehensive image finding capabilities.

Key Features:
- Auto-detection of Excel header rows
- Robust column name normalization
- Smart date/year extraction
- Comprehensive image search strategies
- Error handling and logging

Author: Motor Test Analysis Team
Version: 2.0
"""

import logging
from pathlib import Path
from typing import Optional, List, Dict, Any, Union

import pandas as pd

# Use proper relative imports
from ..config.app_config import AppConfig
from ..data.models import NoiseTestInfo
from ..utils.common import normalize_sap_code
from .noise_chart_generator import NoiseChartGenerator, NoiseTestData
from ..services.noise_registry_reader import (
    load_registry_dataframe,
    N_PROVA_STD,
    CODICE_SAP_STD,
    ANNO_ORIGINAL_STD,
    TEST_LAB_STD,
    NOTE_STD,
    ANNO_YYYY_STD,
    ANNO_DATETIME_STD,
)
from ..services.noise_directory_cache import get_noise_directory_cache


class NoiseDataHandler:
    """Handles noise test data from Excel registry and finds associated images."""
    
    def __init__(self, app_config: AppConfig):
        """Initialize the noise data handler."""
        self.app_config = app_config
        self.logger = logging.getLogger(__name__)
        self.registro_df: pd.DataFrame = pd.DataFrame()
        self.chart_generator = NoiseChartGenerator()  # Add chart generator
        self.directory_cache = get_noise_directory_cache()
        self._load_and_prepare_registro()

    def _load_and_prepare_registro(self):
        """Load and prepare the noise registry Excel file."""
        if not self.app_config.noise_registry_path:
            self.logger.error("Noise registry path is not configured. Cannot load noise data.")
            self.registro_df = pd.DataFrame()
            return

        excel_path = Path(self.app_config.noise_registry_path)
        
        if not excel_path.exists():
            self.logger.error(f"Noise registry file not found: {excel_path}")
            self.registro_df = pd.DataFrame()
            return
        
        try:
            sheet_name = self.app_config.noise_registry_sheet_name
            sheet_label = sheet_name if sheet_name else "Default (first sheet)"
            self.logger.info(f"Loading noise registry '{excel_path}' from sheet: {sheet_label}")

            self.registro_df, column_mapping = load_registry_dataframe(
                excel_path,
                sheet_name=sheet_name,
                log=self.logger,
            )

            if self.registro_df.empty:
                self.logger.warning("Noise registry contains no usable data after cleaning")
                return

            if column_mapping:
                self.logger.info(f"Registry column mapping applied: {column_mapping}")

            self.logger.info(
                "Registry loaded successfully: %s valid records, %s columns",
                len(self.registro_df),
                len(self.registro_df.columns),
            )

        except Exception as e:
            self.logger.error(f"Error loading registry file: {e}", exc_info=True)
            self.registro_df = pd.DataFrame()

    def get_noise_data_comprehensive(self, test_number: str, year: str, sap_code: Optional[str] = None) -> dict:
        """
        Enhanced noise data finder - optimized for performance when processing multiple tests.
        
        PERFORMANCE OPTIMIZATIONS:
        - Single directory scan with efficient file filtering
        - Lazy loading of TXT file processing
        - Memory-efficient file handling
        - Caching of directory structures
        
        Args:
            test_number: The noise test number (N.PROVA from registry)
            year: The year for the test
            sap_code: Optional SAP code (kept for compatibility, not used for filtering)
            
        Returns:
            Dict containing:
            - 'type': 'txt_data', 'images', or 'none'
            - 'txt_files': List of TXT file paths (if found)
            - 'images': List of image file paths (if found)
            - 'processed_data': Parsed noise measurement data (if TXT files)
        """
        self.logger.debug(f"🔍 Fast noise search: Test #{test_number}, Year: {year}")
        
        if not self.app_config.noise_dir:
            self.logger.warning("❌ Noise directory not configured")
            return {'type': 'none', 'txt_files': [], 'images': [], 'processed_data': None}
        
        # Performance optimization: Use cached base directory check
        base_noise_dir = Path(self.app_config.noise_dir)
        if not hasattr(self, '_base_dir_exists'):
            self._base_dir_exists = base_noise_dir.exists()
        
        if not self._base_dir_exists:
            return {'type': 'none', 'txt_files': [], 'images': [], 'processed_data': None}
        
        # Clean and format inputs efficiently
        year_clean = str(year).strip()
        test_num_padded = str(test_number).strip().zfill(4)
        
        # Resolve directory using shared cache (handles fallbacks internally)
        target_folder = self.directory_cache.resolve_test_folder(
            base_noise_dir,
            year_clean,
            test_num_padded,
        )

        if target_folder is None:
            return {'type': 'none', 'txt_files': [], 'images': [], 'processed_data': None}
        
        # Performance optimization: Single directory scan with efficient filtering
        try:
            # Use set comprehensions for O(1) lookups
            txt_extensions = {'.txt'}
            image_extensions = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.tif'}

            found_txt_files = []
            found_images = []

            # Single iteration through directory with efficient filtering
            for item in target_folder.iterdir():
                if not item.is_file():
                    continue

                item_suffix = item.suffix.lower()
                if item_suffix in txt_extensions:
                    found_txt_files.append(item)
                elif item_suffix in image_extensions:
                    found_images.append(item)
            
            # Process results based on priority (TXT files first)
            if found_txt_files:
                self.logger.debug(f"✅ Found {len(found_txt_files)} TXT files")
                # Performance: Only process TXT files if specifically needed
                processed_data = self._process_noise_txt_files_optimized(found_txt_files)
                return {
                    'type': 'txt_data',
                    'txt_files': found_txt_files,
                    'images': found_images,
                    'processed_data': processed_data
                }
            else:
                if not found_images:
                    # Populate from cache to avoid second scan if images were missed earlier
                    found_images = self.directory_cache.list_image_files(target_folder)

                if found_images:
                    self.logger.debug(f"✅ Found {len(found_images)} images")
                return {
                    'type': 'images',
                    'txt_files': [],
                    'images': found_images,
                    'processed_data': None
                }

            return {'type': 'none', 'txt_files': [], 'images': [], 'processed_data': None}
            
        except (OSError, PermissionError) as e:
            self.logger.error(f"❌ Error accessing folder {target_folder}: {e}")
            return {'type': 'error', 'txt_files': [], 'images': [], 'processed_data': None}

    def get_noise_images_simple(self, test_number: str, year: str) -> List[Path]:
        """
        Optimized noise image finder for better performance when processing multiple tests.
        
        PERFORMANCE OPTIMIZATIONS:
        - Cached directory existence checks
        - Single directory iteration with efficient filtering
        - Early termination for empty directories
        - Memory-efficient path operations
        
        Args:
            test_number: The noise test number (N.PROVA from registry)
            year: The year for the test
            
        Returns:
            List of image file paths found
        """
        if not self.app_config.noise_dir:
            return []
        
        # Performance: Cache base directory existence
        base_noise_dir = Path(self.app_config.noise_dir)
        if not hasattr(self, '_base_dir_exists'):
            self._base_dir_exists = base_noise_dir.exists()
        
        if not self._base_dir_exists:
            return []
        
        # Performance: Efficient input cleaning
        test_num_padded = str(test_number).strip().zfill(4)
        year_clean = str(int(float(str(year).strip()))) if year else ""
        
        if not year_clean:
            return []
        
        target_folder = self.directory_cache.resolve_test_folder(
            base_noise_dir,
            year_clean,
            test_num_padded,
        )

        if target_folder is None:
            return []

        images = self.directory_cache.list_image_files(target_folder)
        if images:
            self.logger.debug(f"✅ Found {len(images)} images in {target_folder.name}")
        return images

    def get_noise_test_info(self, sap_code: str, test_date_str: Optional[str] = None, 
                           test_number_for_lab_matching: Optional[str] = None) -> Optional[NoiseTestInfo]:
        """
        Get noise test information for a given SAP code.
        The date and test number come from the REGISTRO RUMORE, not from performance tests.
        
        Args:
            sap_code: The SAP code to search for
            test_date_str: Optional performance test date for filtering (UNUSED - dates come from registry)
            test_number_for_lab_matching: Optional test number for matching (UNUSED - test numbers come from registry)
            
        Returns:
            NoiseTestInfo object if found, None otherwise
        """
        if self.registro_df.empty:
            self.logger.warning("Registry DataFrame is empty - no noise data available")
            return None
        
        if not sap_code or not sap_code.strip():
            self.logger.warning("Empty SAP code provided")
            return None
        
        # Ensure required columns exist
        required_cols = [N_PROVA_STD, CODICE_SAP_STD]
        missing_cols = [col for col in required_cols if col not in self.registro_df.columns]
        if missing_cols:
            self.logger.error(f"Missing required columns: {missing_cols}")
            return None
        
        # Find candidates by SAP code
        sap_normalized = normalize_sap_code(sap_code)
        
        # DEBUG: Log SAP code matching
        self.logger.debug(f"Looking for SAP code: {sap_code} (normalized: {sap_normalized})")
        
        # Try exact match first
        candidates = self.registro_df[self.registro_df[CODICE_SAP_STD] == sap_normalized].copy()
        
        # If no exact match, try flexible matching
        if candidates.empty:
            self.logger.debug(f"No exact match for SAP {sap_normalized}, trying flexible matching...")
            
            # Try different variations of the SAP code
            sap_variations = [
                sap_code.strip(),  # Original
                sap_code.strip().upper(),  # Uppercase
                sap_code.strip().lower(),  # Lowercase
                sap_code.replace(" ", "").strip(),  # No spaces
                sap_code.replace("-", "").strip(),  # No dashes
            ]
            
            for sap_var in sap_variations:
                if sap_var != sap_normalized:  # Don't repeat the normalized version
                    candidates = self.registro_df[self.registro_df[CODICE_SAP_STD] == sap_var].copy()
                    if not candidates.empty:
                        self.logger.debug(f"Found match with SAP variation: {sap_var}")
                        break
        
        if candidates.empty:
            self.logger.info(f"No noise tests found for SAP code: {sap_code}")
            # DEBUG: Show available SAP codes for comparison
            available_saps = self.registro_df[CODICE_SAP_STD].dropna().unique()[:10]  # First 10
            self.logger.debug(f"Available SAP codes (first 10): {list(available_saps)}")
            return None
        
        self.logger.info(f"Found {len(candidates)} candidates for SAP {sap_code}")
        
        # For noise tests, we don't filter by performance test dates or numbers
        # We use the date and test number from the REGISTRO RUMORE itself
        
        if candidates.empty:
            self.logger.info(f"No noise tests found for SAP code: {sap_code}")
            return None
        
        # Select the best candidate (most recent, then highest N_PROVA)
        sort_columns = []
        sort_ascending = []
        
        if ANNO_DATETIME_STD in candidates.columns:
            sort_columns.append(ANNO_DATETIME_STD)
            sort_ascending.append(False)  # Most recent first
        
        sort_columns.append(N_PROVA_STD)
        sort_ascending.append(False)  # Highest N_PROVA first
        
        selected = candidates.sort_values(
            by=sort_columns,
            ascending=sort_ascending,
            na_position='last'
        ).iloc[0]
        
        self.logger.info(f"Selected noise test N.PROVA: {selected[N_PROVA_STD]}")
        
        # Extract information
        n_prova = str(selected[N_PROVA_STD])
        
        # Extract year
        year = ""
        if ANNO_YYYY_STD in selected and pd.notna(selected[ANNO_YYYY_STD]):
            try:
                year = str(int(selected[ANNO_YYYY_STD]))
            except (ValueError, TypeError):
                pass
        
        # Extract test lab
        test_lab = ""
        if TEST_LAB_STD in selected and pd.notna(selected[TEST_LAB_STD]):
            test_lab = str(selected[TEST_LAB_STD])
          # Find associated images using simplified approach
        images = self.get_noise_images_simple(n_prova, year)
        
        return NoiseTestInfo(
            nprova=n_prova,
            test_lab=test_lab,
            year=year,
            image_paths=images
        )

    def get_all_noise_tests_with_images(self, sap_code: str, test_date_str: Optional[str] = None) -> List[NoiseTestInfo]:
        """
        Get ALL noise test records with their images for a given SAP code.
        
        Args:
            sap_code: The SAP code to search for
            test_date_str: Optional performance test date for filtering
            
        Returns:
            List of NoiseTestInfo objects for all matching tests
        """
        if self.registro_df.empty:
            self.logger.warning("Registry DataFrame is empty - no noise data available")
            return []
        
        if not sap_code or not sap_code.strip():
            self.logger.warning("Empty SAP code provided")
            return []
        
        # Ensure required columns exist
        required_cols = [N_PROVA_STD, CODICE_SAP_STD]
        missing_cols = [col for col in required_cols if col not in self.registro_df.columns]
        if missing_cols:
            self.logger.error(f"Missing required columns: {missing_cols}")
            return []
        
        # Find all candidates by SAP code
        sap_normalized = normalize_sap_code(sap_code)
        candidates = self.registro_df[self.registro_df[CODICE_SAP_STD] == sap_normalized].copy()
        
        if candidates.empty:
            self.logger.info(f"No noise tests found for SAP code: {sap_code}")
            return []
        
        self.logger.info(f"Found {len(candidates)} noise tests for SAP {sap_code}")
        
        # Filter by test date if provided
        if test_date_str and ANNO_DATETIME_STD in candidates.columns:
            try:
                # Try parsing the date
                test_date = pd.to_datetime(test_date_str, errors='coerce', dayfirst=True)
                if pd.isna(test_date):
                    test_date = pd.to_datetime(test_date_str, errors='coerce')
                
                if pd.notna(test_date):
                    self.logger.info(f"Filtering by test date: {test_date.date()}")
                    # Keep noise tests that are on or before the performance test date
                    candidates = candidates[candidates[ANNO_DATETIME_STD] <= test_date]
                    self.logger.info(f"{len(candidates)} candidates remaining after date filter")
                
            except Exception as e:
                self.logger.warning(f"Could not parse test date '{test_date_str}': {e}")
        
        if candidates.empty:
            self.logger.info(f"No candidates remaining after filters for SAP {sap_code}")
            return []
        
        # Sort by date (most recent first), then by N_PROVA (highest first)
        sort_columns = []
        sort_ascending = []
        
        if ANNO_DATETIME_STD in candidates.columns:
            sort_columns.append(ANNO_DATETIME_STD)
            sort_ascending.append(False)  # Most recent first
        
        sort_columns.append(N_PROVA_STD)
        sort_ascending.append(False)  # Highest N_PROVA first
        
        sorted_candidates = candidates.sort_values(
            by=sort_columns,
            ascending=sort_ascending,
            na_position='last'
        )
        
        # Process each test record to create NoiseTestInfo objects
        noise_test_infos = []
        
        for _, row in sorted_candidates.iterrows():
            # Extract information for this test
            n_prova = str(row[N_PROVA_STD])
            
            # Extract year
            year = ""
            if ANNO_YYYY_STD in row and pd.notna(row[ANNO_YYYY_STD]):
                try:
                    year = str(int(row[ANNO_YYYY_STD]))
                except (ValueError, TypeError):
                    pass
            
            # Extract test lab
            test_lab = ""
            
            # Find associated images for this specific test using simplified approach
            images = self.get_noise_images_simple(n_prova, year)
            
            # Create NoiseTestInfo object
            noise_info = NoiseTestInfo(
                nprova=n_prova,
                test_lab=test_lab,
                year=year,
                image_paths=images
            )
            
            noise_test_infos.append(noise_info)
            
            self.logger.info(f"Processed N.PROVA {n_prova}: {len(images)} images found")
        
        total_images = sum(len(info.image_paths) for info in noise_test_infos)
        self.logger.info(f"Total: {len(noise_test_infos)} tests with {total_images} total images for SAP {sap_code}")
        
        return noise_test_infos

    def get_all_noise_tests_for_sap(self, sap_code: str) -> List[Dict[str, Any]]:
        """Get all noise test records for a SAP code."""
        if self.registro_df.empty or not sap_code:
            return []
        
        sap_normalized = normalize_sap_code(sap_code)
        matches = self.registro_df[self.registro_df[CODICE_SAP_STD] == sap_normalized]
        
        if matches.empty:
            return []
        
        # Convert to list of dictionaries with string keys
        result = []
        for _, row in matches.iterrows():
            record = {}
            for col, val in row.items():
                record[str(col)] = val
            result.append(record)
        
        return result

    def get_registro_df_for_display(self) -> pd.DataFrame:
        """Get a copy of the processed registry DataFrame."""
        return self.registro_df.copy()

    def get_noise_images_from_gui(self, test_number: str, year: str, sap_code: Optional[str] = None) -> NoiseTestInfo:
        """
        Simplified method for GUI - just takes test number and year from the noise registry
        and finds all images in the corresponding folder.
        
        Args:
            test_number: The test number from the noise registry (N.PROVA)
            year: The year for the test
            sap_code: Optional SAP code for additional context
            
        Returns:
            NoiseTestInfo object with found images
        """
        self.logger.info(f"🎯 GUI Noise Request: Test #{test_number}, Year: {year}, SAP: {sap_code}")
        
        # Get images using the simplified approach
        images = self.get_noise_images_simple(test_number, year)
        
        # Create and return NoiseTestInfo
        return NoiseTestInfo(
            nprova=test_number,
            year=year,
            test_lab="",  # Not needed for simple approach
            image_paths=images
        )

    def _process_noise_txt_files(self, txt_files: List[Path]) -> dict:
        """
        Process noise measurement TXT files and extract acoustic data.
        
        Args:
            txt_files: List of TXT file paths to process
            
        Returns:
            Dict containing processed noise measurement data
        """
        processed_data = {
            'summary': {},
            'measurements': [],
            'frequency_data': {},
            'statistics': {},
            'charts_data': {},
            'tables_data': {}
        }
        
        try:
            all_measurements = []
            
            for txt_file in txt_files:
                self.logger.info(f"📖 Processing noise file: {txt_file.name}")
                
                try:
                    # Read and parse the TXT file
                    file_data = self._parse_noise_txt_file(txt_file)
                    if file_data:
                        all_measurements.append(file_data)
                        
                except Exception as e:
                    self.logger.error(f"❌ Error processing {txt_file.name}: {e}")
                    continue
            
            if not all_measurements:
                self.logger.warning("⚠️ No valid measurements found in TXT files")
                return processed_data
            
            # Process and analyze all measurements
            processed_data = self._analyze_noise_measurements(all_measurements)
            
            self.logger.info(f"✅ Successfully processed {len(all_measurements)} noise measurement files")
            return processed_data
            
        except Exception as e:
            self.logger.error(f"❌ Error in _process_noise_txt_files: {e}")
            return processed_data

    def _parse_noise_txt_file(self, txt_file: Path) -> dict:
        """
        Parse a single noise measurement TXT file.
        
        Args:
            txt_file: Path to the TXT file
            
        Returns:
            Dict containing parsed measurement data
        """
        try:
            with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            
            # Initialize data structure
            measurement_data = {
                'filename': txt_file.name,
                'test_info': {},
                'sound_pressure_levels': {},
                'sound_power_levels': {},
                'frequency_data': {},
                'microphone_data': {},
                'raw_content': content
            }
            
            lines = content.split('\n')
            current_section = None
            
            for line_num, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue
                
                # Extract test information
                if 'Data:' in line or 'Date:' in line:
                    measurement_data['test_info']['date'] = line.split(':', 1)[-1].strip()
                elif 'Ora:' in line or 'Time:' in line:
                    measurement_data['test_info']['time'] = line.split(':', 1)[-1].strip()
                elif 'File:' in line:
                    measurement_data['test_info']['source_file'] = line.split(':', 1)[-1].strip()
                
                # Extract sound levels
                elif 'dB splA' in line and 'Lps' in line:
                    # Sound pressure level
                    try:
                        value = float(line.split()[0])
                        measurement_data['sound_pressure_levels']['Lps_dB_splA'] = value
                    except (ValueError, IndexError):
                        pass
                elif 'dB wA' in line and 'Lws' in line:
                    # Sound power level
                    try:
                        value = float(line.split()[0])
                        measurement_data['sound_power_levels']['Lws_dB_wA'] = value
                    except (ValueError, IndexError):
                        pass
                
                # Extract frequency data
                elif 'Hz' in line and any(char.isdigit() for char in line):
                    # This looks like frequency data
                    try:
                        parts = line.split()
                        if len(parts) >= 2:
                            freq_str = parts[0].replace(',', '.')
                            if 'Hz' in freq_str:
                                freq_str = freq_str.replace('Hz', '').strip()
                            frequency = float(freq_str)
                            
                            # Extract measurement values
                            values = []
                            for part in parts[1:]:
                                try:
                                    val = float(part.replace(',', '.'))
                                    values.append(val)
                                except ValueError:
                                    continue
                            
                            if values:
                                measurement_data['frequency_data'][frequency] = values
                    except (ValueError, IndexError):
                        pass
                
                # Extract microphone-specific data
                elif 'MIC' in line.upper() and any(char.isdigit() for char in line):
                    # Microphone data line
                    try:
                        if ':' in line:
                            mic_info, values_str = line.split(':', 1)
                            mic_name = mic_info.strip()
                            
                            # Parse values
                            values = []
                            for val_str in values_str.split():
                                try:
                                    val = float(val_str.replace(',', '.'))
                                    values.append(val)
                                except ValueError:
                                    continue
                            
                            if values:
                                measurement_data['microphone_data'][mic_name] = values
                    except (ValueError, IndexError):
                        pass
            
            # Calculate basic statistics
            measurement_data['statistics'] = self._calculate_basic_statistics(measurement_data)
            
            return measurement_data
            
        except Exception as e:
            self.logger.error(f"❌ Error parsing TXT file {txt_file}: {e}")
            return {}

    def _analyze_noise_measurements(self, measurements: List[dict]) -> dict:
        """
        Analyze processed noise measurements and prepare data for charts/tables.
        
        Args:
            measurements: List of parsed measurement data
            
        Returns:
            Dict containing analysis results and chart/table data
        """
        try:
            analysis_result = {
                'summary': {
                    'total_measurements': len(measurements),
                    'date_range': {},
                    'avg_sound_pressure': 0,
                    'avg_sound_power': 0,
                    'frequency_range': {}
                },
                'measurements': measurements,
                'frequency_data': {},
                'statistics': {},
                'charts_data': {
                    'frequency_response': {},
                    'sound_levels_comparison': {},
                    'microphone_comparison': {}
                },
                'tables_data': {
                    'summary_table': [],
                    'frequency_table': [],
                    'microphone_table': []
                }
            }
            
            # Aggregate sound pressure levels
            pressure_levels = []
            power_levels = []
            all_frequencies = set()
            
            for measurement in measurements:
                # Collect sound levels
                if 'Lps_dB_splA' in measurement.get('sound_pressure_levels', {}):
                    pressure_levels.append(measurement['sound_pressure_levels']['Lps_dB_splA'])
                if 'Lws_dB_wA' in measurement.get('sound_power_levels', {}):
                    power_levels.append(measurement['sound_power_levels']['Lws_dB_wA'])
                
                # Collect frequencies
                all_frequencies.update(measurement.get('frequency_data', {}).keys())
            
            # Calculate averages
            if pressure_levels:
                analysis_result['summary']['avg_sound_pressure'] = sum(pressure_levels) / len(pressure_levels)
            if power_levels:
                analysis_result['summary']['avg_sound_power'] = sum(power_levels) / len(power_levels)
            
            # Frequency analysis
            if all_frequencies:
                sorted_frequencies = sorted(all_frequencies)
                analysis_result['summary']['frequency_range'] = {
                    'min': min(sorted_frequencies),
                    'max': max(sorted_frequencies),
                    'count': len(sorted_frequencies)
                }
                
                # Prepare frequency response chart data
                freq_chart_data = {}
                for freq in sorted_frequencies:
                    freq_values = []
                    for measurement in measurements:
                        if freq in measurement.get('frequency_data', {}):
                            freq_values.extend(measurement['frequency_data'][freq])
                    
                    if freq_values:
                        freq_chart_data[freq] = {
                            'average': sum(freq_values) / len(freq_values),
                            'min': min(freq_values),
                            'max': max(freq_values),
                            'values': freq_values
                        }
                
                analysis_result['charts_data']['frequency_response'] = freq_chart_data
            
            # Prepare summary table data
            summary_table = []
            for i, measurement in enumerate(measurements):
                row = {
                    'measurement_id': i + 1,
                    'filename': measurement.get('filename', f'Measurement {i+1}'),
                    'date': measurement.get('test_info', {}).get('date', 'N/A'),
                    'time': measurement.get('test_info', {}).get('time', 'N/A'),
                    'sound_pressure': measurement.get('sound_pressure_levels', {}).get('Lps_dB_splA', 'N/A'),
                    'sound_power': measurement.get('sound_power_levels', {}).get('Lws_dB_wA', 'N/A'),
                    'frequency_points': len(measurement.get('frequency_data', {}))
                }
                summary_table.append(row)
            
            analysis_result['tables_data']['summary_table'] = summary_table
            
            # Prepare frequency table
            if all_frequencies:
                frequency_table = []
                for freq in sorted(all_frequencies):
                    row = {'frequency_hz': freq}
                    for i, measurement in enumerate(measurements):
                        if freq in measurement.get('frequency_data', {}):
                            values = measurement['frequency_data'][freq]
                            row[f'measurement_{i+1}'] = values[0] if values else 'N/A'
                        else:
                            row[f'measurement_{i+1}'] = 'N/A'
                    frequency_table.append(row)
                
                analysis_result['tables_data']['frequency_table'] = frequency_table
            
            # Calculate overall statistics
            analysis_result['statistics'] = {
                'pressure_levels': {
                    'mean': sum(pressure_levels) / len(pressure_levels) if pressure_levels else 0,
                    'min': min(pressure_levels) if pressure_levels else 0,
                    'max': max(pressure_levels) if pressure_levels else 0,
                    'count': len(pressure_levels)
                },
                'power_levels': {
                    'mean': sum(power_levels) / len(power_levels) if power_levels else 0,
                    'min': min(power_levels) if power_levels else 0,
                    'max': max(power_levels) if power_levels else 0,
                    'count': len(power_levels)
                },
                'frequencies': {
                    'total_unique': len(all_frequencies),
                    'range': f"{min(all_frequencies):.1f} - {max(all_frequencies):.1f} Hz" if all_frequencies else "N/A"
                }
            }
            
            self.logger.info(f"📊 Analysis complete: {len(measurements)} measurements, {len(all_frequencies)} frequency points")
            return analysis_result
            
        except Exception as e:
            self.logger.error(f"❌ Error in _analyze_noise_measurements: {e}")
            return {
                'summary': {}, 'measurements': measurements, 'frequency_data': {},
                'statistics': {}, 'charts_data': {}, 'tables_data': {}
            }

    def _calculate_basic_statistics(self, measurement_data: dict) -> dict:
        """Calculate basic statistics for a single measurement."""
        try:
            stats = {}
            
            # Frequency data statistics
            freq_data = measurement_data.get('frequency_data', {})
            if freq_data:
                all_values = []
                for values in freq_data.values():
                    all_values.extend(values)
                
                if all_values:
                    stats['frequency_stats'] = {
                        'mean': sum(all_values) / len(all_values),
                        'min': min(all_values),
                        'max': max(all_values),
                        'count': len(all_values)
                    }
            
            # Microphone data statistics
            mic_data = measurement_data.get('microphone_data', {})
            if mic_data:
                mic_stats = {}
                for mic_name, values in mic_data.items():
                    if values:
                        mic_stats[mic_name] = {
                            'mean': sum(values) / len(values),
                            'min': min(values),
                            'max': max(values)
                        }
                stats['microphone_stats'] = mic_stats
            
            return stats
            
        except Exception as e:
            self.logger.error(f"❌ Error calculating statistics: {e}")
            return {}

    def get_noise_test_info_comprehensive(self, sap_code: str) -> List[NoiseTestInfo]:
        """
        Get comprehensive noise test information for a SAP code - OPTIMIZED VERSION.
        
        PERFORMANCE OPTIMIZATIONS:
        - Uses existing registry data efficiently
        - Single pass through candidates
        - Optimized file operations
        """
        try:
            if self.registro_df.empty:
                self.logger.info(f"No registry data available for SAP {sap_code}")
                return []
            
            # Use existing method to get noise test info
            noise_info = self.get_noise_test_info(sap_code)
            if noise_info:
                return [noise_info]
            else:
                return []
            
        except Exception as e:
            self.logger.error(f"❌ Error getting comprehensive noise info for {sap_code}: {e}")
            return []
    
    def _process_noise_txt_files_optimized(self, txt_files: List[Path]) -> dict:
        """
        Optimized version of TXT file processing for better performance.
        
        PERFORMANCE OPTIMIZATIONS:
        - Lazy loading: Only process files when needed
        - Memory efficient: Process files one at a time
        - Faster parsing: Skip unnecessary processing for summary data
        - Early termination: Stop if basic data found
        
        Args:
            txt_files: List of TXT file paths to process
            
        Returns:
            Dict containing optimized noise measurement data
        """
        if not txt_files:
            return {'summary': {}, 'measurements': [], 'processed': False}
        
        try:
            # Performance: Process only first few files for summary if many files exist
            max_files_to_process = min(len(txt_files), 10)  # Limit processing for performance
            
            processed_measurements = []
            total_pressure_levels = []
            total_power_levels = []
            
            for i, txt_file in enumerate(txt_files[:max_files_to_process]):
                try:
                    # Fast parse: Extract only essential data
                    measurement_data = self._parse_noise_txt_file_fast(txt_file)
                    if measurement_data:
                        processed_measurements.append(measurement_data)
                        
                        # Collect key metrics for summary
                        if 'sound_pressure_levels' in measurement_data:
                            pressure_val = measurement_data['sound_pressure_levels'].get('Lps_dB_splA')
                            if pressure_val is not None:
                                total_pressure_levels.append(pressure_val)
                        
                        if 'sound_power_levels' in measurement_data:
                            power_val = measurement_data['sound_power_levels'].get('Lws_dB_wA')
                            if power_val is not None:
                                total_power_levels.append(power_val)
                        
                except Exception as e:
                    self.logger.debug(f"⚠️ Skipping file {txt_file.name}: {e}")
                    continue
            
            # Create optimized summary
            summary = {
                'total_files': len(txt_files),
                'processed_files': len(processed_measurements),
                'avg_sound_pressure': sum(total_pressure_levels) / len(total_pressure_levels) if total_pressure_levels else 0,
                'avg_sound_power': sum(total_power_levels) / len(total_power_levels) if total_power_levels else 0,
                'optimization': 'fast_processing_enabled'
            }
            
            self.logger.debug(f"✅ Fast processed {len(processed_measurements)}/{len(txt_files)} TXT files")
            
            return {
                'summary': summary,
                'measurements': processed_measurements,
                'processed': True
            }
            
        except Exception as e:
            self.logger.error(f"❌ Error in optimized TXT processing: {e}")
            return {'summary': {}, 'measurements': [], 'processed': False}

    def _parse_noise_txt_file_fast(self, txt_file: Path) -> dict:
        """
        Fast parser for TXT files - extracts only essential data for performance.
        
        Args:
            txt_file: Path to the TXT file
            
        Returns:
            Dict containing essential measurement data
        """
        try:
            # Performance: Read file with optimized encoding and error handling
            with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
                # Performance: Read only first portion of file for essential data
                content = f.read(2048)  # Read first 2KB only for basic info
            
            # Initialize essential data structure
            measurement_data = {
                'filename': txt_file.name,
                'sound_pressure_levels': {},
                'sound_power_levels': {},
                'test_info': {}
            }
            
            lines = content.split('\n')
            
            # Performance: Process only essential lines
            for line in lines[:50]:  # Process only first 50 lines
                line = line.strip()
                if not line:
                    continue
                
                # Extract essential sound levels only
                if 'dB splA' in line and 'Lps' in line:
                    try:
                        value = float(line.split()[0])
                        measurement_data['sound_pressure_levels']['Lps_dB_splA'] = value
                    except (ValueError, IndexError):
                        pass
                elif 'dB wA' in line and 'Lws' in line:
                    try:
                        value = float(line.split()[0])
                        measurement_data['sound_power_levels']['Lws_dB_wA'] = value
                    except (ValueError, IndexError):
                        pass
                elif 'Data:' in line or 'Date:' in line:
                    measurement_data['test_info']['date'] = line.split(':', 1)[-1].strip()
                
                # Early termination if we have essential data
                if (measurement_data['sound_pressure_levels'] and 
                    measurement_data['sound_power_levels']):
                    break
            
            return measurement_data
            
        except Exception as e:
            self.logger.debug(f"❌ Fast parse error for {txt_file}: {e}")
            return {}
