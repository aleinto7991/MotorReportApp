import logging
import re
from pathlib import Path
from typing import Optional
import pandas as pd

from .models import InfData

class InfParser:
    """Parses .inf files."""
    def __init__(self):
        self.logger = logging.getLogger(__class__.__name__)

    def parse(self, inf_path: Path) -> InfData:
        info = InfData()
        if not inf_path.exists():
            self.logger.error(f"INF file not found: {inf_path}")
            return info # Return empty InfData

        content_lines = None
        for encoding in ['latin1', 'utf-8', 'cp1252']:
            try:
                with inf_path.open('r', encoding=encoding) as f:
                    content_lines = f.readlines()
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                self.logger.error(f"Error opening INF file {inf_path} with encoding {encoding}: {e}")
        
        if content_lines is None:
            self.logger.error(f"Could not decode INF file: {inf_path}")
            return info

        try:
            in_info_aggiuntive = False
            comment_lines = []
            for line in content_lines:
                line = line.strip()
                if line.startswith('[Info_Aggiuntive]'):
                    in_info_aggiuntive = True
                    continue
                elif in_info_aggiuntive and line.startswith('['):
                    in_info_aggiuntive = False
                
                if '=' not in line: continue
                parts = line.split('=', 1)
                if len(parts) < 2: continue
                key, value = parts[0].strip(), parts[1].strip().strip('"')
                if not value: continue

                if key in ('Tipo motore', 'TipoMotore', 'Motor Type'): info.motor_type = value
                elif key in ('Data', 'Date'): info.date = value
                elif key in ('Tensione', 'Voltage'):
                    match = re.search(r'(\d+(?:\.\d+)?)', value.replace(',', '.'))
                    info.voltage = match.group(1) if match else value
                elif key in ('Frequenza', 'Frequency'):
                    info.hz = value.replace('Hz', '').replace('hz', '').strip()
                elif in_info_aggiuntive and key in ('Note', 'Notes', 'Comment'):
                    comment_lines.append(value.replace('\\0A', '\n').replace('\0A', '\n').replace('\\n', '\n'))
            
            info.comment = '\n'.join(comment_lines)
            if not info.motor_type: self.logger.warning(f"No motor type found in {inf_path}")
        except Exception as e:
            self.logger.error(f"Error parsing INF file content {inf_path}: {e}")
        return info

class CsvParser:
    """Parses .csv files and performs initial processing."""
    def __init__(self):
        self.logger = logging.getLogger(__class__.__name__)

    def parse(self, csv_path: Path) -> Optional[pd.DataFrame]:
        if not csv_path.exists():
            self.logger.error(f"CSV file not found: {csv_path}")
            return None
        try:
            df = pd.read_csv(csv_path, sep=';', encoding='latin1', decimal=',') # Added decimal=',',
            # Standardize column names to simplify processing later
            df.columns = [str(col).strip() for col in df.columns]

            # Convert all potential numeric columns, handling commas as decimal separators
            for column in df.columns:
                if df[column].dtype == object:
                    try:
                        # Attempt to convert to numeric, replacing comma with dot for decimals
                        df[column] = pd.to_numeric(df[column].astype(str).str.replace(',', '.'), errors='raise')
                    except (AttributeError, ValueError, TypeError):
                        # If conversion fails, keep original, log warning if it looks like it should be numeric
                        if any(char.isdigit() for char in str(df[column].iloc[0] if not df[column].empty else "")):
                            self.logger.debug(f"Column '{column}' in {csv_path.name} could not be fully converted to numeric and was left as object.")
                        pass # Keep as object if not convertible
            
            self._perform_unit_conversions(df)
            # self._add_calculated_columns(df) # This can be called after conversions if needed
            return df
        except FileNotFoundError:
            self.logger.error(f"CSV file not found during parse: {csv_path}")
            return None
        except pd.errors.EmptyDataError:
            self.logger.error(f"CSV file is empty: {csv_path}")
            return None
        except Exception as e:
            self.logger.error(f"Error reading or processing CSV file {csv_path}: {e}", exc_info=True)
            return None

    def _perform_unit_conversions(self, df: pd.DataFrame) -> None:
        """Convert units for vacuum (to kPa) and air flow (to m³/h)."""
        self.logger.debug(f"Starting unit conversions. Columns: {df.columns.tolist()}")
        # Vacuum conversion: mmH2O to kPa (1 mmH2O = 0.00980665 kPa)
        vacuum_col_mmh2o = None
        if 'Vacuum Corrected mmH2O' in df.columns:
            vacuum_col_mmh2o = 'Vacuum Corrected mmH2O'
        elif 'Vacuum mmH2O' in df.columns: # Fallback if corrected is not present
            vacuum_col_mmh2o = 'Vacuum mmH2O'
        elif 'Pressione Finale vuoto (mmH2O)' in df.columns: # Italian name
             vacuum_col_mmh2o = 'Pressione Finale vuoto (mmH2O)'

        if vacuum_col_mmh2o and pd.api.types.is_numeric_dtype(df[vacuum_col_mmh2o]):
            df['Vacuum Corrected (kPa)'] = df[vacuum_col_mmh2o] * 0.00980665
            self.logger.info(f"Converted '{vacuum_col_mmh2o}' to 'Vacuum Corrected (kPa)'.")
            # Optionally drop the original mmH2O column if it's not 'Vacuum Corrected (kPa)' already
            # if vacuum_col_mmh2o != 'Vacuum Corrected (kPa)' and vacuum_col_mmh2o in df.columns:
            #     df.drop(columns=[vacuum_col_mmh2o], inplace=True)
        elif vacuum_col_mmh2o: # Column exists but is not numeric
            self.logger.warning(f"Vacuum column '{vacuum_col_mmh2o}' found but is not numeric. Cannot convert to kPa.")
        else:
            self.logger.warning("No suitable vacuum column (mmH2O) found for kPa conversion.")

        # Air Flow conversion: l/sec to m³/h (1 l/sec = 3.6 m³/h)
        airflow_col_lsec = None
        if 'Air Flow l/sec.' in df.columns:
            airflow_col_lsec = 'Air Flow l/sec.'
        elif 'Portata Finale Aria (l/s)' in df.columns: # Italian name variant
            airflow_col_lsec = 'Portata Finale Aria (l/s)'
        
        if airflow_col_lsec and pd.api.types.is_numeric_dtype(df[airflow_col_lsec]):
            df['Air Flow (m³/h)'] = df[airflow_col_lsec] * 3.6
            self.logger.info(f"Converted '{airflow_col_lsec}' to 'Air Flow (m³/h)'.")
            # Optionally drop the original l/sec column
            # if airflow_col_lsec != 'Air Flow (m³/h)' and airflow_col_lsec in df.columns:
            #    df.drop(columns=[airflow_col_lsec], inplace=True)
        elif airflow_col_lsec: # Column exists but is not numeric
            self.logger.warning(f"Air flow column '{airflow_col_lsec}' found but is not numeric. Cannot convert to m³/h.")
        else:
            self.logger.warning("No suitable air flow column (l/sec) found for m³/h conversion.")

    # _add_calculated_columns can be kept or modified if other calculations are needed
    # def _add_calculated_columns(self, df: pd.DataFrame) -> None:
    #     pass
