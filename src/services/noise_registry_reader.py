"""Shared utilities for reading and normalising the noise registry Excel files."""
from __future__ import annotations

import logging
import re
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple, Any

import pandas as pd

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Standard column keys used throughout the application
# ---------------------------------------------------------------------------
N_PROVA_STD = "N_PROVA"
CODICE_SAP_STD = "CODICE_SAP"
ANNO_ORIGINAL_STD = "ANNO_ORIGINAL"
TEST_LAB_STD = "TEST_LAB"
NOTE_STD = "NOTE"
VOLTAGE_STD = "TENSIONE"
CLIENT_STD = "CLIENTE"
APPLICATION_STD = "APPARECCHIO"
RESPONSIBLE_STD = "RESPONSABILE"
ANNO_YYYY_STD = "ANNO_YYYY"
ANNO_DATETIME_STD = "ANNO_DATETIME"

COL_NAME_VARIATIONS: Dict[str, Iterable[str]] = {
    N_PROVA_STD: [
        "N. PROVA",
        "N° PROVA",
        "N.PROVA",
        "N PROVA",
    "PROVA",
        "NUMERO PROVA",
        "NUM PROVA",
        "TEST_NO",
        "N_PROVA",
        "NUM. PROVA",
        "TEST NO.",
        "TEST NO",
        "TEST NUMBER",
    ],
    CODICE_SAP_STD: [
        "CODICE SAP",
        "COD.MOTORE",
        "CODICE MOTOR",
        "COD. SAP",
        "SAP",
        "MOTORE",
        "CODICE_SAP",
        "COD. MOTORE",
        "CODIGO MOTORE",
        "SAP CODE",
    ],
    ANNO_ORIGINAL_STD: [
        "ANNO",
        "DATA",
        "DATE",
        "YEAR",
        "DATETIME",
        "DATA PROVA",
        "DATA TEST",
    ],
    TEST_LAB_STD: [
        "TEST LAB",
        "RIF. TEST LAB",
        "LAB TEST REF",
        "RIFERIMENTO TEST DI LABORATORIO",
        "REF LAB",
        "TESTLAB",
        "LAB",
        "LABORATORIO",
        "TEST LAB.",
    ],
    NOTE_STD: [
        "NOTE",
        "OSSERVAZIONI",
        "ANNOTAZIONI",
        "VARIE",
        "COMMENTS",
        "COMMENTI",
    ],
    VOLTAGE_STD: [
        "TENSIONE",
        "VOLTAGE",
        "VOLT",
        "V",
        "TENS.",
        "TENS",
        "VN",
        "VOLTAGE (V)",
        "TENSIONE (V)",
    ],
    CLIENT_STD: [
        "CLIENTE",
        "CLIENT",
        "CUSTOMER",
        "COMM.",
        "COMMITTENTE",
    ],
    APPLICATION_STD: [
        "APPARECCHIO",
        "APPLICATION",
        "DEVICE",
        "EQUIPMENT",
        "APPAR.",
        "APPARATO",
    ],
    RESPONSIBLE_STD: [
        "RESP.",
        "RESPONSIBLE",
        "RESPONSABILE",
        "RESP",
        "RESPONSAB.",
    ],
}

HEADER_KEYWORDS = [variation.upper() for variations in COL_NAME_VARIATIONS.values() for variation in variations]


def _normalise_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    return re.sub(r"[^A-Z0-9]", "", text)

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def find_header_row(df_sample: pd.DataFrame, max_rows: int = 15, *, log: Optional[logging.Logger] = None) -> Optional[int]:
    """Attempt to locate the registry header row in a raw dataframe sample."""
    logger_to_use = log or logger

    for idx in range(min(max_rows, len(df_sample))):
        row_values = [str(val).strip().upper() for val in df_sample.iloc[idx].values if pd.notna(val)]
        matches = sum(1 for value in row_values if any(keyword in value for keyword in HEADER_KEYWORDS))
        if matches >= 2:
            logger_to_use.info("Detected registry header row at index %s", idx)
            return idx

    logger_to_use.warning("Could not detect registry header row within first %s rows", max_rows)
    return None


def build_column_mapping(columns: Iterable[str]) -> Dict[str, str]:
    """Return a mapping of canonical column names to the actual column titles."""
    mapping: Dict[str, str] = {}
    columns_list = list(columns)

    normalised_lookup: Dict[str, str] = {}
    for original in columns_list:
        key = _normalise_header(original)
        if key and key not in normalised_lookup:
            normalised_lookup[key] = original

    for canonical, variations in COL_NAME_VARIATIONS.items():
        for variation in variations:
            normalised_variation = _normalise_header(variation)
            if normalised_variation in normalised_lookup:
                mapping[canonical] = normalised_lookup[normalised_variation]
                break

    return mapping


def normalise_dataframe_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Rename dataframe columns to their canonical counterparts when possible."""
    mapping = build_column_mapping(df.columns)
    if not mapping:
        return df.copy(), {}

    rename_map = {original: canonical for canonical, original in mapping.items()}
    return df.rename(columns=rename_map), mapping


def _extract_year_from_value(value: Any) -> Optional[int]:
    if pd.isna(value) or value == "":
        return None

    try:
        if isinstance(value, datetime):
            return value.year

        value_str = str(value).strip()
        if value_str.isdigit() and len(value_str) == 4:
            year = int(value_str)
            if 1900 <= year <= 2100:
                return year

        try:
            dt = pd.to_datetime(value_str, errors="raise", dayfirst=True)
            return dt.year
        except Exception:
            pass

        match = re.search(r"\b(19|20)\d{2}\b", value_str)
        if match:
            return int(match.group())

        if value_str.replace(".", "").isdigit():
            try:
                excel_date = pd.to_datetime(float(value_str), origin="1899-12-30", unit="D")
                return excel_date.year
            except Exception:
                return None
    except Exception:
        return None

    return None


def _convert_to_datetime(value: Any) -> Optional[pd.Timestamp]:
    if pd.isna(value) or value == "":
        return None

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    try:
        return pd.to_datetime(value, errors="raise", dayfirst=True)
    except Exception:
        try:
            value_str = str(value).strip()
            if value_str.replace(".", "").isdigit():
                return pd.to_datetime(float(value_str), origin="1899-12-30", unit="D")
        except Exception:
            return None
    return None


def clean_registry_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Clean and standardise the noise registry dataframe."""
    df_clean = df.copy()

    if N_PROVA_STD in df_clean.columns:
        def pad_n_prova(val: Any) -> Optional[str]:
            if pd.isna(val) or val in {"nan", "NaN", "None", ""}:
                return None
            val_clean = str(val).strip().split(".")[0]
            if val_clean.isdigit():
                return val_clean.zfill(4)
            return val_clean

        df_clean[N_PROVA_STD] = df_clean[N_PROVA_STD].astype(str).str.strip().apply(pad_n_prova)

    if CODICE_SAP_STD in df_clean.columns:
        df_clean[CODICE_SAP_STD] = (
            df_clean[CODICE_SAP_STD]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace(r"\\+$", "", regex=True)
        )
        df_clean[CODICE_SAP_STD] = df_clean[CODICE_SAP_STD].replace(["NAN", "NONE", ""], None)

    if ANNO_ORIGINAL_STD in df_clean.columns:
        df_clean[ANNO_YYYY_STD] = df_clean[ANNO_ORIGINAL_STD].apply(_extract_year_from_value)
        df_clean[ANNO_DATETIME_STD] = df_clean[ANNO_ORIGINAL_STD].apply(_convert_to_datetime)

    if TEST_LAB_STD in df_clean.columns:
        df_clean[TEST_LAB_STD] = (
            df_clean[TEST_LAB_STD]
            .astype(str)
            .str.strip()
            .str.upper()
            .replace(["NAN", "NONE", ""], None)
        )

    if NOTE_STD not in df_clean.columns:
        df_clean[NOTE_STD] = None

    initial_count = len(df_clean)
    subset_columns = [col for col in (N_PROVA_STD, CODICE_SAP_STD) if col in df_clean.columns]
    if subset_columns:
        df_clean = df_clean.dropna(subset=subset_columns, how="all")

    if N_PROVA_STD in df_clean.columns:
        df_clean = df_clean[df_clean[N_PROVA_STD].notna()]
    else:
        logger.warning("Noise registry missing '%s' column after normalisation", N_PROVA_STD)

    if CODICE_SAP_STD in df_clean.columns:
        df_clean = df_clean[df_clean[CODICE_SAP_STD].notna()]
    else:
        logger.warning("Noise registry missing '%s' column after normalisation", CODICE_SAP_STD)

    removed = initial_count - len(df_clean)
    logger.info("Registry cleaning removed %s rows (from %s)", removed, initial_count)

    return df_clean


def load_registry_dataframe(
    excel_path: Path | str,
    *,
    sheet_name: Optional[str | int] = None,
    header_search_rows: int = 15,
    log: Optional[logging.Logger] = None,
    nrows: Optional[int] = None,
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Load and prepare the noise registry Excel file with intelligent caching.
    
    This function uses file modification time for cache invalidation,
    ensuring cached data is always current.
    """
    logger_to_use = log or logger
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Noise registry file not found: {path}")

    # Get file modification time for cache key
    mtime = path.stat().st_mtime
    
    # Use cached version if available
    return _load_registry_dataframe_cached(
        str(path),
        mtime,
        sheet_name=sheet_name,
        header_search_rows=header_search_rows,
        nrows=nrows,
        logger_name=logger_to_use.name if logger_to_use else None
    )


@lru_cache(maxsize=8)  # Cache up to 8 different registry files
def _load_registry_dataframe_cached(
    excel_path_str: str,
    mtime: float,  # File modification time for cache invalidation
    *,
    sheet_name: Optional[str | int] = None,
    header_search_rows: int = 15,
    nrows: Optional[int] = None,
    logger_name: Optional[str] = None
) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Cached implementation of registry loading.
    
    Cache key includes file path and modification time, so cache
    is automatically invalidated when file changes.
    """
    logger_to_use = logging.getLogger(logger_name) if logger_name else logger
    path = Path(excel_path_str)
    
    logger_to_use.debug(f"Loading noise registry from {path} (mtime: {mtime})")

    sheet = sheet_name if sheet_name is not None else 0

    df_peek = pd.read_excel(path, header=None, engine="openpyxl", sheet_name=sheet, nrows=nrows)
    header_row = find_header_row(df_peek, max_rows=header_search_rows, log=logger_to_use)

    if header_row is None:
        logger_to_use.warning("Falling back to default header row 0 for noise registry")
        header_row = 0

    df = pd.read_excel(path, header=header_row, engine="openpyxl", sheet_name=sheet, nrows=nrows)
    df = df.dropna(how="all")

    normalised_df, mapping = normalise_dataframe_columns(df)
    cleaned_df = clean_registry_dataframe(normalised_df)

    logger_to_use.info(f"Loaded noise registry: {len(cleaned_df)} rows (cached)")
    return cleaned_df, mapping
