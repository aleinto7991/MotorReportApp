"""Utilities for applying measurement unit preferences across the report."""
from __future__ import annotations

from typing import Dict, Iterable, Tuple

import pandas as pd

# Use try/except for flexible import paths
try:
    from .app_config import AppConfig
except ImportError:
    from .app_config import AppConfig

# Base column names generated during CSV parsing
PRESSURE_BASE_COLUMN = "Vacuum Corrected (kPa)"
FLOW_BASE_COLUMN = "Air Flow (m³/h)"
POWER_BASE_COLUMN = "Power Corrected Watts in"
SPEED_BASE_COLUMN = "Speed RPM"

# Candidate source columns that may appear before normalization
PRESSURE_SOURCE_CANDIDATES = [
    "Vacuum Corrected mmH2O",
    "Vacuum mmH2O",
    "Vacuum (mmH2O)",
    "Vacuum Corrected (mmH2O)",
    "Pressione Finale vuoto (mmH2O)",
]

FLOW_SOURCE_CANDIDATES = [
    "Air Flow l/sec.",
    "Air Flow l/sec",
    "Air Flow (l/sec)",
    "Air Flow (l/s)",
    "Portata Finale Aria (l/s)",
    "Airflow (l/sec)",
    "Airflow l/sec",
]

POWER_SOURCE_CANDIDATES = [
    "Power Watts in",
    "Power Input (W)",
    "Power (W)",
]

SPEED_SOURCE_CANDIDATES = [
    "Speed (rpm)",
    "Speed (RPM)",
    "Speed (r/min)",
]

PRESSURE_UNIT_SETTINGS = {
    "kPa": {
        "label": "Vacuum Corrected (kPa)",
        "axis_label": "Vacuum Corrected (kPa)",
        "chart_name": "Vacuum (kPa)",
        "factor": 1.0,
    },
    "mmH2O": {
        "label": "Vacuum Corrected (mmH₂O)",
        "axis_label": "Vacuum Corrected (mmH₂O)",
        "chart_name": "Vacuum (mmH₂O)",
        "factor": 101.971621,
    },
    "psi": {
        "label": "Vacuum Corrected (psi)",
        "axis_label": "Vacuum Corrected (psi)",
        "chart_name": "Vacuum (psi)",
        "factor": 0.145037738,
    },
}

FLOW_UNIT_SETTINGS = {
    "m³/h": {
        "label": "Air Flow (m³/h)",
        "axis_label": "Air Flow (m³/h)",
        "chart_name": "Air Flow (m³/h)",
        "factor": 1.0,
    },
    "l/s": {
        "label": "Air Flow (l/s)",
        "axis_label": "Air Flow (l/s)",
        "chart_name": "Air Flow (l/s)",
        "factor": 1 / 3.6,
    },
    "CFM": {
        "label": "Air Flow (CFM)",
        "axis_label": "Air Flow (CFM)",
        "chart_name": "Air Flow (CFM)",
        "factor": 1 / 1.69901082,
    },
}

POWER_UNIT_SETTINGS = {
    "W": {
        "label": "Power Corrected Input (W)",
        "axis_label": "Power Corrected Input (W)",
        "chart_name": "Power Input (W)",
        "factor": 1.0,
    },
    "kW": {
        "label": "Power Corrected Input (kW)",
        "axis_label": "Power Corrected Input (kW)",
        "chart_name": "Power Input (kW)",
        "factor": 0.001,
    },
    "HP": {
        "label": "Power Corrected Input (HP)",
        "axis_label": "Power Corrected Input (HP)",
        "chart_name": "Power Input (HP)",
        "factor": 1 / 745.699872,
    },
}

SPEED_UNIT_SETTINGS = {
    "rpm": {
        "label": "Speed (rpm)",
        "axis_label": "Speed (rpm)",
        "chart_name": "Speed (rpm)",
        "factor": 1.0,
    },
    "rps": {
        "label": "Speed (rps)",
        "axis_label": "Speed (rps)",
        "chart_name": "Speed (rps)",
        "factor": 1 / 60.0,
    },
}

UNIT_CONFIG = {
    "pressure": {
        "base": PRESSURE_BASE_COLUMN,
        "candidates": PRESSURE_SOURCE_CANDIDATES,
        "settings": PRESSURE_UNIT_SETTINGS,
    },
    "flow": {
        "base": FLOW_BASE_COLUMN,
        "candidates": FLOW_SOURCE_CANDIDATES,
        "settings": FLOW_UNIT_SETTINGS,
    },
    "power": {
        "base": POWER_BASE_COLUMN,
        "candidates": POWER_SOURCE_CANDIDATES,
        "settings": POWER_UNIT_SETTINGS,
    },
    "speed": {
        "base": SPEED_BASE_COLUMN,
        "candidates": SPEED_SOURCE_CANDIDATES,
        "settings": SPEED_UNIT_SETTINGS,
    },
}


def apply_unit_preferences(df: pd.DataFrame, config: AppConfig) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    """Return a copy of ``df`` with measurement columns converted to the units selected in ``config``.

    The resulting metadata dictionary contains column names and labels for each measurement category
    so that downstream builders can update headers, chart titles, and axis labels consistently.
    """

    converted_df = df.copy()
    metadata: Dict[str, Dict[str, str]] = {}

    _ensure_base_columns(converted_df)

    # Apply measurement-specific conversions
    metadata["pressure"] = _convert_measurement(
        converted_df,
        measurement="pressure",
        selected_unit=config.pressure_unit,
    )
    metadata["flow"] = _convert_measurement(
        converted_df,
        measurement="flow",
        selected_unit=config.flow_unit,
    )
    metadata["power"] = _convert_measurement(
        converted_df,
        measurement="power",
        selected_unit=config.power_unit,
    )
    metadata["speed"] = _convert_measurement(
        converted_df,
        measurement="speed",
        selected_unit=config.speed_unit,
    )

    _drop_unused_unit_columns(converted_df, metadata)

    return converted_df, metadata


def _ensure_base_columns(df: pd.DataFrame) -> None:
    """Make sure the canonical base measurement columns exist using best-effort conversion."""
    _ensure_pressure_base(df)
    _ensure_flow_base(df)
    _ensure_power_base(df)
    _ensure_speed_base(df)


def _ensure_pressure_base(df: pd.DataFrame) -> None:
    if PRESSURE_BASE_COLUMN in df.columns:
        df[PRESSURE_BASE_COLUMN] = pd.to_numeric(df[PRESSURE_BASE_COLUMN], errors="coerce")
        return

    source = _find_first_numeric(df, PRESSURE_SOURCE_CANDIDATES)
    if source:
        df[PRESSURE_BASE_COLUMN] = pd.to_numeric(df[source], errors="coerce") * 0.00980665
    else:
        df[PRESSURE_BASE_COLUMN] = pd.NA


def _ensure_flow_base(df: pd.DataFrame) -> None:
    if FLOW_BASE_COLUMN in df.columns:
        df[FLOW_BASE_COLUMN] = pd.to_numeric(df[FLOW_BASE_COLUMN], errors="coerce")
        return

    source = _find_first_numeric(df, FLOW_SOURCE_CANDIDATES)
    if source:
        df[FLOW_BASE_COLUMN] = pd.to_numeric(df[source], errors="coerce") * 3.6
    else:
        df[FLOW_BASE_COLUMN] = pd.NA


def _ensure_power_base(df: pd.DataFrame) -> None:
    if POWER_BASE_COLUMN in df.columns:
        df[POWER_BASE_COLUMN] = pd.to_numeric(df[POWER_BASE_COLUMN], errors="coerce")
        return

    source = _find_first_numeric(df, POWER_SOURCE_CANDIDATES)
    if source:
        df[POWER_BASE_COLUMN] = pd.to_numeric(df[source], errors="coerce")
    else:
        df[POWER_BASE_COLUMN] = pd.NA


def _ensure_speed_base(df: pd.DataFrame) -> None:
    if SPEED_BASE_COLUMN in df.columns:
        df[SPEED_BASE_COLUMN] = pd.to_numeric(df[SPEED_BASE_COLUMN], errors="coerce")
        return

    source = _find_first_numeric(df, SPEED_SOURCE_CANDIDATES)
    if source:
        df[SPEED_BASE_COLUMN] = pd.to_numeric(df[source], errors="coerce")
    else:
        df[SPEED_BASE_COLUMN] = pd.NA


def _convert_measurement(df: pd.DataFrame, measurement: str, selected_unit: str) -> Dict[str, str]:
    config = UNIT_CONFIG[measurement]
    base_column = config["base"]
    settings = config["settings"]

    unit_info = settings.get(selected_unit) or settings[next(iter(settings))]

    target_column = unit_info["label"]
    axis_label = unit_info["axis_label"]
    chart_name = unit_info["chart_name"]
    factor = unit_info["factor"]

    if base_column in df.columns:
        numeric = pd.to_numeric(df[base_column], errors="coerce")
        
        if target_column != base_column:
            # Get the position of the base column to maintain column order
            base_col_idx = df.columns.get_loc(base_column)
            # Drop the base column
            df.drop(columns=[base_column], inplace=True)
            # Insert the converted column at the same position
            df.insert(base_col_idx, target_column, numeric * factor)
        else:
            # Same column name, just update values in place
            df[base_column] = numeric * factor
    else:
        df[target_column] = pd.NA

    return {
        "column": target_column,
        "axis_label": axis_label,
        "chart_name": chart_name,
        "unit": selected_unit,
    }


def _drop_unused_unit_columns(df: pd.DataFrame, metadata: Dict[str, Dict[str, str]]) -> None:
    keep_columns = {info["column"] for info in metadata.values()}

    candidates: Iterable[str] = set().union(
        PRESSURE_SOURCE_CANDIDATES,
        FLOW_SOURCE_CANDIDATES,
        POWER_SOURCE_CANDIDATES,
        SPEED_SOURCE_CANDIDATES,
        [PRESSURE_BASE_COLUMN, FLOW_BASE_COLUMN, POWER_BASE_COLUMN, SPEED_BASE_COLUMN],
    )

    for column in candidates:
        if column in df.columns and column not in keep_columns:
            df.drop(columns=[column], inplace=True)


def _find_first_numeric(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    for column in candidates:
        if column in df.columns:
            series = pd.to_numeric(df[column], errors="coerce")
            if series.notna().any():
                df[column] = series
                return column
    return None
