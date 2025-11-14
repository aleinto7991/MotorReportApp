from __future__ import annotations

import pandas as pd
import pytest

from src.config.measurement_units import apply_unit_preferences
from src.config.app_config import AppConfig


def _assert_almost(value: float, expected: float, rel: float = 1e-9) -> None:
    assert value == pytest.approx(expected, rel=rel)


def test_apply_unit_preferences_converts_and_renames():
    df = pd.DataFrame(
        {
            "Vacuum Corrected (kPa)": [1.0],
            "Air Flow (m³/h)": [360.0],
            "Power Corrected Watts in": [1000.0],
            "Speed RPM": [1800.0],
        }
    )

    config = AppConfig(
        pressure_unit="mmH2O",
        flow_unit="l/s",
        power_unit="kW",
        speed_unit="rps",
    )

    converted, metadata = apply_unit_preferences(df, config)

    expected_columns = {
        "Vacuum Corrected (mmH₂O)",
        "Air Flow (l/s)",
        "Power Corrected Input (kW)",
        "Speed (rps)",
    }

    assert expected_columns.issubset(set(converted.columns))

    _assert_almost(converted["Vacuum Corrected (mmH₂O)"].iloc[0], 101.971621)
    _assert_almost(converted["Air Flow (l/s)"].iloc[0], 100.0)
    _assert_almost(converted["Power Corrected Input (kW)"].iloc[0], 1.0)
    _assert_almost(converted["Speed (rps)"].iloc[0], 30.0)

    assert metadata["pressure"]["unit"] == "mmH2O"
    assert metadata["flow"]["unit"] == "l/s"
    assert metadata["power"]["unit"] == "kW"
    assert metadata["speed"]["unit"] == "rps"


def test_apply_unit_preferences_drops_unselected_columns():
    df = pd.DataFrame(
        {
            "Vacuum Corrected mmH2O": [150.0],
            "Air Flow l/sec": [120.0],
            "Power Watts in": [800.0],
            "Speed (r/min)": [1200.0],
        }
    )

    config = AppConfig(
        pressure_unit="kPa",
        flow_unit="m³/h",
        power_unit="W",
        speed_unit="rpm",
    )

    converted, _ = apply_unit_preferences(df, config)

    # Legacy columns should be removed in favor of the selected-unit columns
    assert "Vacuum Corrected mmH2O" not in converted.columns
    assert "Air Flow l/sec" not in converted.columns
    assert "Power Watts in" not in converted.columns
    assert "Speed (r/min)" not in converted.columns

    assert "Vacuum Corrected (kPa)" in converted.columns
    assert "Air Flow (m³/h)" in converted.columns
    assert "Power Corrected Input (W)" in converted.columns
    assert "Speed (rpm)" in converted.columns  # Re-created with canonical naming
