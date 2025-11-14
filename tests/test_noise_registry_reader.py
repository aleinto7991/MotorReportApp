from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from src.services import noise_registry_reader as reader


def test_load_registry_dataframe_normalises_columns(sample_registry_excel: Path):
    df, mapping = reader.load_registry_dataframe(sample_registry_excel, sheet_name=0)

    assert reader.N_PROVA_STD in df.columns
    assert reader.CODICE_SAP_STD in df.columns
    assert reader.ANNO_YYYY_STD in df.columns
    assert reader.ANNO_DATETIME_STD in df.columns

    assert mapping[reader.N_PROVA_STD] == "N. PROVA"
    assert mapping[reader.CODICE_SAP_STD] == "COD.MOTORE"

    assert df[reader.N_PROVA_STD].tolist() == ["0001", "0002", "0003"]
    assert df[reader.CODICE_SAP_STD].tolist() == ["SAP001", "SAP002", "SAP003"]
    assert all(year == 2024 for year in df[reader.ANNO_YYYY_STD].tolist())
    assert all(ts.year == 2024 for ts in df[reader.ANNO_DATETIME_STD].tolist())


def test_build_column_mapping_handles_spacing_and_punctuation():
    columns = ["N PROVA", "Codice Sap", "Other"]
    mapping = reader.build_column_mapping(columns)

    assert mapping[reader.N_PROVA_STD] == "N PROVA"
    assert mapping[reader.CODICE_SAP_STD] == "Codice Sap"


def test_clean_registry_dataframe_filters_invalid_rows():
    df = pd.DataFrame({
        reader.N_PROVA_STD: [None, " 123 ", "AB12"],
        reader.CODICE_SAP_STD: ["", "sap100", None],
        reader.ANNO_ORIGINAL_STD: ["2020", "2021", ""],
    })

    cleaned = reader.clean_registry_dataframe(df)

    assert cleaned[reader.N_PROVA_STD].tolist() == ["0123"]
    assert cleaned[reader.CODICE_SAP_STD].tolist() == ["SAP100"]
    assert cleaned[reader.ANNO_YYYY_STD].tolist() == [2021]


def test_load_registry_dataframe_missing_file_raises(tmp_path: Path):
    missing = tmp_path / "missing.xlsx"
    with pytest.raises(FileNotFoundError):
        reader.load_registry_dataframe(missing)
