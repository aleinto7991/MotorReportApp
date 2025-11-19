from __future__ import annotations

from pathlib import Path

import pytest

from src.config.app_config import AppConfig
from src.data.models import (
    InfData,
    MotorTestData,
    TestLabSummary,
    SchedaSummary,
    CollaudoSummary,
)
from src.reports.excel_report import ExcelReport


def test_excel_report_resolves_logo_from_directory_locator(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """If the config omits a logo path, ExcelReport should reuse the DirectoryLocator fallback."""

    fallback_logo = tmp_path / "assets" / "logo.png"
    fallback_logo.parent.mkdir(parents=True, exist_ok=True)
    fallback_logo.write_bytes(b"fake")

    class _DirectoryLocatorStub:
        def __init__(self, logger=None):  # pragma: no cover - trivial plumbing
            self.logo_path = fallback_logo

    monkeypatch.setattr("src.reports.excel_report.DirectoryLocator", _DirectoryLocatorStub)

    config = AppConfig(
        tests_folder=str(tmp_path / "tests"),
        registry_path=str(tmp_path / "registry.xlsx"),
        output_path=str(tmp_path / "report.xlsx"),
    )

    report = ExcelReport(config)

    assert report.logo_path == str(fallback_logo)
    assert config.logo_path == str(fallback_logo)


def test_excel_report_generate_keeps_output_path(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Report generation should keep the configured output path when no post-processing runs."""

    base_output = tmp_path / "report.xlsx"
    config = AppConfig(
        tests_folder=str(tmp_path / "tests"),
        registry_path=str(tmp_path / "registry.xlsx"),
        output_path=str(base_output),
    )

    class _ExcelWriterStub:
        def __init__(self, path: str, engine: str | None = None) -> None:  # pragma: no cover - stub wiring
            self.path = Path(path)
            self.book = object()

        def __enter__(self) -> "_ExcelWriterStub":
            return self

        def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - simple context stub
            return None

    monkeypatch.setattr("src.reports.excel_report.pd.ExcelWriter", _ExcelWriterStub)
    monkeypatch.setattr("src.reports.excel_report.extract_dominant_colors", lambda _p: ["#111111"])

    class _FormatterStub:
        def __init__(self, workbook, colors):  # pragma: no cover - simple data holder
            self.workbook = workbook
            self.colors = colors

    monkeypatch.setattr("src.reports.excel_report.ExcelFormatter", _FormatterStub)

    def _noop(*_args, **_kwargs) -> None:  # pragma: no cover - helper for builder stubs
        return None

    for method_name in [
        "_create_summary_sheet",
        "_create_sap_sheets",
        "_create_comparison_sheet",
        "_create_multiple_comparison_sheets",
    ]:
        monkeypatch.setattr(ExcelReport, method_name, _noop)

    motor_test = MotorTestData(
        test_number="T1",
        inf_data=InfData(motor_type="SAP1", date="2024-01-01"),
        csv_path=tmp_path / "T1.csv",
        csv_data=None,
    )

    report = ExcelReport(config)
    success = report.generate(
        grouped_data={"SAP1": [motor_test]},
        all_tests_summary=[motor_test],
        all_noise_tests_by_sap={"SAP1": []},
        comparison_data={},
        multiple_comparisons=None,
    )

    assert success is True
    assert report.output_path == base_output
    assert config.output_path == str(base_output)


def test_excel_report_builds_carichi_sheets(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """CarichiSheetBuilder should be invoked for SAP groups that carry test-lab summaries."""

    base_output = tmp_path / "report.xlsx"
    config = AppConfig(
        tests_folder=str(tmp_path / "tests"),
        registry_path=str(tmp_path / "registry.xlsx"),
        output_path=str(base_output),
    )

    class _ExcelWriterStub:
        def __init__(self, path: str, engine: str | None = None) -> None:  # pragma: no cover - stub wiring
            self.path = Path(path)
            self.book = object()

        def __enter__(self) -> "_ExcelWriterStub":
            return self

        def __exit__(self, exc_type, exc, tb) -> None:  # pragma: no cover - simple context stub
            return None

    monkeypatch.setattr("src.reports.excel_report.pd.ExcelWriter", _ExcelWriterStub)
    monkeypatch.setattr("src.reports.excel_report.extract_dominant_colors", lambda _p: ["#222222"])

    class _FormatterStub:
        def __init__(self, workbook, colors):  # pragma: no cover - simple data holder
            self.workbook = workbook
            self.colors = colors

        def get(self, _key):
            return None

    monkeypatch.setattr("src.reports.excel_report.ExcelFormatter", _FormatterStub)

    def _noop(*_args, **_kwargs) -> None:  # pragma: no cover - helper for builder stubs
        return None

    for method_name in [
        "_create_summary_sheet",
        "_create_comparison_sheet",
        "_create_multiple_comparison_sheets",
    ]:
        monkeypatch.setattr(ExcelReport, method_name, _noop)

    sap_calls: list[str] = []

    class _SapBuilderStub:
        def __init__(
            self,
            workbook,
            sap_code,
            motor_tests,
            formatter,
            config,
            logo_tab_colors,
            all_noise_tests,
            noise_handler,
        ) -> None:  # pragma: no cover - simple stub
            sap_calls.append(sap_code)

        def build(self) -> None:
            return None

    monkeypatch.setattr("src.reports.excel_report.SapSheetBuilder", _SapBuilderStub)

    carichi_calls: list[tuple[str, list[MotorTestData]]] = []

    class _CarichiBuilderStub:
        def __init__(self, workbook, sap_code, motor_tests, formatter, logo_tab_colors, logo_path=None) -> None:  # pragma: no cover
            carichi_calls.append((sap_code, list(motor_tests)))

        def build(self) -> bool:
            return True

    monkeypatch.setattr("src.reports.excel_report.CarichiSheetBuilder", _CarichiBuilderStub)

    scheda = SchedaSummary(headers=["Or1"], rows={"Media": {"Or1": 1.0}})
    collaudo = CollaudoSummary(headers=["P1"], values={"P1": 0.9})
    tls = TestLabSummary(source_path=str(tmp_path / "test_lab.xlsx"), scheda=scheda, collaudo_media=collaudo)

    motor_test = MotorTestData(
        test_number="T1",
        inf_data=InfData(motor_type="SAP1", date="2024-01-01"),
        csv_path=tmp_path / "T1.csv",
        csv_data=None,
        test_lab_summary=tls,
    )

    report = ExcelReport(config)
    success = report.generate(
        grouped_data={"SAP1": [motor_test]},
        all_tests_summary=[motor_test],
        all_noise_tests_by_sap={"SAP1": []},
        comparison_data={},
        multiple_comparisons=None,
    )

    assert success is True
    assert sap_calls == ["SAP1"], "SAP builder should still run"
    assert len(carichi_calls) == 1, "Carichi builder should be invoked once"
    assert carichi_calls[0][0] == "SAP1"
    assert carichi_calls[0][1][0].test_lab_summary is tls
