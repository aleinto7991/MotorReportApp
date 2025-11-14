from __future__ import annotations

from pathlib import Path
from typing import List, Tuple, Dict, Optional

import pytest

from src.ui.core.report_manager import ReportManager
from src.ui.core.state_manager import StateManager
from src.data.models import Test, MotorTestData, InfData
from src.config.app_config import AppConfig
from src.validators.noise_test_validator import NoiseTestValidationInfo
from src.core.motor_report_engine import MotorReportApp


class _DummyStatusManager:
    def __init__(self) -> None:
        self.messages: List[Tuple[str, str]] = []

    def update_status(self, message: str, color: str) -> None:
        self.messages.append((message, color))


class _DummyGUI:
    def __init__(self, state_manager: StateManager) -> None:
        self.state_manager = state_manager
        self.status_manager = _DummyStatusManager()


def _create_test(sap_code: str, test_no: str, date: str = "2024-01-01") -> Test:
    return Test(test_lab_number=test_no, sap_code=sap_code, voltage="230", notes="", date=date)


def test_validate_noise_tests_includes_gui_only_selections(monkeypatch: pytest.MonkeyPatch) -> None:
    """Ensure GUI-selected noise tests missing from the registry are preserved."""

    state_manager = StateManager()
    state = state_manager.state
    state.include_noise = True
    state.selected_tests_folder = "C:/tmp/tests"
    state.selected_registry_file = "C:/tmp/registry.xlsx"
    state.selected_noise_registry = "C:/tmp/noise_registry.xlsx"

    # GUI selections include an SAP that will be absent from the registry payload
    state.selected_noise_saps = {"612057"}
    state.selected_noise_test_labs = {
        "612057": {"28332", "28333"},
        "6120890824": {"55555"},
    }

    # Populate selected performance tests so the validator can enrich supplemental entries
    state.selected_tests = {
        "28332": _create_test("612057", "28332"),
        "28333": _create_test("612057", "28333"),
        "55555": _create_test("6120890824", "55555"),
    }

    gui = _DummyGUI(state_manager)
    report_manager = ReportManager(gui)  # type: ignore[arg-type]

    # Registry returns only a single matching entry; the others must be fabricated by the validator
    registry_entry = NoiseTestValidationInfo(
        sap_code="612057",
        test_no="28332",
        file_path="registry",
        exists=True,
        is_valid=True,
    )

    class _ValidatorStub:
        def __init__(self, noise_folder: str, sheet_name: str) -> None:  # pragma: no cover - simple wiring
            self.noise_folder = noise_folder
            self.sheet_name = sheet_name

        def validate_from_registry(self, registry_path: str) -> List[NoiseTestValidationInfo]:
            return [registry_entry]

    monkeypatch.setattr("src.validators.noise_test_validator.NoiseTestValidator", _ValidatorStub)
    monkeypatch.setattr("src.config.directory_config.NOISE_TEST_DIR", None, raising=False)
    monkeypatch.setattr("src.config.directory_config.NOISE_REGISTRY_FILE", None, raising=False)

    captured: dict[str, List[NoiseTestValidationInfo]] = {}

    def _capture(selected_noise_tests: List[NoiseTestValidationInfo]) -> None:
        captured["tests"] = selected_noise_tests

    report_manager.validate_noise_tests_if_needed(_capture)

    assert "tests" in captured, "callback should receive selected noise tests"

    selected_pairs = {
        (entry.sap_code, str(entry.test_no).strip())
        for entry in captured["tests"]
    }

    assert selected_pairs == {
        ("612057", "28332"),
        ("612057", "28333"),
        ("6120890824", "55555"),
    }

    # The second SAP should be promoted into the selected set even though the registry skipped it
    assert "6120890824" in state_manager.state.selected_noise_saps


class _DirectoryLocatorStub:
    def __init__(self) -> None:
        self.logger_calls: List[Tuple[str, Dict[str, object]]] = []

    def apply_defaults(self, config: AppConfig) -> None:  # pragma: no cover - simple stub
        # Keep paths untouched for tests while recording that we were invoked.
        self.logger_calls.append(("apply_defaults", {
            "tests_folder": config.tests_folder,
            "noise_dir": config.noise_dir,
        }))


class _NoiseHandlerStub:
    def __init__(self, image_root: Path) -> None:
        self.calls: List[Tuple[str, str]] = []
        self.image_root = image_root

    def get_noise_test_info_by_test_year(self, test_number: str, year: str, sap_code: Optional[str] = None):
        """Mock the new method that returns NoiseTestInfo with both images and TXT files."""
        from src.data.models import NoiseTestInfo
        from pathlib import Path as PathLib
        
        self.calls.append((test_number, year))
        image_path = self.image_root / f"noise_{test_number}_{year}.png"
        image_path.write_bytes(b"fake")
        
        return NoiseTestInfo(
            nprova=test_number,
            year=year,
            image_paths=[PathLib(image_path)],
            txt_files=[],  # No TXT files in this mock
            data_type="images",
            sap_code=sap_code
        )
    
    def get_noise_images_simple(self, test_number: str, year: str) -> List[str]:
        """Fallback method for backwards compatibility."""
        self.calls.append((test_number, year))
        image_path = self.image_root / f"noise_{test_number}_{year}.png"
        image_path.write_bytes(b"fake")
        return [str(image_path)]


def test_motor_report_app_propagates_prefiltered_noise(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """MotorReportApp should pass every pre-filtered noise test through to the Excel layer."""

    (tmp_path / "noise").mkdir()

    config = AppConfig(
        tests_folder=str(tmp_path),
        registry_path=str(tmp_path / "registry.xlsx"),
        output_path=str(tmp_path / "out.xlsx"),
        noise_registry_path=str(tmp_path / "noise_registry.xlsx"),
        noise_dir=str(tmp_path / "noise"),
        include_noise=True,
    )

    noise_handler = _NoiseHandlerStub(tmp_path)

    captured: Dict[str, object] = {}

    class _ExcelReportStub:
        def __init__(self, cfg: AppConfig, noise_handler=None) -> None:
            self.config = cfg
            self.noise_handler = noise_handler

        def generate(self, *, grouped_data, all_tests_summary, all_noise_tests_by_sap, comparison_data, multiple_comparisons, lf_tests_by_sap=None) -> bool:
            captured["noise_keys"] = {sap: [info.test_number for info in infos] for sap, infos in all_noise_tests_by_sap.items()}
            captured["grouped_keys"] = sorted(grouped_data.keys())
            captured["tests_summary"] = [data.test_number for data in all_tests_summary]
            return True

    # Patch ExcelReport where MotorReportApp references it (imported into core.motor_report_engine)
    monkeypatch.setattr("src.core.motor_report_engine.ExcelReport", _ExcelReportStub)

    directory_locator = _DirectoryLocatorStub()
    app = MotorReportApp(config, noise_handler=noise_handler, directory_locator=directory_locator)  # type: ignore[arg-type]

    motor_data_map = {
        "28332": MotorTestData(
            test_number="28332",
            inf_data=InfData(motor_type="612057", date="2024-01-15"),
            csv_path=tmp_path / "28332.csv",
            csv_data=None,
        ),
        "55555": MotorTestData(
            test_number="55555",
            inf_data=InfData(motor_type="6120890824", date="2024-02-20"),
            csv_path=tmp_path / "55555.csv",
            csv_data=None,
        ),
    }

    def _fake_process(self, test_number: str):  # pragma: no cover - used via monkeypatch
        return motor_data_map[test_number]

    monkeypatch.setattr(MotorReportApp, "_process_single_test", _fake_process, raising=False)

    selected_tests = [
        Test(test_lab_number="28332", sap_code="612057", voltage="230", notes="", date="2024-01-15"),
        Test(test_lab_number="55555", sap_code="6120890824", voltage="230", notes="", date="2024-02-20"),
    ]

    selected_noise_tests = [
        NoiseTestValidationInfo(
            sap_code="612057",
            test_no="28332",
            file_path="registry",
            exists=True,
            is_valid=True,
            date="2024",
        ),
        NoiseTestValidationInfo(
            sap_code="6120890824",
            test_no="55555",
            file_path="supplemental",
            exists=True,
            is_valid=True,
            date="2024",
        ),
    ]

    app.generate_report(
        selected_tests=selected_tests,
        selected_noise_tests=selected_noise_tests,
        include_noise_data=True,
    )

    assert captured["grouped_keys"] == ["612057", "6120890824"]
    assert captured["tests_summary"] == ["28332", "55555"]
    assert captured["noise_keys"] == {
        "612057": ["28332"],
        "6120890824": ["55555"],
    }
    assert noise_handler.calls == [("28332", "2024"), ("55555", "2024")]
