from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd
import pytest


@pytest.fixture
def directory_config_stub(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Dict[str, Path]:
    """Provide a sandboxed set of directories/files and patch `directory_config` to use them."""
    from src import directory_config as config

    base = tmp_path / "sandbox"
    performance_dir = base / "performance"
    noise_dir = base / "noise"
    output_dir = base / "output"
    logs_dir = base / "logs"
    project_root = base / "project"
    onedrive_root = base / "onedrive"
    logo_path = project_root / "assets" / "logo.png"
    lab_registry = onedrive_root / "registries" / "lab.xlsx"
    noise_registry = onedrive_root / "registries" / "noise.xlsx"

    for path in [
        performance_dir,
        noise_dir,
        output_dir,
        logs_dir,
        project_root / "assets",
        project_root / "logs",
        lab_registry.parent,
        noise_registry.parent,
    ]:
        path.mkdir(parents=True, exist_ok=True)

    logo_path.write_bytes(b"logo")
    lab_registry.write_bytes(b"lab")
    noise_registry.write_bytes(b"noise")

    monkeypatch.setattr(config, "PROJECT_ROOT", project_root)
    monkeypatch.setattr(config, "ONEDRIVE_ROOT", onedrive_root)
    monkeypatch.setattr(config, "PERFORMANCE_TEST_DIR", performance_dir)
    monkeypatch.setattr(config, "NOISE_TEST_DIR", noise_dir)
    monkeypatch.setattr(config, "LAB_REGISTRY_FILE", lab_registry)
    monkeypatch.setattr(config, "NOISE_REGISTRY_FILE", noise_registry)
    monkeypatch.setattr(config, "OUTPUT_DIR", output_dir)
    monkeypatch.setattr(config, "LOGS_DIR", logs_dir)
    monkeypatch.setattr(config, "LOGO_PATH", logo_path)

    return {
        "performance": performance_dir,
        "noise": noise_dir,
        "output": output_dir,
        "logs": logs_dir,
        "project": project_root,
        "onedrive": onedrive_root,
        "lab_registry": lab_registry,
        "noise_registry": noise_registry,
        "logo": logo_path,
    }


@pytest.fixture
def sample_registry_excel(tmp_path: Path) -> Path:
    """Create a minimal noise registry workbook using common column variations."""
    data = {
        "N. PROVA": [1, 2, 3],
        "COD.MOTORE": ["SAP001", "sap002", " Sap003 "],
        "DATA": ["2024-01-02", "2024-03-05", "2024-06-07"],
        "TEST LAB": ["L1", "L2", "L3"],
        "NOTE": ["", "Check", None],
        "TENSIONE": [400, 380, 400],
    }
    df = pd.DataFrame(data)
    path = tmp_path / "noise_registry.xlsx"
    df.to_excel(path, index=False)
    return path
