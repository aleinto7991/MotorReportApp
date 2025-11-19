"""Data models for Motor Report Application."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict
import pandas as pd


@dataclass
class Test:
    """Represents a single test entry from the registry."""
    __test__ = False  # Prevent pytest from mistaking this dataclass for a test suite
    test_lab_number: str
    sap_code: str
    voltage: str
    notes: str
    date: Optional[str] = None  # Date information from INF file or registry


@dataclass
class InfData:
    """Data parsed from an .inf file."""
    motor_type: str = ''
    date: str = ''
    voltage: str = ''
    hz: str = ''
    comment: str = ''


@dataclass
class NoiseTestInfo:
    """Data class for noise test information."""
    test_number: Optional[str] = None
    date: Optional[str] = None
    mic_position: Optional[str] = None
    background_noise: Optional[float] = None
    motor_noise: Optional[float] = None
    result: Optional[str] = None
    image_path: Optional[Path] = None
    # SAP code for this noise test
    sap_code: Optional[str] = None
    # Legacy fields, can be removed if not used elsewhere
    nprova: Optional[str] = None
    test_lab: Optional[str] = None
    year: Optional[str] = None
    image_paths: List[Path] = field(default_factory=list)
    # New comprehensive data fields for TXT file analysis
    data_type: str = "images"  # "txt_data", "images", or "none"
    txt_files: List[Path] = field(default_factory=list)
    processed_data: Optional[dict] = None  # Contains parsed measurements, charts, tables data


@dataclass
class SchedaSummary:
    """Summary block collected from the Scheda SR sheet."""
    headers: List[str]
    rows: Dict[str, Dict[str, Optional[float]]]
    notes: List[str] = field(default_factory=list)


@dataclass
class CollaudoSummary:
    """Media row collected from the Collaudo SR sheet."""
    headers: List[str]
    values: Dict[str, Optional[float]]


@dataclass
class TestLabSummary:
    """Bundle of extracted data from the external test-lab workbook."""
    __test__ = False  # Prevent pytest from trying to collect this dataclass
    source_path: Optional[str] = None
    scheda: Optional[SchedaSummary] = None
    collaudo_media: Optional[CollaudoSummary] = None
    matched_test_number: Optional[str] = None  # Candidate stem that matched during lookup
    match_strategy: str = "exact"  # How the workbook was resolved (exact, prefix, fallback_*)
    raw_sheets: List[Dict] = field(default_factory=list)  # List of raw sheet data (name, values, merges, col_widths)


@dataclass
class LifeTestInfo:
    """Data class for Life Test (LF) information."""
    test_number: str  # e.g., "LF053/07"
    sap_code: Optional[str] = None
    notes: Optional[str] = None
    year: Optional[str] = None  # Extracted from test number (e.g., "2007" from "LF053/07")
    test_id: Optional[str] = None  # Extracted from test number (e.g., "053" from "LF053/07")
    file_path: Optional[Path] = None  # Path to the LF Excel file
    file_exists: bool = False  # Whether the file was found
    responsible: Optional[str] = None  # Responsabile column from registry


@dataclass
class MotorTestData:
    """Comprehensive data for a single motor test."""
    test_number: str
    inf_data: InfData
    csv_path: Path
    csv_data: Optional[pd.DataFrame] = None
    noise_info: Optional[NoiseTestInfo] = None
    noise_status: str = 'NOT FOUND' # For summary
    status_message: str = 'OK' # For summary
    test_lab_summary: Optional[TestLabSummary] = None

    @property
    def sap_code(self) -> str:
        return self.inf_data.motor_type if self.inf_data else 'UNKNOWN_SAP'
