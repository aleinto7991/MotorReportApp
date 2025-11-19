"""Application configuration data class."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict, Set
try:
    from src._version import VERSION
except Exception:
    # Fallback if package import context differs
    from .._version import VERSION


@dataclass
class AppConfig:
    """Application configuration settings."""
    # Core paths, can be set by GUI or CLI
    tests_folder: Optional[str] = None
    registry_path: Optional[str] = None # This will be the main performance registry
    output_path: Optional[str] = None
    log_path: Optional[str] = None
    logo_path: Optional[str] = None

    # Auto-detected paths
    project_root: Optional[Path] = None
    onedrive_root: Optional[Path] = None
    performance_test_dir: Optional[Path] = None
    lab_registry_file: Optional[Path] = None

    # Optional paths, now populated from directory_config at runtime
    noise_registry_path: Optional[str] = None
    noise_dir: Optional[str] = None
    test_lab_root: Optional[str] = None  # Base directory for test lab "A" workbooks

    # Execution control
    no_gui: bool = False # To run in CLI mode
    direct_inputs: List[str] = field(default_factory=list) # For direct GUI input
    sap_codes: List[str] = field(default_factory=list) # SAPs to include in report
    compare_saps: List[str] = field(default_factory=list) # SAPs to compare
    include_noise: bool = True # Workflow toggle
    include_comparison: bool = True # Workflow toggle
    open_after_creation: bool = False
    verbose: bool = False
    # Keep a single source of truth for version
    script_version: str = VERSION
    registry_sheet_name: str = "REGISTRO" # Changed to match actual Excel sheet
    noise_registry_sheet_name: Optional[str] = None
    noise_saps: List[str] = field(default_factory=list)  # Track selected noise SAP codes
    
    # Fine-grained test lab selection for comparison (SAP code -> Set of test lab numbers)
    comparison_test_labs: Dict[str, List[str]] = field(default_factory=dict)
    
    # Pre-filtered noise tests from GUI (bypasses noise lookup if provided)
    selected_noise_tests: Optional[List] = None  # List of NoiseTestInfo objects
    
    # Life Test (LF) selections (SAP code -> Set of test numbers)
    selected_lf_test_numbers: Dict[str, Set[str]] = field(default_factory=dict)

    # Carichi/TestLab selections (Test Number -> File Path)
    selected_carichi_map: Dict[str, str] = field(default_factory=dict)
    
    # Multiple comparison groups for custom comparison sheets
    multiple_comparisons: List = field(default_factory=list)  # List of comparison group definitions

    # Measurement unit preferences
    pressure_unit: str = "kPa"
    flow_unit: str = "m³/h"
    speed_unit: str = "rpm"
    power_unit: str = "W"

    def __post_init__(self):
        # Resolve paths to be absolute
        if self.tests_folder:
            self.tests_folder = str(Path(self.tests_folder).resolve())
        if self.registry_path:
            self.registry_path = str(Path(self.registry_path).resolve())
        if self.output_path:
            self.output_path = str(Path(self.output_path).resolve())
        if self.log_path:
            self.log_path = str(Path(self.log_path).resolve())
        if self.logo_path:
            self.logo_path = str(Path(self.logo_path).resolve())
        if self.noise_registry_path:
            self.noise_registry_path = str(Path(self.noise_registry_path).resolve())
        if self.noise_dir:
            self.noise_dir = str(Path(self.noise_dir).resolve())
        if self.test_lab_root:
            self.test_lab_root = str(Path(self.test_lab_root).resolve())
