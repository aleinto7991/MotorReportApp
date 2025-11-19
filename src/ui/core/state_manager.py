"""
Application State Manager for the Motor Report GUI
Manages all application state including selected tests, configurations, and workflow state.
"""
import logging
import os
from typing import List, Dict, Set, Optional, Callable, Any
from dataclasses import dataclass, field
from ..utils.selection_cache import SelectionCache
from ...data.models import Test
from ...services.noise_registry_reader import load_registry_dataframe
from ...services.carichi_locator import CarichiLocator, CarichiTestInfo

logger = logging.getLogger(__name__)


@dataclass
class AppState:
    """Central application state"""
    # Test data
    found_tests: List[Test] = field(default_factory=list)
    selected_tests: Dict[str, Test] = field(default_factory=dict)
    found_sap_codes: List[str] = field(default_factory=list)
    
    # Workflow state
    workflow_step: int = 1
    search_selection_applied: bool = False
    config_selection_applied: bool = False
    
    # File paths
    selected_tests_folder: str = ""
    selected_registry_file: str = ""
    selected_noise_folder: str = ""
    selected_noise_registry: str = ""
    test_lab_directory: str = ""
    
    # SAP selections
    selected_performance_saps: Set[str] = field(default_factory=set)
    selected_noise_saps: Set[str] = field(default_factory=set)
    selected_comparison_saps: Set[str] = field(default_factory=set)
    
    # Fine-grained test lab selection for comparison (SAP -> Set of test lab numbers) - LEGACY
    selected_comparison_test_labs: Dict[str, Set[str]] = field(default_factory=dict)
    
    # NEW: Multiple comparison groups support
    # Each comparison group is a dict with: {"name": str, "test_labs": List[str], "description": str}
    multiple_comparisons: List[Dict[str, Any]] = field(default_factory=list)
    
    # NEW: Comparison groups from Config tab (group_id -> {sap_code -> Set[test_lab_numbers]})
    comparison_groups: Dict[str, Dict[str, Set[str]]] = field(default_factory=dict)
    
    # Fine-grained test lab selection for noise (SAP -> Set of test lab numbers)
    selected_noise_test_labs: Dict[str, Set[str]] = field(default_factory=dict)
    
    # Life Test (LF) selections (SAP -> Set of test numbers like "LF053/07")
    selected_lf_test_numbers: Dict[str, Set[str]] = field(default_factory=dict)
    selected_lf_saps: Set[str] = field(default_factory=set)
    
    # Configuration
    include_noise: bool = True
    include_comparison: bool = True
    registry_sheet_name: str = "REGISTRO"
    noise_registry_sheet_name: str = "Registro"

    # Measurement unit preferences
    pressure_unit: str = "kPa"
    flow_unit: str = "m³/h"
    speed_unit: str = "rpm"
    power_unit: str = "W"
    
    # Loaded registry data (to avoid redundant loading)
    noise_registry_dataframe: Optional[Any] = None  # pandas DataFrame

    # Carichi Nominali (Test Lab) tracking
    carichi_matches: Dict[str, Optional[CarichiTestInfo]] = field(default_factory=dict)
    carichi_last_checked: Optional[str] = None
    carichi_errors: List[str] = field(default_factory=list)
    
    # Internal flags
    shutting_down: bool = False
    operation_in_progress: bool = False
    current_operation: str = ""
    
    # Final validated data for report generation
    final_test_data: Dict = field(default_factory=dict)  # Complete structured data from config validation
    final_performance_data: Dict[str, Test] = field(default_factory=dict)
    final_comparison_data: Dict[str, List] = field(default_factory=dict)  # SAP -> List of tests
    final_noise_data: Dict[str, List] = field(default_factory=dict)  # SAP -> List of noise tests
    generation_summary: Dict = field(default_factory=dict)
    
    # Picker context for distinguishing different file picker uses
    picker_context: str = ""


class StateManager:
    """Manages application state and provides state change notifications"""
    
    def __init__(self):
        self.state = AppState()
        self._observers: List[Callable] = []
        self.selection_cache = SelectionCache()
        self._carichi_locator: Optional[CarichiLocator] = None
        self._carichi_locator_path: Optional[str] = None
        self._carichi_lookup_signature: Optional[tuple[str, tuple[str, ...]]] = None
    
    def add_observer(self, callback):
        """Add a callback to be notified when state changes"""
        self._observers.append(callback)
    
    def notify_observers(self, event_type: str, data: Optional[dict] = None):
        """Notify all observers of a state change"""
        for callback in self._observers:
            try:
                callback(event_type, data or {})
            except Exception as e:
                logger.warning(f"Observer callback failed: {e}")

    def start_operation(self, operation_name: str) -> bool:
        """
        Start an operation if no other operation is in progress.
        
        Args:
            operation_name: Name of the operation to start
            
        Returns:
            True if operation started successfully, False if another operation is in progress
        """
        if self.state.operation_in_progress:
            logger.warning(f"Cannot start '{operation_name}' - '{self.state.current_operation}' is already in progress")
            return False
            
        self.state.operation_in_progress = True
        self.state.current_operation = operation_name
        self.notify_observers("operation_started", {"operation": operation_name})
        logger.info(f"Started operation: {operation_name}")
        return True
        
    def end_operation(self):
        """End the current operation"""
        if self.state.operation_in_progress:
            operation_name = self.state.current_operation
            self.state.operation_in_progress = False
            self.state.current_operation = ""
            self.notify_observers("operation_ended", {"operation": operation_name})
            logger.info(f"Ended operation: {operation_name}")
        
    def is_operation_in_progress(self) -> bool:
        """Check if any operation is currently in progress"""
        return self.state.operation_in_progress
        
    def get_current_operation(self) -> str:
        """Get the name of the current operation"""
        return self.state.current_operation
    
    def reset_search(self):
        """Reset search-related state"""
        self.state.workflow_step = 1
        self.state.found_tests.clear()
        self.state.selected_tests.clear()
        self.state.found_sap_codes.clear()
        self.state.search_selection_applied = False
        self.selection_cache.selected_test_labs.clear()
        self.selection_cache.selected_sap_codes.clear()
        
        # Also clear SAP selections and fine-grained test lab selections
        self.state.selected_noise_saps.clear()
        self.state.selected_comparison_saps.clear()
        self.state.selected_comparison_test_labs.clear()
        self.state.config_selection_applied = False
        
        self.notify_observers("search_reset")
    
    def update_test_selection(self, test_id: str, test: Test, selected: bool):
        """Update test selection state with optimized notifications"""
        # Check if state actually changed to avoid unnecessary updates
        current_selected = test_id in self.state.selected_tests
        if current_selected == selected:
            return  # No change needed
        
        if selected:
            self.state.selected_tests[test_id] = test
        else:
            self.state.selected_tests.pop(test_id, None)
        
        self._invalidate_carichi_cache()

        # Only notify if there are observers and state actually changed
        if self._observers:
            self.notify_observers("test_selection_changed", {
                "test_id": test_id,
                "selected": selected,
                "total_selected": len(self.state.selected_tests)
            })

    def remove_selected_test(self, test_id: str) -> bool:
        """Remove a test from the selection and clean up dependent selections."""
        test = self.state.selected_tests.get(test_id)
        if not test:
            return False

        self.update_test_selection(test_id, test, False)
        self.apply_search_selection()

        sap_code = test.sap_code

        # Remove from noise selections
        self._remove_noise_test_reference(sap_code, test_id)

        # Remove from comparison selections
        self._remove_comparison_test_reference(sap_code, test_id)

        # If no tests remain for the SAP, clean up SAP-specific selections
        if sap_code and not any(t.sap_code == sap_code for t in self.state.selected_tests.values()):
            self.state.selected_noise_saps.discard(sap_code)
            self.state.selected_comparison_saps.discard(sap_code)
            self.state.selected_noise_test_labs.pop(sap_code, None)
            self.state.selected_comparison_test_labs.pop(sap_code, None)

            # Remove SAP from comparison groups (new structure)
            for group_id in list(self.state.comparison_groups.keys()):
                group = self.state.comparison_groups[group_id]
                if sap_code in group:
                    group.pop(sap_code, None)
                if not group:
                    self.state.comparison_groups.pop(group_id, None)

        self.notify_observers("test_removed", {"test_id": test_id, "sap_code": sap_code})
        return True

    def remove_tests_for_sap(self, sap_code: str) -> int:
        """Remove all selected tests associated with a SAP code."""
        to_remove = [tid for tid, test in self.state.selected_tests.items() if test.sap_code == sap_code]
        removed_count = 0
        for test_id in to_remove:
            if self.remove_selected_test(test_id):
                removed_count += 1
        return removed_count

    def remove_noise_selection(self, sap_code: str, test_lab: Optional[str] = None) -> bool:
        """Remove noise selection for a SAP or a specific test lab."""
        if sap_code not in self.state.selected_noise_saps and sap_code not in self.state.selected_noise_test_labs:
            return False

        if test_lab is None:
            self.state.selected_noise_saps.discard(sap_code)
            self.state.selected_noise_test_labs.pop(sap_code, None)
            self.notify_observers("noise_selection_removed", {"sap_code": sap_code})
            return True

        tests = self.state.selected_noise_test_labs.get(sap_code)
        if not tests or test_lab not in tests:
            return False

        tests.discard(test_lab)
        if not tests:
            self.state.selected_noise_test_labs.pop(sap_code, None)
            self.state.selected_noise_saps.discard(sap_code)

        self.notify_observers("noise_test_removed", {"sap_code": sap_code, "test_lab": test_lab})
        return True

    def remove_comparison_selection(self, sap_code: str, test_lab: Optional[str] = None) -> bool:
        """Remove comparison selection for a SAP or a specific test lab (legacy system)."""
        if sap_code not in self.state.selected_comparison_saps and sap_code not in self.state.selected_comparison_test_labs:
            return False

        if test_lab is None:
            self.state.selected_comparison_saps.discard(sap_code)
            self.state.selected_comparison_test_labs.pop(sap_code, None)
            self.notify_observers("comparison_selection_removed", {"sap_code": sap_code})
            return True

        tests = self.state.selected_comparison_test_labs.get(sap_code)
        if not tests or test_lab not in tests:
            return False

        tests.discard(test_lab)
        if not tests:
            self.state.selected_comparison_test_labs.pop(sap_code, None)
            self.state.selected_comparison_saps.discard(sap_code)

        self.notify_observers("comparison_test_removed", {"sap_code": sap_code, "test_lab": test_lab})
        return True

    def remove_comparison_group_entry(self, group_id: str, sap_code: Optional[str] = None,
                                       test_lab: Optional[str] = None) -> bool:
        """Remove entries from the new comparison group structure."""
        group = self.state.comparison_groups.get(group_id)
        if group is None:
            return False

        if sap_code is None:
            # Remove entire group
            self.state.comparison_groups.pop(group_id, None)
            self.notify_observers("comparison_group_removed", {"group_id": group_id})
            return True

        if sap_code not in group:
            return False

        if test_lab is None:
            group.pop(sap_code, None)
        else:
            tests = group.get(sap_code)
            if not tests or test_lab not in tests:
                return False
            tests.discard(test_lab)
            if not tests:
                group.pop(sap_code, None)

        if not group:
            self.state.comparison_groups.pop(group_id, None)
            self.notify_observers("comparison_group_removed", {"group_id": group_id})
        else:
            self.notify_observers("comparison_group_updated", {"group_id": group_id, "sap_code": sap_code})

        return True
    
    def apply_search_selection(self):
        """Apply the current test selection and update cache efficiently"""
        # Only update if there's a change
        new_test_labs = set(self.state.selected_tests.keys())
        new_sap_codes = set(t.sap_code for t in self.state.selected_tests.values() if t.sap_code)
        
        if (self.selection_cache.selected_test_labs == new_test_labs and 
            self.selection_cache.selected_sap_codes == new_sap_codes):
            return  # No changes to apply
        
        self.selection_cache.selected_test_labs = new_test_labs
        self.selection_cache.selected_sap_codes = new_sap_codes
        self.state.search_selection_applied = len(self.state.selected_tests) > 0
        
        # Only notify if there are observers
        if self._observers:
            self.notify_observers("search_selection_applied", {
                "selected_count": len(self.state.selected_tests)
            })
    
    def clear_search_selection(self):
        """Clear all selected tests"""
        self.state.selected_tests.clear()
        self.selection_cache.selected_test_labs.clear()
        self.selection_cache.selected_sap_codes.clear()
        self.state.search_selection_applied = False
        self._invalidate_carichi_cache()
        
        self.notify_observers("search_selection_cleared")

    def _remove_noise_test_reference(self, sap_code: Optional[str], test_id: str):
        if not sap_code:
            return

        tests = self.state.selected_noise_test_labs.get(sap_code)
        if tests and test_id in tests:
            tests.discard(test_id)
            if not tests:
                self.state.selected_noise_test_labs.pop(sap_code, None)
                self.state.selected_noise_saps.discard(sap_code)

    def _remove_comparison_test_reference(self, sap_code: Optional[str], test_id: str):
        if sap_code is None:
            return

        tests = self.state.selected_comparison_test_labs.get(sap_code)
        if tests and test_id in tests:
            tests.discard(test_id)
            if not tests:
                self.state.selected_comparison_test_labs.pop(sap_code, None)
                self.state.selected_comparison_saps.discard(sap_code)

        # Handle new comparison group structure
        for group_id in list(self.state.comparison_groups.keys()):
            group = self.state.comparison_groups[group_id]
            tests_for_sap = group.get(sap_code)
            if tests_for_sap and test_id in tests_for_sap:
                tests_for_sap.discard(test_id)
                if not tests_for_sap:
                    group.pop(sap_code, None)
                if not group:
                    self.state.comparison_groups.pop(group_id, None)
    
    def update_sap_selection(self, sap_type: str, sap_code: str, selected: bool):
        """Update SAP selection for a specific type (performance, noise, comparison)"""
        # Validate sap_type
        valid_types = {"performance", "noise", "comparison"}
        if sap_type not in valid_types:
            logger.warning(f"Invalid SAP type: {sap_type}. Valid types: {valid_types}")
            return
        
        try:
            sap_set = getattr(self.state, f"selected_{sap_type}_saps")
        except AttributeError:
            logger.error(f"SAP set not found for type: {sap_type}")
            return
        
        if selected:
            sap_set.add(sap_code)
        else:
            sap_set.discard(sap_code)
        
        self.notify_observers("sap_selection_changed", {
            "sap_type": sap_type,
            "sap_code": sap_code,
            "selected": selected
        })
    
    def update_paths(self, tests_folder: Optional[str] = None, registry_file: Optional[str] = None, 
                     noise_folder: Optional[str] = None, noise_registry: Optional[str] = None,
                     test_lab_dir: Optional[str] = None):
        """Update file paths"""
        if tests_folder is not None:
            self.state.selected_tests_folder = tests_folder
        if registry_file is not None:
            self.state.selected_registry_file = registry_file
        if noise_folder is not None:
            self.state.selected_noise_folder = noise_folder
        if noise_registry is not None:
            self.state.selected_noise_registry = noise_registry
        if test_lab_dir is not None:
            self.state.test_lab_directory = test_lab_dir or ""
            self._carichi_locator = None
            self._carichi_locator_path = None
            self._invalidate_carichi_cache()
        
        self.notify_observers("paths_updated", {
            "tests_folder": self.state.selected_tests_folder,
            "registry_file": self.state.selected_registry_file,
            "noise_folder": self.state.selected_noise_folder,
            "noise_registry": self.state.selected_noise_registry,
            "test_lab_dir": self.state.test_lab_directory
        })
    
    def update_configuration(self, **config):
        """Update configuration settings"""
        for key, value in config.items():
            if hasattr(self.state, key):
                setattr(self.state, key, value)
        
        self.notify_observers("configuration_updated", config)
    
    def get_tests_to_process(self) -> List[Test]:
        """Get the list of tests that should be processed"""
        return list(self.state.selected_tests.values())
    
    def get_unique_saps(self) -> List[str]:
        """Get unique SAP codes from selected tests"""
        return list(set(
            test.sap_code for test in self.state.selected_tests.values() 
            if test.sap_code
        ))
    
    # Multiple Comparisons Management
    def add_comparison_group(self, name: str, test_labs: List[str], description: str = "") -> str:
        """Add a new comparison group and return its ID"""
        comparison_id = f"comp_{len(self.state.multiple_comparisons) + 1}"
        comparison_group = {
            "id": comparison_id,
            "name": name,
            "test_labs": test_labs.copy(),
            "description": description,
            "created_at": self._get_timestamp()
        }
        self.state.multiple_comparisons.append(comparison_group)
        
        self.notify_observers("comparison_added", comparison_group)
        return comparison_id
    
    def remove_comparison_group(self, comparison_id: str) -> bool:
        """Remove a comparison group by ID"""
        for i, comp in enumerate(self.state.multiple_comparisons):
            if comp.get("id") == comparison_id:
                removed_comp = self.state.multiple_comparisons.pop(i)
                self.notify_observers("comparison_removed", removed_comp)
                return True
        return False
    
    def update_comparison_group(self, comparison_id: str, **updates) -> bool:
        """Update a comparison group"""
        for comp in self.state.multiple_comparisons:
            if comp.get("id") == comparison_id:
                comp.update(updates)
                comp["updated_at"] = self._get_timestamp()
                self.notify_observers("comparison_updated", comp)
                return True
        return False
    
    def get_comparison_group(self, comparison_id: str) -> Optional[Dict[str, Any]]:
        """Get a specific comparison group by ID"""
        for comp in self.state.multiple_comparisons:
            if comp.get("id") == comparison_id:
                return comp.copy()
        return None
    
    def get_all_comparison_groups(self) -> List[Dict[str, Any]]:
        """Get all comparison groups"""
        return [comp.copy() for comp in self.state.multiple_comparisons]
    
    def clear_all_comparisons(self):
        """Clear all comparison groups"""
        self.state.multiple_comparisons.clear()
        self.notify_observers("all_comparisons_cleared", {})
    
    def _get_timestamp(self) -> str:
        """Get current timestamp string"""
        import datetime
        return datetime.datetime.now().isoformat()
    
    def load_noise_registry_data(self) -> bool:
        """
        Load the noise registry DataFrame and store it in state.
        Returns True if successful, False otherwise.
        """
        if not self.state.selected_noise_registry:
            logger.warning("No noise registry file selected")
            return False
        
        try:
            from pathlib import Path

            registry_path = Path(self.state.selected_noise_registry)
            if not registry_path.exists():
                logger.error(f"Noise registry file not found: {registry_path}")
                return False

            sheet_option = self.state.noise_registry_sheet_name or None
            try:
                df, column_mapping = load_registry_dataframe(
                    registry_path,
                    sheet_name=sheet_option,
                    log=logger,
                )
            except ValueError as exc:
                logger.warning(
                    "Noise registry sheet '%s' not found (%s); falling back to first sheet",
                    sheet_option,
                    exc,
                )
                df, column_mapping = load_registry_dataframe(
                    registry_path,
                    sheet_name=0,
                    log=logger,
                )

            if column_mapping:
                logger.debug("Noise registry column mapping applied: %s", column_mapping)

            self.state.noise_registry_dataframe = df
            logger.info(
                "Loaded noise registry with %s records from %s",
                len(df),
                registry_path,
            )
            self.notify_observers("noise_registry_loaded", {"record_count": len(df)})
            return True

        except Exception as e:
            logger.error(f"Failed to load noise registry: {e}")
            self.state.noise_registry_dataframe = None
            return False
    
    def get_noise_registry_dataframe(self):
        """Get the loaded noise registry DataFrame"""
        return self.state.noise_registry_dataframe
    
    def is_noise_registry_loaded(self) -> bool:
        """Check if noise registry data is loaded"""
        return self.state.noise_registry_dataframe is not None

    # ------------------------------------------------------------------
    # Carichi Nominali helpers

    def _invalidate_carichi_cache(self):
        """Clear cached Carichi lookup results."""
        self.state.carichi_matches.clear()
        self.state.carichi_last_checked = None
        self.state.carichi_errors.clear()
        self._carichi_lookup_signature = None

    def _build_carichi_signature(self) -> tuple[str, tuple[str, ...]]:
        """Return a signature representing the current lookup inputs."""
        base_path = (self.state.test_lab_directory or "").strip()
        test_ids = tuple(sorted(self.state.selected_tests.keys()))
        return base_path, test_ids

    def _ensure_carichi_locator(self) -> Optional[CarichiLocator]:
        """Instantiate (or reuse) a Carichi locator for the current path."""
        base_path = (self.state.test_lab_directory or "").strip()
        if not base_path:
            self._carichi_locator = None
            self._carichi_locator_path = None
            return None

        normalized = os.path.normpath(base_path)
        if self._carichi_locator_path != normalized:
            self._carichi_locator = CarichiLocator(normalized, log=logger.getChild("CarichiLocator"))
            self._carichi_locator_path = normalized
        return self._carichi_locator

    def refresh_carichi_matches(self, force: bool = False) -> Dict[str, Optional[CarichiTestInfo]]:
        """Refresh cached Carichi lookup results for selected tests."""
        signature = self._build_carichi_signature()
        if (not force and self._carichi_lookup_signature == signature and
                self.state.carichi_matches):
            return self.state.carichi_matches

        if not self.state.selected_tests:
            self.state.carichi_matches = {}
            self.state.carichi_errors = []
            self.state.carichi_last_checked = None
            self._carichi_lookup_signature = signature
            return {}

        base_dir = (self.state.test_lab_directory or "").strip()
        if not base_dir:
            self.state.carichi_matches = {}
            self.state.carichi_errors = ["Test Lab directory not configured."]
            self.state.carichi_last_checked = None
            self._carichi_lookup_signature = None
            return {}

        locator = self._ensure_carichi_locator()
        if not locator or not locator.available:
            message = f"Test Lab directory not accessible: {base_dir}"
            self.state.carichi_matches = {}
            self.state.carichi_errors = [message]
            self.state.carichi_last_checked = self._get_timestamp()
            self._carichi_lookup_signature = signature
            return {}

        matches: Dict[str, Optional[CarichiTestInfo]] = {}
        for test_id, test in self.state.selected_tests.items():
            try:
                matches[test_id] = locator.find_for_performance_test(test.test_lab_number)
            except Exception as exc:
                logger.warning("Carichi lookup failed for %s: %s", test.test_lab_number, exc)
                matches[test_id] = None

        self.state.carichi_matches = matches
        self.state.carichi_errors = []
        self.state.carichi_last_checked = self._get_timestamp()
        self._carichi_lookup_signature = signature

        resolved = sum(1 for match in matches.values() if match)
        missing = len(matches) - resolved
        if self._observers:
            self.notify_observers("carichi_matches_refreshed", {
                "resolved": resolved,
                "missing": missing,
                "path": base_dir,
            })

        return matches

    def get_carichi_status(self) -> Dict[str, Any]:
        """Return a snapshot summary of Carichi lookup coverage."""
        matches = self.state.carichi_matches or {}
        selected_tests = self.state.selected_tests
        
        # Filter: Only consider tests ending in "A" as relevant for Carichi coverage
        relevant_tests = {
            tid: t for tid, t in selected_tests.items() 
            if t.test_lab_number and t.test_lab_number.strip().upper().endswith("A")
        }
        
        total_tests = len(relevant_tests)
        
        # Count resolved only among relevant tests
        resolved = 0
        for tid in relevant_tests:
            if matches.get(tid):
                resolved += 1

        missing_details: List[Dict[str, str]] = []
        missing_by_sap: Dict[str, List[str]] = {}
        year_counts: Dict[str, int] = {}

        for test_id, test in relevant_tests.items():
            match_info = matches.get(test_id)
            if match_info and match_info.year_folder:
                year_counts[match_info.year_folder] = year_counts.get(match_info.year_folder, 0) + 1

            if not match_info:
                sap_code = test.sap_code or "UNKNOWN"
                missing_details.append({
                    "test_number": test.test_lab_number,
                    "sap_code": sap_code,
                })
                missing_by_sap.setdefault(sap_code, []).append(test.test_lab_number)

        coverage_percent = 0.0
        if total_tests:
            coverage_percent = round((resolved / total_tests) * 100, 1)
        elif not relevant_tests and not selected_tests:
             # No tests selected at all
             coverage_percent = 0.0
        elif not relevant_tests and selected_tests:
             # Tests selected, but none are "A" tests -> effectively 100% (nothing missing) or N/A
             # Let's treat it as 100% to avoid showing errors
             coverage_percent = 100.0

        locator = self._ensure_carichi_locator() if self.state.test_lab_directory else None
        available = bool(locator and locator.available)

        return {
            "enabled": bool(self.state.test_lab_directory),
            "path": self.state.test_lab_directory,
            "available": available,
            "total_tests": total_tests,
            "resolved_count": resolved,
            "missing_count": len(missing_details),
            "missing_details": missing_details,
            "missing_by_sap": missing_by_sap,
            "year_counts": year_counts,
            "last_checked": self.state.carichi_last_checked,
            "errors": list(self.state.carichi_errors),
            "coverage_percent": coverage_percent,
            "matches": matches,
        }

