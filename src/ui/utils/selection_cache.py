"""
Cache for storing SAP code and test lab selections in the GUI workflow.
"""
from typing import List, Dict, Set

class SelectionCache:
    """
    Stores user selections for SAP codes and test lab numbers across steps.
    This cache is used by the GUI to pass the correct data to the backend.
    """
    def __init__(self):
        # All test lab numbers selected by the user (step 2)
        self.selected_test_labs: Set[str] = set()
        # All SAP codes corresponding to selected tests (step 2)
        self.selected_sap_codes: Set[str] = set()
        # SAP codes selected for each feature in step 3
        self.performance_saps: Set[str] = set()
        self.noise_saps: Set[str] = set()
        self.comparison_saps: Set[str] = set()

    def clear(self):
        self.selected_test_labs.clear()
        self.selected_sap_codes.clear()
        self.performance_saps.clear()
        self.noise_saps.clear()
        self.comparison_saps.clear()

    def to_backend_args(self):
        """
        Returns a dict of arguments to pass to the backend for report generation.
        """
        return {
            'selected_test_labs': list(self.selected_test_labs),
            'selected_sap_codes': list(self.selected_sap_codes),
            'performance_saps': list(self.performance_saps),
            'noise_saps': list(self.noise_saps),
            'comparison_saps': list(self.comparison_saps),
        }

