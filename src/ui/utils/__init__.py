"""
Utility modules for the GUI

Note: StatusManager has been moved to src/ui/core/status_manager.py for
architectural consistency with other managers.
Note: WorkflowManager has been moved to src/ui/core/workflow_manager.py
"""
from .helpers import SearchResultsFormatter
# For backward compatibility, re-export StatusManager and WorkflowManager from new locations
from ..core.status_manager import StatusManager
from ..core.workflow_manager import WorkflowManager

__all__ = ['StatusManager', 'WorkflowManager', 'SearchResultsFormatter']

