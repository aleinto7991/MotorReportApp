"""Top-level package for MotorReportApp.

Provide backward-compatible top-level exports expected by tests and tools
that import submodules directly from the `src` package (for example
`from src import directory_config`).
"""

# Expose commonly-used submodules at the package level for convenience
from .config import directory_config as directory_config

__all__ = ["directory_config"]
