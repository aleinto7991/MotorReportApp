"""Core business logic for Motor Report Application.

Exports only existing telemetry helpers (log_duration). Memory usage helper
was removed during refactor.
"""

from .motor_report_engine import MotorReportApp
from .telemetry import log_duration

__all__ = [
    "MotorReportApp",
    "log_duration",
]
