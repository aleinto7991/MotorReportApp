"""Lightweight timing helpers for instrumentation.

Two complementary context managers are provided:

* ``time_scope`` – emits DEBUG messages at start/finish, useful when callers
  only care about timings during local diagnostics.
* ``log_duration`` – emits a single log entry at a configurable level once the
  block completes, making it safe for production telemetry.

Environment Variables:
* ``MOTOR_REPORT_PROFILE`` – When set to "1", "true", or "yes", enables 
  detailed profiling output for performance analysis. When disabled, only
  INFO-level and above messages are logged.
"""
from __future__ import annotations

import logging
import os
import time
from contextlib import contextmanager
from typing import Iterator, Optional, Union


def is_profiling_enabled() -> bool:
    """
    Check if profiling is enabled via environment variable.
    
    Returns:
        True if MOTOR_REPORT_PROFILE is set to "1", "true", or "yes" (case-insensitive)
    """
    value = os.environ.get("MOTOR_REPORT_PROFILE", "").lower()
    return value in ("1", "true", "yes")


@contextmanager
def time_scope(logger: Optional[logging.Logger], label: str) -> Iterator[None]:
    """
    Log the runtime of a code block using debug-level start/finish lines.
    
    Only logs if profiling is enabled or logger is at DEBUG level.
    
    Args:
        logger: Logger instance to use, or None to disable logging
        label: Descriptive label for the timed operation
    
    Yields:
        None
    
    Example:
        >>> with time_scope(logger, "process_data"):
        ...     # expensive operation
        ...     pass
    """
    if logger is None:
        yield
        return
    
    # Only log if profiling enabled or logger is at DEBUG level
    should_log = is_profiling_enabled() or logger.isEnabledFor(logging.DEBUG)
    if not should_log:
        yield
        return

    start = time.perf_counter()
    logger.debug("[timing] %s started", label)
    try:
        yield
    finally:
        elapsed_ms = (time.perf_counter() - start) * 1000
        logger.debug("[timing] %s finished in %.2f ms", label, elapsed_ms)


@contextmanager
def log_duration(
    logger: Union[logging.Logger, logging.LoggerAdapter],
    message: str,
    *,
    level: int = logging.INFO,
    extra: Optional[dict] = None,
) -> Iterator[None]:
    """
    Measure the time spent inside a block and log a single completion line.
    
    Respects profiling environment variable - when profiling is disabled,
    only logs at INFO level and above.
    
    Args:
        logger: Logger or LoggerAdapter instance
        message: Message to log (duration will be appended)
        level: Log level (default: logging.INFO)
        extra: Optional dict of extra logging context
    
    Yields:
        None
    
    Example:
        >>> with log_duration(logger, "Report generation"):
        ...     generate_report()
        # Logs: "Report generation completed in 2.345s"
    """
    # Check if we should log based on profiling setting and log level
    profiling_enabled = is_profiling_enabled()
    should_log = profiling_enabled or level >= logging.INFO
    
    if not should_log:
        yield
        return

    start = time.perf_counter()
    try:
        yield
    finally:
        duration = time.perf_counter() - start
        logger.log(level, "%s completed in %.3fs", message, duration, extra=extra)
