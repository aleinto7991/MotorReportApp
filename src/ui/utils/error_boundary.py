"""
Error Boundary - Graceful error handling and circuit breaker pattern.

This module provides robust error handling for GUI operations:
- Try-catch decorators for UI methods
- Timeout protection for long-running operations
- Circuit breaker pattern for repeated failures
- Fallback UI states

Prevents silent failures that leave UI frozen or unresponsive.
"""
import logging
import time
from functools import wraps
from typing import Callable, Any, Optional, TypeVar
from concurrent.futures import TimeoutError as FutureTimeoutError

logger = logging.getLogger(__name__)

T = TypeVar('T')


class CircuitBreaker:
    """Circuit breaker pattern to prevent cascading failures.
    
    Opens after max_failures consecutive failures, preventing further
    attempts for cooldown_seconds. Useful for operations like search
    or report generation that might repeatedly fail.
    """
    
    def __init__(self, max_failures: int = 3, cooldown_seconds: float = 30.0):
        """Initialize circuit breaker.
        
        Args:
            max_failures: Number of failures before opening circuit
            cooldown_seconds: Time to wait before allowing retry
        """
        self.max_failures = max_failures
        self.cooldown_seconds = cooldown_seconds
        self.failures = 0
        self.last_failure_time = 0
        self.is_open = False
    
    def record_success(self):
        """Record successful operation - resets failure count."""
        self.failures = 0
        self.is_open = False
    
    def record_failure(self):
        """Record failed operation - may open circuit."""
        self.failures += 1
        self.last_failure_time = time.time()
        
        if self.failures >= self.max_failures:
            self.is_open = True
            logger.warning(
                f"Circuit breaker OPEN after {self.failures} failures. "
                f"Cooldown: {self.cooldown_seconds}s"
            )
    
    def can_attempt(self) -> bool:
        """Check if operation can be attempted.
        
        Returns:
            True if circuit is closed or cooldown expired
        """
        if not self.is_open:
            return True
        
        # Check if cooldown period has passed
        time_since_failure = time.time() - self.last_failure_time
        if time_since_failure >= self.cooldown_seconds:
            logger.info("Circuit breaker cooldown expired, allowing retry")
            self.is_open = False
            self.failures = 0
            return True
        
        return False
    
    def get_status(self) -> str:
        """Get human-readable status.
        
        Returns:
            Status string for UI display
        """
        if not self.is_open:
            return "OK"
        
        time_remaining = self.cooldown_seconds - (time.time() - self.last_failure_time)
        return f"Temporary block ({int(time_remaining)}s remaining)"


def with_error_boundary(
    fallback_value: Any = None,
    fallback_ui_message: Optional[str] = None,
    log_level: str = "error"
) -> Callable:
    """Decorator to add error boundary to UI methods.
    
    Catches exceptions and provides graceful degradation instead of
    crashing the UI or leaving it in inconsistent state.
    
    Args:
        fallback_value: Value to return on error
        fallback_ui_message: Message to show user on error
        log_level: Logging level ('error', 'warning', 'debug')
    
    Example:
        @with_error_boundary(fallback_value=[], fallback_ui_message="Failed to load data")
        def load_data(self):
            return risky_operation()
    """
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # Log the error
                log_func = getattr(logger, log_level, logger.error)
                log_func(
                    f"Error in {func.__name__}: {e}",
                    exc_info=log_level == "error"
                )
                
                # Try to show UI message if possible
                if fallback_ui_message:
                    try:
                        # Try to find GUI instance from args
                        for arg in args:
                            if hasattr(arg, 'status_manager'):
                                arg.status_manager.update_status(
                                    f"⚠️ {fallback_ui_message}",
                                    "orange"
                                )
                                break
                    except:
                        pass  # Don't fail showing error message
                
                return fallback_value
        
        return wrapper
    return decorator


def with_timeout(timeout_seconds: float = 30.0, timeout_message: str = "Operation timed out"):
    """Decorator to add timeout protection to operations.
    
    Uses concurrent.futures to enforce timeout on long-running operations.
    
    Args:
        timeout_seconds: Maximum time allowed
        timeout_message: Message to show/log on timeout
    
    Example:
        @with_timeout(timeout_seconds=10.0)
        def slow_operation(self):
            # Long running code
            pass
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            from concurrent.futures import ThreadPoolExecutor
            
            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(func, *args, **kwargs)
                try:
                    return future.result(timeout=timeout_seconds)
                except FutureTimeoutError:
                    logger.error(f"{func.__name__} timed out after {timeout_seconds}s")
                    raise TimeoutError(timeout_message)
        
        return wrapper
    return decorator


def safe_ui_update(func: Callable) -> Callable:
    """Decorator for safe UI updates with fallback.
    
    Wraps UI update methods to prevent crashes from Flet session issues,
    disposed controls, or other UI-related errors.
    
    Example:
        @safe_ui_update
        def update_status(self, message):
            self.status_text.value = message
            self.page.update()
    """
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except (RuntimeError, AttributeError) as e:
            # Common errors when UI is disposed or session closed
            error_msg = str(e).lower()
            if any(word in error_msg for word in ['disposed', 'session', 'closed', 'shutdown']):
                logger.debug(f"UI update skipped (session closed): {func.__name__}")
                return None
            raise  # Re-raise if not a known UI disposal error
        except Exception as e:
            logger.error(f"Unexpected error in UI update {func.__name__}: {e}", exc_info=True)
            return None
    
    return wrapper


# Global circuit breakers for common operations
search_circuit_breaker = CircuitBreaker(max_failures=3, cooldown_seconds=30.0)
report_circuit_breaker = CircuitBreaker(max_failures=2, cooldown_seconds=60.0)
registry_circuit_breaker = CircuitBreaker(max_failures=3, cooldown_seconds=20.0)

