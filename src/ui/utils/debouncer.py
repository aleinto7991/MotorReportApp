"""
Debouncer - Prevents excessive function calls from rapid-fire events.

This module provides debouncing utilities to reduce unnecessary UI updates
and backend calls when users interact rapidly with the interface.

Common use cases:
- Checkbox toggles during bulk selections
- Filter changes (SAP selection, test selection)
- Search input
- Slider movements

Benefits:
- Reduces UI render cycles
- Prevents backend overload
- Smoother user experience
"""
import logging
import threading
import time
from functools import wraps
from typing import Callable, Any, Optional, Dict

logger = logging.getLogger(__name__)


class Debouncer:
    """Debounces function calls - only executes after quiet period.
    
    When a debounced function is called repeatedly, it waits for a
    quiet period (no calls for delay_seconds) before executing.
    
    Perfect for handling rapid-fire events like checkboxes or filters.
    """
    
    def __init__(self, delay_seconds: float = 0.3, name: str = "debouncer"):
        """Initialize debouncer.
        
        Args:
            delay_seconds: Quiet period required before execution
            name: Name for logging purposes
        """
        self.delay_seconds = delay_seconds
        self.name = name
        self._timer: Optional[threading.Timer] = None
        self._lock = threading.Lock()
        self._pending_args = None
        self._pending_kwargs = None
    
    def debounce(self, func: Callable) -> Callable:
        """Decorator to debounce a function.
        
        Args:
            func: Function to debounce
        
        Returns:
            Debounced function
        
        Example:
            debouncer = Debouncer(delay_seconds=0.5)
            
            @debouncer.debounce
            def on_filter_changed(self, value):
                # This only executes 0.5s after the last call
                self.update_results(value)
        """
        @wraps(func)
        def wrapper(*args, **kwargs):
            with self._lock:
                # Cancel pending execution
                if self._timer is not None:
                    self._timer.cancel()
                    logger.debug(f"Debouncer '{self.name}': Cancelled pending call")
                
                # Store latest args
                self._pending_args = args
                self._pending_kwargs = kwargs
                
                # Schedule new execution
                def execute():
                    logger.debug(f"Debouncer '{self.name}': Executing after {self.delay_seconds}s delay")
                    try:
                        func(*self._pending_args, **self._pending_kwargs)
                    except Exception as e:
                        logger.error(f"Debouncer '{self.name}': Error in debounced function: {e}", exc_info=True)
                
                self._timer = threading.Timer(self.delay_seconds, execute)
                self._timer.daemon = True
                self._timer.start()
                logger.debug(f"Debouncer '{self.name}': Scheduled execution in {self.delay_seconds}s")
        
        return wrapper
    
    def cancel(self):
        """Cancel any pending execution."""
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
                self._timer = None
                logger.debug(f"Debouncer '{self.name}': Cancelled")


class Throttler:
    """Throttles function calls - executes at most once per interval.
    
    Unlike debouncing, throttling guarantees execution at regular intervals
    even during continuous activity. First call executes immediately,
    subsequent calls wait for interval to pass.
    
    Perfect for expensive operations that must show progress.
    """
    
    def __init__(self, interval_seconds: float = 0.5, name: str = "throttler"):
        """Initialize throttler.
        
        Args:
            interval_seconds: Minimum time between executions
            name: Name for logging purposes
        """
        self.interval_seconds = interval_seconds
        self.name = name
        self._last_execution_time = 0
        self._lock = threading.Lock()
    
    def throttle(self, func: Callable) -> Callable:
        """Decorator to throttle a function.
        
        Args:
            func: Function to throttle
        
        Returns:
            Throttled function
        
        Example:
            throttler = Throttler(interval_seconds=1.0)
            
            @throttler.throttle
            def on_scroll(self, position):
                # Executes at most once per second
                self.load_more_items(position)
        """
        @wraps(func)
        def wrapper(*args, **kwargs):
            with self._lock:
                current_time = time.time()
                time_since_last = current_time - self._last_execution_time
                
                if time_since_last >= self.interval_seconds:
                    logger.debug(f"Throttler '{self.name}': Executing (last was {time_since_last:.2f}s ago)")
                    self._last_execution_time = current_time
                    return func(*args, **kwargs)
                else:
                    logger.debug(
                        f"Throttler '{self.name}': Skipped (only {time_since_last:.2f}s since last, "
                        f"need {self.interval_seconds}s)"
                    )
                    return None
        
        return wrapper


# Global debouncers for common UI operations
filter_debouncer = Debouncer(delay_seconds=0.25, name="filter_changes")
selection_debouncer = Debouncer(delay_seconds=0.3, name="checkbox_selection")
sap_selection_debouncer = Debouncer(delay_seconds=0.2, name="sap_selection")
search_input_debouncer = Debouncer(delay_seconds=0.4, name="search_input")


def debounce(delay_seconds: float = 0.3, debouncer_name: str = "custom"):
    """Function decorator for one-off debouncing.
    
    Creates a new debouncer for a specific function.
    
    Args:
        delay_seconds: Quiet period required before execution
        debouncer_name: Name for logging
    
    Example:
        @debounce(delay_seconds=0.5, debouncer_name="custom_filter")
        def my_expensive_function(self, value):
            # Heavy computation here
            pass
    """
    debouncer = Debouncer(delay_seconds=delay_seconds, name=debouncer_name)
    return debouncer.debounce


def throttle(interval_seconds: float = 0.5, throttler_name: str = "custom"):
    """Function decorator for one-off throttling.
    
    Creates a new throttler for a specific function.
    
    Args:
        interval_seconds: Minimum time between executions
        throttler_name: Name for logging
    
    Example:
        @throttle(interval_seconds=1.0, throttler_name="scroll_handler")
        def on_scroll_event(self, position):
            # Load more items
            pass
    """
    throttler = Throttler(interval_seconds=interval_seconds, name=throttler_name)
    return throttler.throttle

