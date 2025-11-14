"""
Thread Pool Manager - Centralized thread pool for background tasks.

This module provides a shared ThreadPoolExecutor for all background operations
in the GUI, replacing ad-hoc threading.Thread creations.

Benefits:
- Controlled resource usage (max_workers=4)
- Automatic thread reuse
- Better error handling and logging
- Graceful shutdown
- Future-based async patterns
"""
import logging
import threading
from concurrent.futures import ThreadPoolExecutor, Future
from typing import Callable, Any, Optional
from functools import wraps

logger = logging.getLogger(__name__)

# Shared thread pool for all GUI background tasks
_thread_pool: Optional[ThreadPoolExecutor] = None
_pool_lock = threading.Lock()


def get_thread_pool() -> ThreadPoolExecutor:
    """Get or create the shared thread pool.
    
    Returns:
        ThreadPoolExecutor: Shared thread pool instance
    """
    global _thread_pool
    
    if _thread_pool is None:
        with _pool_lock:
            if _thread_pool is None:
                _thread_pool = ThreadPoolExecutor(
                    max_workers=4,
                    thread_name_prefix="gui_worker"
                )
                logger.info("Initialized thread pool with 4 workers")
    
    return _thread_pool


def shutdown_thread_pool(wait: bool = True) -> None:
    """Shutdown the thread pool gracefully.
    
    Args:
        wait: If True, wait for all tasks to complete
    """
    global _thread_pool
    
    if _thread_pool is not None:
        with _pool_lock:
            if _thread_pool is not None:
                logger.info("Shutting down thread pool...")
                _thread_pool.shutdown(wait=wait)
                _thread_pool = None
                logger.info("Thread pool shutdown complete")


def submit_task(func: Callable, *args, **kwargs) -> Future:
    """Submit a task to the thread pool.
    
    Args:
        func: Function to execute
        *args: Positional arguments for func
        **kwargs: Keyword arguments for func
    
    Returns:
        Future: Future object for the task
    
    Example:
        future = submit_task(my_function, arg1, arg2, kwarg=value)
        result = future.result(timeout=5)  # Wait up to 5 seconds
    """
    pool = get_thread_pool()
    
    # Wrap function to log errors
    @wraps(func)
    def wrapped():
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Error in background task {func.__name__}: {e}", exc_info=True)
            raise
    
    return pool.submit(wrapped)


def run_in_background(func: Callable, *args, on_complete: Optional[Callable] = None, **kwargs) -> Future:
    """Run a function in the background with optional completion callback.
    
    Args:
        func: Function to execute in background
        *args: Positional arguments for func
        on_complete: Callback to run when task completes (receives result or exception)
        **kwargs: Keyword arguments for func
    
    Returns:
        Future: Future object for the task
    
    Example:
        def on_done(future):
            try:
                result = future.result()
                print(f"Success: {result}")
            except Exception as e:
                print(f"Error: {e}")
        
        run_in_background(my_function, arg1, on_complete=on_done)
    """
    future = submit_task(func, *args, **kwargs)
    
    if on_complete:
        future.add_done_callback(on_complete)
    
    return future

