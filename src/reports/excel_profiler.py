"""
Performance profiling utilities for Excel report generation.

Provides timing instrumentation, bottleneck identification, and optimization
recommendations for xlsxwriter operations.
"""
import logging
import time
import functools
from typing import Dict, List, Optional, Callable, Any
from contextlib import contextmanager
from dataclasses import dataclass, field
from collections import defaultdict

logger = logging.getLogger(__name__)


@dataclass
class TimingStats:
    """Statistics for a timed operation."""
    operation_name: str
    call_count: int = 0
    total_time: float = 0.0
    min_time: float = float('inf')
    max_time: float = 0.0
    times: List[float] = field(default_factory=list)
    
    def record(self, elapsed: float) -> None:
        """Record a new timing measurement."""
        self.call_count += 1
        self.total_time += elapsed
        self.min_time = min(self.min_time, elapsed)
        self.max_time = max(self.max_time, elapsed)
        self.times.append(elapsed)
    
    @property
    def avg_time(self) -> float:
        """Calculate average time."""
        return self.total_time / self.call_count if self.call_count > 0 else 0.0
    
    @property
    def percentage_of_total(self) -> float:
        """Calculate percentage of total time (set externally)."""
        return 0.0  # Will be calculated by profiler
    
    def __str__(self) -> str:
        """Format timing stats as string."""
        return (
            f"{self.operation_name}: "
            f"{self.call_count} calls, "
            f"total={self.total_time:.3f}s, "
            f"avg={self.avg_time:.3f}s, "
            f"min={self.min_time:.3f}s, "
            f"max={self.max_time:.3f}s"
        )


class ExcelProfiler:
    """
    Performance profiler for Excel generation operations.
    
    Tracks timing for different operations and provides analysis of bottlenecks.
    
    Usage:
        profiler = ExcelProfiler()
        
        with profiler.time_operation("create_summary_sheet"):
            # ... sheet creation code ...
        
        # Or use as decorator
        @profiler.profile_method("build_sap_sheet")
        def build_sap_sheet(self):
            # ... method code ...
        
        # Get report
        profiler.print_report()
    """
    
    def __init__(self, name: str = "Excel Generation"):
        """
        Initialize profiler.
        
        Args:
            name: Name of the profiling session
        """
        self.name = name
        self.stats: Dict[str, TimingStats] = defaultdict(lambda: TimingStats("unknown"))
        self.session_start_time: Optional[float] = None
        self.session_end_time: Optional[float] = None
        self._operation_stack: List[tuple] = []  # Stack for nested operations
    
    @contextmanager
    def time_operation(self, operation_name: str):
        """
        Context manager to time an operation.
        
        Args:
            operation_name: Name of the operation being timed
        
        Example:
            with profiler.time_operation("write_dataframe"):
                df.to_excel(...)
        """
        start_time = time.perf_counter()
        self._operation_stack.append((operation_name, start_time))
        
        try:
            yield
        finally:
            end_time = time.perf_counter()
            elapsed = end_time - start_time
            
            self._operation_stack.pop()
            
            # Record the timing
            if operation_name not in self.stats:
                self.stats[operation_name] = TimingStats(operation_name)
            
            self.stats[operation_name].record(elapsed)
            
            logger.debug(f"[Profile] {operation_name}: {elapsed:.3f}s")
    
    def profile_method(self, operation_name: Optional[str] = None) -> Callable:
        """
        Decorator to profile a method.
        
        Args:
            operation_name: Optional name for the operation (defaults to method name)
        
        Example:
            @profiler.profile_method()
            def build_sheet(self):
                # ... method code ...
        """
        def decorator(func: Callable) -> Callable:
            name = operation_name or func.__name__
            
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                with self.time_operation(name):
                    return func(*args, **kwargs)
            return wrapper
        return decorator
    
    def start_session(self) -> None:
        """Mark the start of a profiling session."""
        self.session_start_time = time.perf_counter()
        logger.info(f"[Profile] Session started: {self.name}")
    
    def end_session(self) -> None:
        """Mark the end of a profiling session."""
        self.session_end_time = time.perf_counter()
        if self.session_start_time:
            total_time = self.session_end_time - self.session_start_time
            logger.info(f"[Profile] Session ended: {self.name} ({total_time:.3f}s)")
    
    @property
    def total_measured_time(self) -> float:
        """Get total time across all measured operations."""
        return sum(stat.total_time for stat in self.stats.values())
    
    @property
    def session_duration(self) -> Optional[float]:
        """Get total session duration if session was tracked."""
        if self.session_start_time and self.session_end_time:
            return self.session_end_time - self.session_start_time
        return None
    
    def get_sorted_stats(self, by: str = 'total_time') -> List[TimingStats]:
        """
        Get timing stats sorted by a metric.
        
        Args:
            by: Metric to sort by ('total_time', 'avg_time', 'call_count', 'max_time')
        
        Returns:
            List of TimingStats sorted by the specified metric
        """
        sort_keys = {
            'total_time': lambda s: s.total_time,
            'avg_time': lambda s: s.avg_time,
            'call_count': lambda s: s.call_count,
            'max_time': lambda s: s.max_time
        }
        
        sort_key = sort_keys.get(by, sort_keys['total_time'])
        return sorted(self.stats.values(), key=sort_key, reverse=True)
    
    def get_bottlenecks(self, threshold_percent: float = 10.0) -> List[TimingStats]:
        """
        Identify operations that take significant time (bottlenecks).
        
        Args:
            threshold_percent: Minimum percentage of total time to be considered a bottleneck
        
        Returns:
            List of operations exceeding the threshold
        """
        total = self.total_measured_time
        if total == 0:
            return []
        
        bottlenecks = []
        for stat in self.stats.values():
            percentage = (stat.total_time / total) * 100
            if percentage >= threshold_percent:
                bottlenecks.append(stat)
        
        return sorted(bottlenecks, key=lambda s: s.total_time, reverse=True)
    
    def print_report(self, top_n: int = 10, show_all: bool = False) -> None:
        """
        Print a formatted profiling report.
        
        Args:
            top_n: Number of top operations to show
            show_all: If True, show all operations instead of just top N
        """
        print("\n" + "=" * 80)
        print(f"Excel Generation Performance Report: {self.name}")
        print("=" * 80)
        
        if self.session_duration:
            print(f"\nTotal Session Duration: {self.session_duration:.3f}s")
        
        total_measured = self.total_measured_time
        print(f"Total Measured Time: {total_measured:.3f}s")
        
        if self.session_duration and total_measured > 0:
            overhead = self.session_duration - total_measured
            overhead_pct = (overhead / self.session_duration) * 100
            print(f"Unmeasured Overhead: {overhead:.3f}s ({overhead_pct:.1f}%)")
        
        print(f"\nTotal Operations Tracked: {len(self.stats)}")
        
        # Top operations by total time
        print(f"\n{'=' * 80}")
        print(f"Top {top_n if not show_all else 'All'} Operations by Total Time:")
        print(f"{'=' * 80}")
        print(f"{'Operation':<40} {'Calls':>7} {'Total':>10} {'Avg':>10} {'%':>7}")
        print("-" * 80)
        
        sorted_stats = self.get_sorted_stats('total_time')
        stats_to_show = sorted_stats if show_all else sorted_stats[:top_n]
        
        for stat in stats_to_show:
            percentage = (stat.total_time / total_measured * 100) if total_measured > 0 else 0
            print(
                f"{stat.operation_name:<40} "
                f"{stat.call_count:>7} "
                f"{stat.total_time:>9.3f}s "
                f"{stat.avg_time:>9.3f}s "
                f"{percentage:>6.1f}%"
            )
        
        # Bottleneck analysis
        bottlenecks = self.get_bottlenecks(threshold_percent=10.0)
        if bottlenecks:
            print(f"\n{'=' * 80}")
            print("Bottlenecks (operations taking >10% of total time):")
            print(f"{'=' * 80}")
            for stat in bottlenecks:
                percentage = (stat.total_time / total_measured * 100) if total_measured > 0 else 0
                print(f"  ⚠️  {stat.operation_name}: {stat.total_time:.3f}s ({percentage:.1f}%)")
        
        # Recommendations
        print(f"\n{'=' * 80}")
        print("Optimization Recommendations:")
        print(f"{'=' * 80}")
        
        recommendations = self._generate_recommendations(sorted_stats)
        for i, rec in enumerate(recommendations, 1):
            print(f"{i}. {rec}")
        
        print("=" * 80 + "\n")
    
    def _generate_recommendations(self, sorted_stats: List[TimingStats]) -> List[str]:
        """Generate optimization recommendations based on profiling data."""
        recommendations = []
        
        if not sorted_stats:
            return ["No operations tracked. Enable profiling to get recommendations."]
        
        # Check for slow operations
        total = self.total_measured_time
        for stat in sorted_stats[:3]:  # Top 3 slowest
            percentage = (stat.total_time / total * 100) if total > 0 else 0
            if percentage > 20:
                recommendations.append(
                    f"Optimize '{stat.operation_name}' - takes {percentage:.1f}% of total time "
                    f"({stat.total_time:.3f}s)"
                )
        
        # Check for repeated operations
        for stat in sorted_stats:
            if stat.call_count > 100 and stat.avg_time > 0.001:
                recommendations.append(
                    f"Consider batching '{stat.operation_name}' - called {stat.call_count} times "
                    f"(avg {stat.avg_time*1000:.1f}ms per call)"
                )
        
        # Check for xlsx-specific optimizations
        xlsx_operations = [s for s in sorted_stats if 'write' in s.operation_name.lower() or 'format' in s.operation_name.lower()]
        if xlsx_operations:
            recommendations.append(
                "Consider using xlsxwriter's constant_memory mode for large datasets"
            )
        
        if not recommendations:
            recommendations.append("Performance looks good! No critical bottlenecks detected.")
        
        return recommendations
    
    def reset(self) -> None:
        """Reset all profiling data."""
        self.stats.clear()
        self.session_start_time = None
        self.session_end_time = None
        self._operation_stack.clear()
        logger.info(f"[Profile] Reset: {self.name}")


# Global profiler instance for convenience
_global_profiler: Optional[ExcelProfiler] = None


def get_global_profiler() -> ExcelProfiler:
    """Get or create the global profiler instance."""
    global _global_profiler
    if _global_profiler is None:
        _global_profiler = ExcelProfiler("Global Excel Profiler")
    return _global_profiler


def enable_profiling(name: str = "Excel Generation") -> ExcelProfiler:
    """
    Enable profiling and return profiler instance.
    
    Args:
        name: Name for the profiling session
    
    Returns:
        ExcelProfiler instance
    """
    global _global_profiler
    _global_profiler = ExcelProfiler(name)
    logger.info(f"Excel profiling enabled: {name}")
    return _global_profiler


def disable_profiling() -> None:
    """Disable profiling."""
    global _global_profiler
    _global_profiler = None
    logger.info("Excel profiling disabled")
