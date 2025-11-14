"""
Unit tests for Excel profiling utilities.
"""
import pytest
import time
from src.reports.excel_profiler import ExcelProfiler, TimingStats, enable_profiling, disable_profiling


class TestTimingStats:
    """Test cases for TimingStats."""
    
    def test_initialization(self):
        """Test timing stats initialization."""
        stats = TimingStats("test_operation")
        assert stats.operation_name == "test_operation"
        assert stats.call_count == 0
        assert stats.total_time == 0.0
        assert stats.min_time == float('inf')
        assert stats.max_time == 0.0
    
    def test_record_timing(self):
        """Test recording timing measurements."""
        stats = TimingStats("test_op")
        
        stats.record(0.5)
        assert stats.call_count == 1
        assert stats.total_time == 0.5
        assert stats.min_time == 0.5
        assert stats.max_time == 0.5
        
        stats.record(0.3)
        assert stats.call_count == 2
        assert stats.total_time == 0.8
        assert stats.min_time == 0.3
        assert stats.max_time == 0.5
    
    def test_average_calculation(self):
        """Test average time calculation."""
        stats = TimingStats("test_op")
        
        stats.record(0.5)
        stats.record(0.3)
        stats.record(0.4)
        
        assert stats.avg_time == pytest.approx(0.4, rel=0.01)


class TestExcelProfiler:
    """Test cases for ExcelProfiler."""
    
    def test_initialization(self):
        """Test profiler initialization."""
        profiler = ExcelProfiler("Test Session")
        assert profiler.name == "Test Session"
        assert len(profiler.stats) == 0
        assert profiler.session_start_time is None
        assert profiler.session_end_time is None
    
    def test_time_operation_context_manager(self):
        """Test timing operations with context manager."""
        profiler = ExcelProfiler()
        
        with profiler.time_operation("test_operation"):
            time.sleep(0.05)
        
        assert "test_operation" in profiler.stats
        assert profiler.stats["test_operation"].call_count == 1
        assert profiler.stats["test_operation"].total_time >= 0.05
    
    def test_multiple_operations(self):
        """Test tracking multiple different operations."""
        profiler = ExcelProfiler()
        
        with profiler.time_operation("operation_1"):
            time.sleep(0.02)
        
        with profiler.time_operation("operation_2"):
            time.sleep(0.01)
        
        with profiler.time_operation("operation_1"):  # Same operation again
            time.sleep(0.02)
        
        assert len(profiler.stats) == 2
        assert profiler.stats["operation_1"].call_count == 2
        assert profiler.stats["operation_2"].call_count == 1
    
    def test_profile_method_decorator(self):
        """Test profiling methods with decorator."""
        profiler = ExcelProfiler()
        
        @profiler.profile_method("decorated_method")
        def test_function():
            time.sleep(0.01)
            return "result"
        
        result = test_function()
        
        assert result == "result"
        assert "decorated_method" in profiler.stats
        assert profiler.stats["decorated_method"].call_count == 1
    
    def test_session_tracking(self):
        """Test session start and end tracking."""
        profiler = ExcelProfiler()
        
        profiler.start_session()
        assert profiler.session_start_time is not None
        
        time.sleep(0.05)
        
        profiler.end_session()
        assert profiler.session_end_time is not None
        assert profiler.session_duration >= 0.05
    
    def test_total_measured_time(self):
        """Test calculation of total measured time."""
        profiler = ExcelProfiler()
        
        with profiler.time_operation("op1"):
            time.sleep(0.02)
        
        with profiler.time_operation("op2"):
            time.sleep(0.03)
        
        total = profiler.total_measured_time
        assert total >= 0.05
        assert total < 0.1  # Sanity check
    
    def test_get_sorted_stats(self):
        """Test getting stats sorted by different metrics."""
        profiler = ExcelProfiler()
        
        # Create operations with different characteristics
        with profiler.time_operation("slow_op"):
            time.sleep(0.03)
        
        with profiler.time_operation("fast_op"):
            time.sleep(0.01)
        
        with profiler.time_operation("fast_op"):  # Call twice for higher count
            time.sleep(0.01)
        
        # Sort by total time
        sorted_by_time = profiler.get_sorted_stats('total_time')
        assert sorted_by_time[0].operation_name == "slow_op"
        
        # Sort by call count
        sorted_by_count = profiler.get_sorted_stats('call_count')
        assert sorted_by_count[0].operation_name == "fast_op"
    
    def test_get_bottlenecks(self):
        """Test bottleneck identification."""
        profiler = ExcelProfiler()
        
        # Create operations with different time proportions
        with profiler.time_operation("major_op"):
            time.sleep(0.09)  # ~90% of time
        
        with profiler.time_operation("minor_op"):
            time.sleep(0.005)  # ~5% of time
        
        # Bottleneck threshold of 15% - should identify major_op only
        bottlenecks = profiler.get_bottlenecks(threshold_percent=15.0)
        assert len(bottlenecks) >= 1
        assert bottlenecks[0].operation_name == "major_op"
        
        # Lower threshold should identify both
        bottlenecks_all = profiler.get_bottlenecks(threshold_percent=3.0)
        assert len(bottlenecks_all) == 2
    
    def test_reset(self):
        """Test resetting profiler data."""
        profiler = ExcelProfiler()
        
        profiler.start_session()
        with profiler.time_operation("test"):
            time.sleep(0.01)
        profiler.end_session()
        
        assert len(profiler.stats) > 0
        assert profiler.session_start_time is not None
        
        profiler.reset()
        
        assert len(profiler.stats) == 0
        assert profiler.session_start_time is None
        assert profiler.session_end_time is None
    
    def test_nested_operations(self):
        """Test profiling nested operations."""
        profiler = ExcelProfiler()
        
        with profiler.time_operation("outer"):
            time.sleep(0.01)
            with profiler.time_operation("inner"):
                time.sleep(0.02)
            time.sleep(0.01)
        
        assert "outer" in profiler.stats
        assert "inner" in profiler.stats
        # Outer should take longer than inner
        assert profiler.stats["outer"].total_time > profiler.stats["inner"].total_time


class TestGlobalProfiler:
    """Test cases for global profiler functions."""
    
    def test_enable_profiling(self):
        """Test enabling profiling."""
        profiler = enable_profiling("Test Session")
        assert profiler is not None
        assert profiler.name == "Test Session"
    
    def test_disable_profiling(self):
        """Test disabling profiling."""
        enable_profiling()
        disable_profiling()
        # Should not raise errors


class TestProfilerRecommendations:
    """Test profiler's recommendation engine."""
    
    def test_slow_operation_recommendations(self):
        """Test recommendations for slow operations."""
        profiler = ExcelProfiler()
        
        # Create a very slow operation (>20% of total time)
        with profiler.time_operation("slow_operation"):
            time.sleep(0.09)
        
        with profiler.time_operation("fast_operation"):
            time.sleep(0.01)
        
        sorted_stats = profiler.get_sorted_stats('total_time')
        recommendations = profiler._generate_recommendations(sorted_stats)
        
        assert len(recommendations) > 0
        # Should recommend optimizing the slow operation
        assert any("slow_operation" in rec for rec in recommendations)
    
    def test_repeated_operation_recommendations(self):
        """Test recommendations for frequently called operations."""
        profiler = ExcelProfiler()
        
        # Call an operation many times
        for _ in range(150):
            with profiler.time_operation("repeated_op"):
                time.sleep(0.002)  # 2ms each
        
        sorted_stats = profiler.get_sorted_stats('total_time')
        recommendations = profiler._generate_recommendations(sorted_stats)
        
        # Should recommend batching
        assert any("batch" in rec.lower() for rec in recommendations)
