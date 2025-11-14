"""
Unit tests for loading state components.
"""
import pytest
from src.ui.utils.loading_states import (
    SkeletonLoader,
    ProgressIndicator,
    LoadingState
)
import time


class TestSkeletonLoader:
    """Test cases for SkeletonLoader."""
    
    def test_skeleton_line_creation(self):
        """Test skeleton line component creation."""
        line = SkeletonLoader.skeleton_line(width=200, height=12)
        assert line is not None
        assert line.width == 200
        assert line.height == 12
    
    def test_skeleton_row_creation(self):
        """Test skeleton row component creation."""
        row = SkeletonLoader.skeleton_row()
        assert row is not None
        # Row should have content
        assert row.content is not None
    
    def test_search_results_skeleton(self):
        """Test search results skeleton with multiple rows."""
        skeleton = SkeletonLoader.search_results_skeleton(num_rows=5)
        assert skeleton is not None
        # Should have header + column headers + 5 rows = 7 controls
        assert len(skeleton.controls) == 7  # header, column headers, 5 rows
    
    def test_search_results_skeleton_variable_rows(self):
        """Test skeleton with different row counts."""
        skeleton3 = SkeletonLoader.search_results_skeleton(num_rows=3)
        skeleton10 = SkeletonLoader.search_results_skeleton(num_rows=10)
        
        # header + column headers + n rows
        assert len(skeleton3.controls) == 5  # 2 + 3
        assert len(skeleton10.controls) == 12  # 2 + 10


class TestProgressIndicator:
    """Test cases for ProgressIndicator."""
    
    def test_initialization(self):
        """Test progress indicator initialization."""
        progress = ProgressIndicator(total_steps=100, operation_name="Test")
        assert progress.total_steps == 100
        assert progress.current_step == 0
        assert progress.operation_name == "Test"
        assert progress.percentage == 0
    
    def test_update_progress(self):
        """Test updating progress to specific step."""
        progress = ProgressIndicator(total_steps=100)
        progress.update(50)
        assert progress.current_step == 50
        assert progress.percentage == 50
    
    def test_increment_progress(self):
        """Test incrementing progress."""
        progress = ProgressIndicator(total_steps=100)
        progress.increment(10)
        assert progress.current_step == 10
        assert progress.percentage == 10
        
        progress.increment(25)
        assert progress.current_step == 35
        assert progress.percentage == 35
    
    def test_percentage_calculation(self):
        """Test percentage calculation at various steps."""
        progress = ProgressIndicator(total_steps=200)
        
        progress.update(50)
        assert progress.percentage == 25  # 50/200 = 25%
        
        progress.update(100)
        assert progress.percentage == 50  # 100/200 = 50%
        
        progress.update(200)
        assert progress.percentage == 100  # 200/200 = 100%
    
    def test_progress_clamping(self):
        """Test that progress doesn't exceed total steps."""
        progress = ProgressIndicator(total_steps=100)
        
        # Try to set beyond total
        progress.update(150)
        assert progress.current_step == 100
        assert progress.percentage == 100
        
        # Try to increment beyond total
        progress.update(90)
        progress.increment(20)
        assert progress.current_step == 100
        assert progress.percentage == 100
    
    def test_elapsed_time(self):
        """Test elapsed time tracking."""
        progress = ProgressIndicator(total_steps=100)
        time.sleep(0.1)  # Wait a bit
        elapsed = progress.elapsed_time
        assert elapsed >= 0.1
        assert elapsed < 0.2  # Should be close to 0.1
    
    def test_estimated_remaining_time(self):
        """Test estimated remaining time calculation."""
        progress = ProgressIndicator(total_steps=100)
        
        # At 0%, no estimate
        assert progress.estimated_remaining is None
        
        # Simulate some progress
        time.sleep(0.05)
        progress.update(50)  # 50% complete
        
        remaining = progress.estimated_remaining
        if remaining is not None:
            # Should estimate roughly equal time remaining
            assert remaining > 0
    
    def test_status_text_generation(self):
        """Test status text formatting."""
        progress = ProgressIndicator(total_steps=100, operation_name="Processing")
        
        # At 0%
        status = progress.get_status_text()
        assert "Processing" in status
        assert "0%" in status
        
        # At 50%
        progress.update(50)
        status = progress.get_status_text()
        assert "Processing" in status
        assert "50%" in status
    
    def test_minimum_total_steps(self):
        """Test that total_steps is at least 1."""
        progress = ProgressIndicator(total_steps=0)
        assert progress.total_steps == 1
        
        progress = ProgressIndicator(total_steps=-10)
        assert progress.total_steps == 1
    
    def test_thread_safety(self):
        """Test that updates are thread-safe."""
        progress = ProgressIndicator(total_steps=1000)
        
        import threading
        
        def increment_many():
            for _ in range(100):
                progress.increment(1)
        
        threads = [threading.Thread(target=increment_many) for _ in range(5)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        
        # Should have incremented 500 times total (5 threads * 100 each)
        assert progress.current_step == 500
        assert progress.percentage == 50


class TestLoadingState:
    """Test cases for LoadingState factory methods."""
    
    def test_create_search_loading(self):
        """Test search loading state creation."""
        loading = LoadingState.create_search_loading(query="TEST123")
        assert loading is not None
        # Should be a container
        assert loading.content is not None
    
    def test_create_search_loading_without_query(self):
        """Test search loading with empty query."""
        loading = LoadingState.create_search_loading(query="")
        assert loading is not None
    
    def test_create_report_loading(self):
        """Test report loading state creation."""
        loading = LoadingState.create_report_loading(
            filename="test_report.xlsx",
            num_tests=5
        )
        assert loading is not None
        assert loading.content is not None
    
    def test_create_report_loading_minimal(self):
        """Test report loading with minimal parameters."""
        loading = LoadingState.create_report_loading()
        assert loading is not None
    
    def test_create_generic_loading(self):
        """Test generic loading state creation."""
        loading = LoadingState.create_generic_loading(message="Processing...")
        assert loading is not None
    
    def test_create_generic_loading_with_details(self):
        """Test generic loading with detail items."""
        details = ["Step 1", "Step 2", "Step 3"]
        loading = LoadingState.create_generic_loading(
            message="Processing...",
            show_details=True,
            details=details
        )
        assert loading is not None
        # Should have more controls when details are shown
        assert loading.content is not None
