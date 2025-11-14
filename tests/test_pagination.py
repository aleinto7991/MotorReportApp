"""
Unit tests for the pagination utility.
"""
import pytest
from src.ui.utils.pagination import Paginator


class TestPaginator:
    """Test cases for the Paginator class."""
    
    def test_basic_initialization(self):
        """Test paginator initializes with correct defaults."""
        paginator = Paginator()
        assert paginator.total_items == 0
        assert paginator.total_pages == 0
        assert paginator.current_page == 0
        assert paginator.page_size == 50
    
    def test_custom_page_size(self):
        """Test custom page size."""
        items = list(range(100))
        paginator = Paginator(items=items, page_size=25)
        assert paginator.page_size == 25
        assert paginator.total_pages == 4
    
    def test_pagination_calculation(self):
        """Test correct page calculations."""
        items = list(range(125))  # 125 items
        paginator = Paginator(items=items, page_size=50)
        
        assert paginator.total_items == 125
        assert paginator.total_pages == 3  # 50 + 50 + 25
        assert len(paginator.get_current_page()) == 50
    
    def test_page_navigation(self):
        """Test navigating between pages."""
        items = list(range(100))
        paginator = Paginator(items=items, page_size=30)
        
        # Initial page
        assert paginator.current_page == 0
        assert not paginator.has_previous
        assert paginator.has_next
        
        # Next page
        assert paginator.next_page() is True
        assert paginator.current_page == 1
        assert paginator.has_previous
        assert paginator.has_next
        
        # Last page
        assert paginator.last_page() is True
        assert paginator.current_page == 3
        assert not paginator.has_next
        
        # Previous page
        assert paginator.previous_page() is True
        assert paginator.current_page == 2
        
        # First page
        assert paginator.first_page() is True
        assert paginator.current_page == 0
    
    def test_get_current_page_content(self):
        """Test getting correct page content."""
        items = list(range(10, 60))  # 50 items: 10, 11, ..., 59
        paginator = Paginator(items=items, page_size=20)
        
        # First page: items 10-29
        page1 = paginator.get_current_page()
        assert len(page1) == 20
        assert page1[0] == 10
        assert page1[-1] == 29
        
        # Second page: items 30-49
        paginator.next_page()
        page2 = paginator.get_current_page()
        assert len(page2) == 20
        assert page2[0] == 30
        assert page2[-1] == 49
        
        # Third page: items 50-59 (partial page)
        paginator.next_page()
        page3 = paginator.get_current_page()
        assert len(page3) == 10
        assert page3[0] == 50
        assert page3[-1] == 59
    
    def test_page_info_display(self):
        """Test page info string generation."""
        items = list(range(125))
        paginator = Paginator(items=items, page_size=50)
        
        assert paginator.get_page_info() == "Showing 1-50 of 125"
        
        paginator.next_page()
        assert paginator.get_page_info() == "Showing 51-100 of 125"
        
        paginator.next_page()
        assert paginator.get_page_info() == "Showing 101-125 of 125"
    
    def test_empty_items(self):
        """Test paginator with no items."""
        paginator = Paginator(items=[])
        
        assert paginator.total_items == 0
        assert paginator.total_pages == 0
        assert paginator.get_current_page() == []
        assert paginator.get_page_info() == "No items"
        assert not paginator.has_next
        assert not paginator.has_previous
    
    def test_items_update_resets_page(self):
        """Test that updating items resets to first page."""
        paginator = Paginator(items=list(range(100)), page_size=25)
        
        # Navigate to page 2
        paginator.go_to_page(2)
        assert paginator.current_page == 2
        
        # Update items - should reset to page 0
        paginator.items = list(range(50))
        assert paginator.current_page == 0
    
    def test_go_to_page_validation(self):
        """Test go_to_page validates page numbers."""
        items = list(range(100))
        paginator = Paginator(items=items, page_size=25)
        
        # Valid page
        assert paginator.go_to_page(2) is True
        assert paginator.current_page == 2
        
        # Invalid pages
        assert paginator.go_to_page(-1) is False
        assert paginator.go_to_page(10) is False
        assert paginator.current_page == 2  # Unchanged
    
    def test_page_size_minimum(self):
        """Test page size enforces minimum of 1."""
        paginator = Paginator(items=list(range(10)), page_size=0)
        assert paginator.page_size == 1
        
        paginator.page_size = -5
        assert paginator.page_size == 1
    
    def test_navigation_controls_creation(self):
        """Test that navigation controls are created without errors."""
        items = list(range(100))
        paginator = Paginator(items=items, page_size=25)
        
        # Should create controls without error
        controls = paginator.create_navigation_controls()
        assert controls is not None
        
        # Test with callback
        callback_called = []
        
        def test_callback():
            callback_called.append(True)
        
        controls_with_callback = paginator.create_navigation_controls(test_callback)
        assert controls_with_callback is not None
