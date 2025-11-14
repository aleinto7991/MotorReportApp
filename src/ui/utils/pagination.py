"""
Pagination utility for large result sets.
Improves performance by rendering only visible items instead of entire datasets.
"""
import flet as ft
import logging
from typing import List, Callable, TypeVar, Generic, Optional

logger = logging.getLogger(__name__)

T = TypeVar('T')


class Paginator(Generic[T]):
    """
    Manages pagination for large lists of items.
    
    Features:
    - Configurable page size (default: 50 items)
    - Automatic page calculation
    - Navigation controls with page info
    - Callback-based rendering for flexibility
    
    Example usage:
        paginator = Paginator(items=all_tests, page_size=50)
        visible_items = paginator.get_current_page()
        controls = paginator.create_navigation_controls(on_page_change_callback)
    """
    
    def __init__(
        self,
        items: Optional[List[T]] = None,
        page_size: int = 50,
        current_page: int = 0
    ):
        """
        Initialize paginator.
        
        Args:
            items: List of items to paginate (can be set later)
            page_size: Number of items per page
            current_page: Starting page (0-indexed)
        """
        self._items: List[T] = items or []
        self._page_size = max(1, page_size)  # Ensure at least 1
        self._current_page = max(0, current_page)  # Ensure non-negative
    
    @property
    def items(self) -> List[T]:
        """Get current items list."""
        return self._items
    
    @items.setter
    def items(self, value: List[T]):
        """
        Set items and reset to first page.
        
        Args:
            value: New list of items
        """
        self._items = value or []
        self._current_page = 0  # Reset to first page when items change
        logger.debug(f"Paginator items updated: {len(self._items)} total items")
    
    @property
    def page_size(self) -> int:
        """Get items per page."""
        return self._page_size
    
    @page_size.setter
    def page_size(self, value: int):
        """
        Set page size and recalculate current page.
        
        Args:
            value: New page size (minimum 1)
        """
        self._page_size = max(1, value)
        # Ensure current page is still valid
        if self._current_page >= self.total_pages:
            self._current_page = max(0, self.total_pages - 1)
    
    @property
    def current_page(self) -> int:
        """Get current page (0-indexed)."""
        return self._current_page
    
    @property
    def total_items(self) -> int:
        """Get total number of items."""
        return len(self._items)
    
    @property
    def total_pages(self) -> int:
        """Get total number of pages."""
        if not self._items:
            return 0
        return (len(self._items) + self._page_size - 1) // self._page_size
    
    @property
    def has_previous(self) -> bool:
        """Check if there's a previous page."""
        return self._current_page > 0
    
    @property
    def has_next(self) -> bool:
        """Check if there's a next page."""
        return self._current_page < self.total_pages - 1
    
    def get_current_page(self) -> List[T]:
        """
        Get items for the current page.
        
        Returns:
            List of items for current page (slice of all items)
        """
        if not self._items:
            return []
        
        start_idx = self._current_page * self._page_size
        end_idx = min(start_idx + self._page_size, len(self._items))
        
        return self._items[start_idx:end_idx]
    
    def go_to_page(self, page: int) -> bool:
        """
        Navigate to specific page.
        
        Args:
            page: Page number (0-indexed)
        
        Returns:
            True if page changed, False if invalid page
        """
        if 0 <= page < self.total_pages:
            self._current_page = page
            logger.debug(f"Navigated to page {page + 1}/{self.total_pages}")
            return True
        return False
    
    def next_page(self) -> bool:
        """
        Navigate to next page.
        
        Returns:
            True if moved to next page, False if already at last page
        """
        if self.has_next:
            self._current_page += 1
            logger.debug(f"Next page: {self._current_page + 1}/{self.total_pages}")
            return True
        return False
    
    def previous_page(self) -> bool:
        """
        Navigate to previous page.
        
        Returns:
            True if moved to previous page, False if already at first page
        """
        if self.has_previous:
            self._current_page -= 1
            logger.debug(f"Previous page: {self._current_page + 1}/{self.total_pages}")
            return True
        return False
    
    def first_page(self) -> bool:
        """Go to first page."""
        if self._current_page != 0:
            self._current_page = 0
            logger.debug("Navigated to first page")
            return True
        return False
    
    def last_page(self) -> bool:
        """Go to last page."""
        last = max(0, self.total_pages - 1)
        if self._current_page != last:
            self._current_page = last
            logger.debug(f"Navigated to last page: {last + 1}")
            return True
        return False
    
    def get_page_info(self) -> str:
        """
        Get human-readable page info.
        
        Returns:
            String like "Showing 51-100 of 250" or "Page 2 of 5"
        """
        if not self._items:
            return "No items"
        
        start = self._current_page * self._page_size + 1
        end = min((self._current_page + 1) * self._page_size, len(self._items))
        total = len(self._items)
        
        return f"Showing {start}-{end} of {total}"
    
    def create_navigation_controls(
        self,
        on_page_change: Optional[Callable[[], None]] = None
    ) -> ft.Row:
        """
        Create Flet navigation controls for pagination.
        
        Args:
            on_page_change: Callback to execute after page changes
        
        Returns:
            ft.Row with navigation buttons and page info
        """
        def make_nav_handler(action: Callable[[], bool]):
            """Wrapper to call page change callback after navigation."""
            def handler(e):
                if action():  # If page actually changed
                    if on_page_change:
                        on_page_change()
            return handler
        
        return ft.Row(
            controls=[
                ft.IconButton(
                    icon=ft.Icons.FIRST_PAGE,
                    tooltip="First page",
                    disabled=not self.has_previous,
                    on_click=make_nav_handler(self.first_page)
                ),
                ft.IconButton(
                    icon=ft.Icons.CHEVRON_LEFT,
                    tooltip="Previous page",
                    disabled=not self.has_previous,
                    on_click=make_nav_handler(self.previous_page)
                ),
                ft.Container(
                    content=ft.Text(
                        self.get_page_info(),
                        size=13,
                        weight=ft.FontWeight.W_500,
                        text_align=ft.TextAlign.CENTER
                    ),
                    padding=ft.padding.symmetric(horizontal=15, vertical=8),
                    bgcolor="#f0f0f0",
                    border_radius=6
                ),
                ft.IconButton(
                    icon=ft.Icons.CHEVRON_RIGHT,
                    tooltip="Next page",
                    disabled=not self.has_next,
                    on_click=make_nav_handler(self.next_page)
                ),
                ft.IconButton(
                    icon=ft.Icons.LAST_PAGE,
                    tooltip="Last page",
                    disabled=not self.has_next,
                    on_click=make_nav_handler(self.last_page)
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=5
        )

