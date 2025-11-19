"""
Loading state components for better user feedback.

Provides skeleton loaders, progress indicators, and animated placeholders
to improve perceived performance during long operations.
"""
import flet as ft
import logging
from typing import Optional, List, Callable
import time
import threading

logger = logging.getLogger(__name__)

ColorResolver = Optional[Callable[[str, str], Optional[str]]]


def _color(resolver: ColorResolver, token: str, fallback: str) -> str:
    """Resolve a semantic color token via the active theme when possible."""

    if callable(resolver):
        try:
            resolved = resolver(token, fallback)
            if resolved:
                return resolved
        except Exception as exc:
            logger.debug("loading_states color fallback (%s): %s", token, exc)
    return fallback


class SkeletonLoader:
    """
    Creates skeleton loading placeholders that mimic the structure of content.
    
    Skeleton loaders are animated placeholders that show where content will appear,
    improving perceived performance by giving immediate visual feedback.
    """
    
    @staticmethod
    def create_shimmer_effect() -> dict:
        """
        Create a shimmer animation effect for skeleton elements.
        
        Returns:
            Animation configuration dict for Flet
        """
        return {
            "bgcolor": "#e0e0e0",
            "animate_opacity": ft.Animation(
                duration=1500,
                curve=ft.AnimationCurve.EASE_IN_OUT
            )
        }
    
    @staticmethod
    def skeleton_line(
        width: int = 200,
        height: int = 12,
        opacity: float = 0.6,
        *,
        color_resolver: ColorResolver = None,
    ) -> ft.Container:
        """
        Create a skeleton line placeholder.
        
        Args:
            width: Width in pixels
            height: Height in pixels
            opacity: Opacity (0.0 to 1.0)
        
        Returns:
            Animated skeleton container
        """
        return ft.Container(
            width=width,
            height=height,
            bgcolor=_color(color_resolver, 'surface_variant', '#e0e0e0'),
            border_radius=4,
            opacity=opacity,
            animate_opacity=1500
        )
    
    @staticmethod
    def skeleton_row(*, color_resolver: ColorResolver = None) -> ft.Container:
        """
        Create a skeleton placeholder for a table row.
        
        Returns:
            Container with skeleton row structure
        """
        return ft.Container(
            content=ft.Row([
                ft.Container(
                    width=30,
                    height=30,
                    bgcolor=_color(color_resolver, 'surface_variant', '#e0e0e0'),
                    border_radius=6
                ),  # Checkbox
                SkeletonLoader.skeleton_line(width=110, height=12, color_resolver=color_resolver),  # Test lab
                SkeletonLoader.skeleton_line(width=100, height=12, color_resolver=color_resolver),  # Date
                SkeletonLoader.skeleton_line(width=70, height=12, color_resolver=color_resolver),   # Voltage
                SkeletonLoader.skeleton_line(width=250, height=12, color_resolver=color_resolver),  # Notes
            ], spacing=15),
            padding=ft.padding.symmetric(horizontal=12, vertical=8),
            bgcolor=_color(color_resolver, 'surface_variant', '#f8f8f8'),
            border_radius=6,
            margin=ft.margin.only(bottom=4),
            opacity=0.7,
            animate_opacity=ft.Animation(
                duration=1000,
                curve=ft.AnimationCurve.EASE_IN_OUT
            )
        )
    
    @staticmethod
    def search_results_skeleton(num_rows: int = 5, *, color_resolver: ColorResolver = None) -> ft.Column:
        """
        Create a skeleton placeholder for search results.
        
        Args:
            num_rows: Number of skeleton rows to show
        
        Returns:
            Column with skeleton structure
        """
        rows = [SkeletonLoader.skeleton_row(color_resolver=color_resolver) for _ in range(num_rows)]
        
        return ft.Column(
            controls=[
                # Header skeleton
                ft.Container(
                    content=ft.Row([
                        SkeletonLoader.skeleton_line(width=300, height=16, color_resolver=color_resolver),
                    ]),
                    padding=10,
                    bgcolor=_color(color_resolver, 'surface', '#fafafa'),
                    border_radius=8,
                    margin=ft.margin.only(bottom=10)
                ),
                # Column headers skeleton
                ft.Container(
                    content=ft.Row([
                        SkeletonLoader.skeleton_line(width=110, height=14, color_resolver=color_resolver),
                        SkeletonLoader.skeleton_line(width=100, height=14, color_resolver=color_resolver),
                        SkeletonLoader.skeleton_line(width=70, height=14, color_resolver=color_resolver),
                        SkeletonLoader.skeleton_line(width=200, height=14, color_resolver=color_resolver),
                    ], spacing=15),
                    padding=ft.padding.symmetric(horizontal=12, vertical=10),
                    bgcolor=_color(color_resolver, 'surface_variant', '#f0f0f0'),
                    border_radius=6,
                    margin=ft.margin.only(bottom=8)
                ),
                # Rows
                *rows
            ],
            spacing=0
        )


class ProgressIndicator:
    """
    Enhanced progress indicators with percentage, time estimation, and step tracking.
    """
    
    def __init__(self, total_steps: int = 100, operation_name: str = "Operation"):
        """
        Initialize progress indicator.
        
        Args:
            total_steps: Total number of steps (for percentage calculation)
            operation_name: Name of the operation being tracked
        """
        self.total_steps = max(1, total_steps)
        self.current_step = 0
        self.operation_name = operation_name
        self.start_time = time.time()
        self._lock = threading.Lock()
    
    def update(self, step: int, status_message: str = "") -> None:
        """
        Update progress to a specific step.
        
        Args:
            step: Current step number
            status_message: Optional status message
        """
        with self._lock:
            self.current_step = min(step, self.total_steps)
    
    def increment(self, amount: int = 1) -> None:
        """
        Increment progress by a number of steps.
        
        Args:
            amount: Number of steps to increment
        """
        with self._lock:
            self.current_step = min(self.current_step + amount, self.total_steps)
    
    @property
    def percentage(self) -> int:
        """Get current progress percentage (0-100)."""
        with self._lock:
            return int((self.current_step / self.total_steps) * 100)
    
    @property
    def elapsed_time(self) -> float:
        """Get elapsed time in seconds."""
        return time.time() - self.start_time
    
    @property
    def estimated_remaining(self) -> Optional[float]:
        """
        Estimate remaining time in seconds.
        
        Returns:
            Estimated seconds remaining, or None if not enough data
        """
        with self._lock:
            if self.current_step == 0:
                return None
            
            elapsed = self.elapsed_time
            progress_fraction = self.current_step / self.total_steps
            
            if progress_fraction == 0:
                return None
            
            total_estimated = elapsed / progress_fraction
            return total_estimated - elapsed
    
    def get_status_text(self) -> str:
        """
        Get formatted status text with percentage and time.
        
        Returns:
            String like "Processing: 45% (5s remaining)"
        """
        percentage = self.percentage
        status = f"{self.operation_name}: {percentage}%"
        
        remaining = self.estimated_remaining
        if remaining is not None and remaining > 0:
            if remaining < 60:
                status += f" ({int(remaining)}s remaining)"
            else:
                minutes = int(remaining / 60)
                status += f" (~{minutes}m remaining)"
        
        return status
    
    def create_progress_widget(self, *, color_resolver: ColorResolver = None) -> ft.Column:
        """
        Create a Flet widget for displaying progress.
        
        Returns:
            Column with progress bar and status text
        """
        progress_bar = ft.ProgressBar(
            value=self.percentage / 100,
            width=300,
            height=8,
            bgcolor=_color(color_resolver, 'surface_variant', '#e0e0e0'),
            color=_color(color_resolver, 'primary', 'blue')
        )
        
        status_text = ft.Text(
            self.get_status_text(),
            size=13,
            weight=ft.FontWeight.W_500,
            text_align=ft.TextAlign.CENTER,
            color=_color(color_resolver, 'text_muted', '#475467')
        )
        
        return ft.Column(
            controls=[
                status_text,
                progress_bar
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=8
        )


class LoadingState:
    """
    Centralized loading state components for common operations.
    """
    
    @staticmethod
    def create_search_loading(
        query: str = "",
        *,
        color_resolver: ColorResolver = None,
    ) -> ft.Container:
        """
        Create a search loading state with animation.
        
        Args:
            query: Search query being processed
        
        Returns:
            Animated loading container
        """
        message = f"Searching for '{query}'..." if query else "Searching..."
        
        color = lambda token, fallback: _color(color_resolver, token, fallback)

        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.ProgressRing(
                        width=24,
                        height=24,
                        stroke_width=3,
                        color=color('info', '#0288d1')
                    ),
                    ft.Text(
                        message,
                        size=15,
                        weight=ft.FontWeight.W_500,
                        color=color('on_surface', '#1f2933')
                    ),
                ], spacing=12, alignment=ft.MainAxisAlignment.CENTER),
                ft.Container(height=20),  # Spacer
                SkeletonLoader.search_results_skeleton(num_rows=3, color_resolver=color_resolver)
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=15),
            padding=ft.padding.all(30),
            alignment=ft.alignment.center,
            bgcolor=color('surface', '#ffffff'),
            border_radius=10
        )
    
    @staticmethod
    def create_report_loading(
        filename: str = "",
        num_tests: int = 0,
        *,
        color_resolver: ColorResolver = None,
    ) -> ft.Container:
        """
        Create a report generation loading state.
        
        Args:
            filename: Report filename being generated
            num_tests: Number of tests in the report
        
        Returns:
            Loading container with details
        """
        message = f"Generating report: {filename}" if filename else "Generating report..."
        
        color = lambda token, fallback: _color(color_resolver, token, fallback)

        details = []
        if num_tests > 0:
            details.append(ft.Text(f"📊 Processing {num_tests} test(s)", size=13, color=color('text_muted', 'grey')))
        details.extend([
            ft.Text("⚙️ Building Excel workbook", size=13, color=color('text_muted', 'grey')),
            ft.Text("📈 Creating charts and tables", size=13, color=color('text_muted', 'grey')),
            ft.Text("💾 Writing to file", size=13, color=color('text_muted', 'grey')),
        ])
        
        return ft.Container(
            content=ft.Column([
                ft.Row([
                    ft.ProgressRing(
                        width=28,
                        height=28,
                        stroke_width=4,
                        color=color('success', 'green')
                    ),
                    ft.Text(
                        message,
                        size=16,
                        weight=ft.FontWeight.W_600,
                        color=color('on_surface', '#111827')
                    ),
                ], spacing=15, alignment=ft.MainAxisAlignment.CENTER),
                ft.Container(height=15),
                ft.Container(
                    content=ft.Column(details, spacing=8),
                    padding=15,
                    bgcolor=color('surface_variant', '#f8f8f8'),
                    border_radius=8,
                    border=ft.border.all(1, color('outline', '#e0e0e0'))
                ),
                ft.Container(height=10),
                ft.Text(
                    "This may take a moment...",
                    size=12,
                    color=color('text_muted', 'grey'),
                    italic=True
                )
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=5),
            padding=ft.padding.all(35),
            alignment=ft.alignment.center,
            bgcolor=color('surface', '#ffffff'),
            border_radius=12
        )
    
    @staticmethod
    def create_generic_loading(
        message: str = "Loading...",
        show_details: bool = False,
        details: Optional[List[str]] = None,
        *,
        color_resolver: ColorResolver = None,
    ) -> ft.Container:
        """
        Create a generic loading state.
        
        Args:
            message: Loading message
            show_details: Whether to show detail items
            details: List of detail strings to display
        
        Returns:
            Loading container
        """
        color = lambda token, fallback: _color(color_resolver, token, fallback)

        controls = [
            ft.Row([
                ft.ProgressRing(
                    width=20,
                    height=20,
                    stroke_width=3,
                    color=color('primary', '#2563eb')
                ),
                ft.Text(
                    message,
                    size=14,
                    weight=ft.FontWeight.W_500,
                    color=color('on_surface', '#1f2933')
                ),
            ], spacing=10, alignment=ft.MainAxisAlignment.CENTER)
        ]
        
        if show_details and details:
            controls.append(ft.Container(height=10))
            for detail in details:
                controls.append(
                    ft.Text(f"• {detail}", size=12, color=color('text_muted', 'grey'))
                )
        
        return ft.Container(
            content=ft.Column(
                controls,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=5
            ),
            padding=ft.padding.all(25),
            alignment=ft.alignment.center,
            bgcolor=color('surface', '#ffffff'),
            border_radius=8
        )

