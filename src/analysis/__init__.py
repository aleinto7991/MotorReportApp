"""Analysis modules for Motor Report Application.

Exports only symbols that exist in the refactored analysis package.
"""

from .noise_handler import NoiseDataHandler
from .noise_chart_generator import NoiseChartGenerator
from .image_utils import extract_dominant_colors

__all__ = [
    "NoiseDataHandler",
    "NoiseChartGenerator",
    "extract_dominant_colors",
]
