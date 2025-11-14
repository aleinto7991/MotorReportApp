from collections import Counter
from pathlib import Path
import logging
from typing import List

from PIL import Image

logger = logging.getLogger(__name__)

def extract_dominant_colors(
    image_path_str: str, 
    num_colors: int = 2, 
    default_colors: List[str] = ["#4F81BD", "#C0504D"]
) -> List[str]:
    """Extracts dominant colors from an image file."""
    if not image_path_str or not Path(image_path_str).exists():
        logger.warning(f"Logo file not found at '{image_path_str}'. Using default tab colors.")
        return default_colors[:num_colors]
    try:
        img = Image.open(image_path_str).convert('RGB')
        img.thumbnail((100, 100))
        
        pixels = list(img.getdata())
        if not pixels:
            logger.warning(f"Could not get pixel data from logo {image_path_str}. Using default tab colors.")
            return default_colors[:num_colors]

        color_counts = Counter(pixels)
        common_colors_rgb = [item[0] for item in color_counts.most_common(num_colors + 5)]

        filtered_colors_rgb = []
        for r, g, b in common_colors_rgb:
            if not ((r > 240 and g > 240 and b > 240) or (r < 15 and g < 15 and b < 15)):
                filtered_colors_rgb.append((r, g, b))
        
        if not filtered_colors_rgb:
            filtered_colors_rgb = [item[0] for item in color_counts.most_common(num_colors)]

        hex_colors = [f"#{r:02x}{g:02x}{b:02x}".upper() for r, g, b in filtered_colors_rgb[:num_colors]]
        
        final_colors = hex_colors
        if len(final_colors) < num_colors:
            logger.info(f"Extracted fewer than {num_colors} distinct non-white/black colors. Padding with defaults.")
            final_colors.extend(default_colors[:num_colors - len(final_colors)])
        
        logger.info(f"Extracted logo colors: {final_colors[:num_colors]}")
        return final_colors[:num_colors]

    except ImportError:
        logger.warning("Pillow (PIL) library is not installed. Cannot extract logo colors. Please install it (`pip install Pillow`). Using default tab colors.")
        return default_colors[:num_colors]
    except Exception as e:
        logger.error(f"Error extracting colors from logo {image_path_str}: {e}. Using default tab colors.")
        return default_colors[:num_colors]
