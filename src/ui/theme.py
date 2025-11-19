"""Theme helpers and semantic tokens for the Motor Report App UI."""

from __future__ import annotations

from functools import lru_cache
from typing import Dict, Literal, Optional
import logging

try:
    import flet as ft
except Exception:  # pragma: no cover - Flet may not be installed during tests
    ft = None

logger = logging.getLogger(__name__)

ThemeModeLiteral = Literal["light", "dark", "system"]
SemanticPalette = Dict[str, str]


LIGHT_PALETTE: SemanticPalette = {
    "primary": "#1565c0",
    "on_primary": "#ffffff",
    "primary_container": "#e3f2fd",
    "on_primary_container": "#082848",
    "secondary": "#00897b",
    "on_secondary": "#ffffff",
    "secondary_container": "#e0f2f1",
    "on_secondary_container": "#004d43",
    "surface": "#f7f9fc",
    "on_surface": "#1f2933",
    "surface_variant": "#e0e7ef",
    "on_surface_variant": "#3d4a57",
    "background": "#f5f5f5",
    "on_background": "#1a1a1a",
    "outline": "#c5cbd3",
    "success": "#2e7d32",
    "on_success": "#ffffff",
    "success_container": "#c8e6c9",
    "warning": "#f9a825",
    "on_warning": "#1a1a1a",
    "warning_container": "#fff8e1",
    "error": "#c62828",
    "on_error": "#ffffff",
    "error_container": "#ffdde0",
    "info": "#0288d1",
    "on_info": "#ffffff",
    "info_container": "#e1f5fe",
    "text_muted": "#5f6b7a",
}

DARK_PALETTE: SemanticPalette = {
    "primary": "#90caf9",
    "on_primary": "#082848",
    "primary_container": "#1553c1",
    "on_primary_container": "#f7fbff",
    "secondary": "#4db6ac",
    "on_secondary": "#062320",
    "secondary_container": "#004d43",
    "on_secondary_container": "#a7ffeb",
    "surface": "#171d25",
    "on_surface": "#fefefe",
    "surface_variant": "#222a35",
    "on_surface_variant": "#fefefe",
    "background": "#141920",
    "on_background": "#fefefe",
    "outline": "#8895a8",
    "success": "#6fdc8c",
    "on_success": "#06220f",
    "success_container": "#1b5e20",
    "warning": "#ffd54f",
    "on_warning": "#241a00",
    "warning_container": "#7b5e00",
    "error": "#ef9a9a",
    "on_error": "#2b0000",
    "error_container": "#8a1c1c",
    "info": "#81d4fa",
    "on_info": "#042333",
    "info_container": "#01579b",
    "text_muted": "#fefefe",
}


class ThemeManager:
    """Provides a central source of truth for theming and semantic tokens."""

    def __init__(self) -> None:
        self._palettes: Dict[str, SemanticPalette] = {
            "light": LIGHT_PALETTE,
            "dark": DARK_PALETTE,
        }
        self._theme_cache: Dict[str, Optional["ft.Theme"]] = {"light": None, "dark": None}
        self._last_explicit_mode: ThemeModeLiteral = "system"
        self._last_palette_key: str = "light"

    def apply(self, page, default_mode: ThemeModeLiteral = "system") -> None:
        if ft is None:
            logger.debug("Flet not available; ThemeManager.apply is a no-op.")
            return

        try:
            mode = self._resolve_requested_mode(page, default_mode)
            self._last_explicit_mode = mode
            palette_key = self._select_palette_key(page, mode)
            self._last_palette_key = palette_key

            # Attach concrete Theme objects so Flet widgets inherit defaults
            page.theme = self._theme_cache[palette_key] or self._cache_theme(palette_key)
            opposite_key = "dark" if palette_key == "light" else "light"
            page.dark_theme = self._theme_cache[opposite_key] or self._cache_theme(opposite_key)

            # Apply the requested ThemeMode for Flet internals
            self._set_theme_mode(page, mode)

            # Align default surfaces/backgrounds for consistent look
            try:
                page.bgcolor = self._palettes[palette_key]["surface"]
            except Exception:
                pass

            try:
                page.update()
            except Exception:
                pass
        except Exception as exc:  # pragma: no cover - defensive best-effort
            logger.debug("ThemeManager.apply encountered an error: %s", exc)

    def resolve_token(
        self,
        token: str,
        *,
        page=None,
        fallback: Optional[str] = None,
        mode_override: Optional[ThemeModeLiteral] = None,
    ) -> Optional[str]:
        """Resolve a semantic color token for the current (or requested) mode."""

        palette_key = self._select_palette_key(page, mode_override or self._last_explicit_mode)
        palette = self._palettes[palette_key]
        if token in palette:
            return palette[token]

        # Fall back to Flet color scheme attributes if available
        scheme_attr = self._alias_color_scheme_attribute(token)
        if page and scheme_attr:
            try:
                scheme = getattr(getattr(page, "theme", None), "color_scheme", None)
                if scheme and hasattr(scheme, scheme_attr):
                    val = getattr(scheme, scheme_attr)
                    if val:
                        return val
            except Exception:
                pass

        return fallback

    def _resolve_requested_mode(self, page, default_mode: ThemeModeLiteral) -> ThemeModeLiteral:
        mode = default_mode
        if hasattr(page, "client_storage") and page.client_storage is not None:
            try:
                stored = page.client_storage.get("theme")
                if stored in ("light", "dark", "system"):
                    mode = stored  # persist user preference
            except Exception:
                pass
        return mode

    def _detect_platform_brightness(self, page) -> Optional[str]:
        if ft is None:
            return None
        try:
            brightness = getattr(page, "platform_brightness", None)
            if brightness == ft.Brightness.DARK:
                return "dark"
            if brightness == ft.Brightness.LIGHT:
                return "light"
        except Exception:
            pass
        return None

    def _select_palette_key(self, page, mode: ThemeModeLiteral) -> str:
        if mode == "dark":
            return "dark"
        if mode == "light":
            return "light"

        detected = self._detect_platform_brightness(page)
        if detected in ("light", "dark"):
            return detected
        return self._last_palette_key

    def _cache_theme(self, palette_key: str):
        palette = self._palettes[palette_key]
        theme = self._build_theme(palette)
        self._theme_cache[palette_key] = theme
        return theme

    def _build_theme(self, palette: SemanticPalette):
        if ft is None:
            return None
        try:
            theme = ft.Theme(
                color_scheme=ft.ColorScheme(
                    primary=palette["primary"],
                    on_primary=palette["on_primary"],
                    primary_container=palette["primary_container"],
                    on_primary_container=palette["on_primary_container"],
                    secondary=palette["secondary"],
                    on_secondary=palette["on_secondary"],
                    secondary_container=palette["secondary_container"],
                    on_secondary_container=palette["primary"],
                    surface=palette["surface"],
                    on_surface=palette["on_surface"],
                    surface_variant=palette["surface_variant"],
                    on_surface_variant=palette["on_surface_variant"],
                    background=palette["background"],
                    on_background=palette["on_background"],
                    error=palette["error"],
                    on_error=palette["on_error"],
                    outline=palette["outline"],
                )
            )
            return theme
        except Exception as exc:
            logger.debug("Unable to build Flet Theme: %s", exc)
            return None

    def _set_theme_mode(self, page, mode: ThemeModeLiteral) -> None:
        if ft is None:
            return
        try:
            if mode == "dark":
                page.theme_mode = ft.ThemeMode.DARK
            elif mode == "light":
                page.theme_mode = ft.ThemeMode.LIGHT
            else:
                page.theme_mode = ft.ThemeMode.SYSTEM
        except Exception:
            pass

    @staticmethod
    @lru_cache(maxsize=None)
    def _alias_color_scheme_attribute(token: str) -> Optional[str]:
        aliases = {
            "primary": "primary",
            "on_primary": "on_primary",
            "primary_container": "primary_container",
            "on_primary_container": "on_primary_container",
            "surface": "surface",
            "on_surface": "on_surface",
            "surface_variant": "surface_variant",
            "on_surface_variant": "on_surface_variant",
            "background": "background",
            "on_background": "on_background",
            "error": "error",
            "on_error": "on_error",
            "outline": "outline",
            "secondary": "secondary",
            "on_secondary": "on_secondary",
            "secondary_container": "secondary_container",
            "on_secondary_container": "on_secondary_container",
        }
        return aliases.get(token)


theme_manager = ThemeManager()


def apply_theme(page, default_mode: ThemeModeLiteral = "system") -> None:
    theme_manager.apply(page, default_mode)


def set_user_theme(page, mode: ThemeModeLiteral) -> None:
    if ft is None:
        return
    try:
        if hasattr(page, "client_storage") and page.client_storage is not None:
            try:
                page.client_storage.set("theme", mode)
            except Exception:
                pass
        theme_manager.apply(page, default_mode=mode)
    except Exception:
        pass


def toggle_theme(page) -> None:
    if ft is None:
        return
    try:
        current = getattr(page, "theme_mode", None)
        if current == ft.ThemeMode.DARK:
            set_user_theme(page, "light")
        else:
            set_user_theme(page, "dark")
    except Exception:
        pass


def resolve_token(page, token: str, fallback: Optional[str] = None) -> Optional[str]:
    """Helper exposed to UI components so they can request semantic colors."""

    try:
        return theme_manager.resolve_token(token, page=page, fallback=fallback)
    except Exception:
        return fallback
