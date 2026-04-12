"""WCAG 2.1 color contrast utilities for accessible presentation generation.

Provides functions to compute relative luminance, contrast ratios, and
dynamically choose text colors that guarantee readability against any
background — following WCAG AA standards (≥4.5:1 normal text, ≥3:1 large).
"""
from __future__ import annotations

from pptx.dml.color import RGBColor


# ── WCAG 2.1 relative luminance ─────────────────────────────────────

def _linearize(channel: int) -> float:
    """Convert an 8-bit sRGB channel to linear-light per WCAG 2.1."""
    s = channel / 255.0
    return s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4


def relative_luminance(hex_color: str) -> float:
    """Return WCAG 2.1 relative luminance (0.0–1.0) for a hex color string.

    Accepts 6-char hex with or without '#' prefix.
    """
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return 0.0
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return 0.2126 * _linearize(r) + 0.7152 * _linearize(g) + 0.0722 * _linearize(b)


def relative_luminance_rgb(color: RGBColor) -> float:
    """Relative luminance from an ``RGBColor`` object."""
    return 0.2126 * _linearize(color[0]) + 0.7152 * _linearize(color[1]) + 0.0722 * _linearize(color[2])


# ── Contrast ratio ──────────────────────────────────────────────────

def contrast_ratio(hex_fg: str, hex_bg: str) -> float:
    """Return WCAG contrast ratio between two hex colors (always ≥1.0)."""
    l1 = relative_luminance(hex_fg)
    l2 = relative_luminance(hex_bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def contrast_ratio_rgb(fg: RGBColor, bg: RGBColor) -> float:
    """Contrast ratio from two ``RGBColor`` objects."""
    l1 = relative_luminance_rgb(fg)
    l2 = relative_luminance_rgb(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


# ── Dynamic text-color selection ────────────────────────────────────

_WHITE = "FFFFFF"
_DARK = "1A1A2E"  # near-black with slight warmth — softer than pure #000


def pick_text_color(bg_hex: str, *, large_text: bool = False) -> str:
    """Return the best text color hex for the given background.

    Uses WCAG AA thresholds:
    - Normal text: ≥4.5:1
    - Large text (≥18pt or ≥14pt bold): ≥3.0:1

    Returns ``"FFFFFF"`` (white) if it passes; otherwise ``"1A1A2E"`` (dark).
    """
    threshold = 3.0 if large_text else 4.5
    if contrast_ratio(_WHITE, bg_hex) >= threshold:
        return _WHITE
    return _DARK


def pick_text_color_rgb(bg: RGBColor, *, large_text: bool = False) -> RGBColor:
    """Same as :func:`pick_text_color` but accepts/returns ``RGBColor``."""
    hex_result = pick_text_color(f"{bg[0]:02X}{bg[1]:02X}{bg[2]:02X}", large_text=large_text)
    return _hex_to_rgb(hex_result)


# ── Color manipulation helpers ──────────────────────────────────────

def darken_hex(hex_color: str, amount: float = 0.15) -> str:
    """Darken a hex color by *amount* (0.0–1.0). Returns 6-char hex."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    factor = max(0.0, 1.0 - amount)
    r2 = max(0, min(255, int(r * factor)))
    g2 = max(0, min(255, int(g * factor)))
    b2 = max(0, min(255, int(b * factor)))
    return f"{r2:02X}{g2:02X}{b2:02X}"


def lighten_hex(hex_color: str, amount: float = 0.15) -> str:
    """Lighten a hex color by *amount* (0.0–1.0). Returns 6-char hex."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r2 = max(0, min(255, int(r + (255 - r) * amount)))
    g2 = max(0, min(255, int(g + (255 - g) * amount)))
    b2 = max(0, min(255, int(b + (255 - b) * amount)))
    return f"{r2:02X}{g2:02X}{b2:02X}"


def _hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert a 6-char hex string to ``RGBColor``."""
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── WCAG-compliant fallback accent palettes ─────────────────────────
# These replace the original bright fallbacks that failed contrast with
# white text.  Every color below achieves ≥4.5:1 against #FFFFFF.

FALLBACK_ACCENT_HEX: list[str] = [
    "2B5797",  # deep blue   (contrast vs white ≈ 7.0:1)
    "C44A1C",  # burnt orange (≈ 4.6:1)
    "5A5A5A",  # medium gray  (≈ 5.9:1)
    "8B6914",  # dark gold    (≈ 4.8:1)
    "2E75B6",  # mid blue     (≈ 4.6:1)
    "3D7A2E",  # forest green (≈ 4.9:1)
]

# Smaller rotation for infographics (Guidelines §4B: avoid random colors)
FLOW_ACCENT_HEX: list[str] = FALLBACK_ACCENT_HEX[:3]
CMP_ACCENT_HEX: list[str] = [FALLBACK_ACCENT_HEX[0], FALLBACK_ACCENT_HEX[5]]
KPI_ACCENT_HEX: list[str] = [FALLBACK_ACCENT_HEX[0], FALLBACK_ACCENT_HEX[5], FALLBACK_ACCENT_HEX[1], FALLBACK_ACCENT_HEX[3]]
