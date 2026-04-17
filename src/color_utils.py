"""WCAG 2.1 color contrast utilities for accessible presentation generation.

Provides functions to compute relative luminance, contrast ratios, and
dynamically choose text colors that guarantee readability against any
background — following WCAG AA standards (≥4.5:1 normal text, ≥3:1 large).
"""
from __future__ import annotations

import re

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


# ── Brightness-aware effective color ────────────────────────────────

def effective_hex_after_brightness(base_hex: str, brightness: float) -> str:
    """Approximate the hex color rendered by PowerPoint after applying a
    ``theme_color.brightness`` adjustment.

    PowerPoint's brightness value maps to a tint (positive) or shade
    (negative) applied via ``lumMod`` / ``lumOff``. This helper gives us a
    close-enough approximation so we can pick text colors that stay readable
    on the *rendered* fill — which may be radically different from the raw
    theme accent.

    *brightness* is in the range [-1.0, 1.0]:
      - 0.0   → no change (return base)
      - +0.85 → very light tint (the default ``CARD_BRIGHTNESS_LEVELS`` produce colors ≈ 90 % toward white)
      - -0.25 → slightly darker shade
    """
    if not -1.0 <= brightness <= 1.0:
        return base_hex
    h = base_hex.lstrip("#")
    if len(h) != 6:
        return base_hex
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    if brightness > 0:
        r = r + int((255 - r) * brightness)
        g = g + int((255 - g) * brightness)
        b = b + int((255 - b) * brightness)
    elif brightness < 0:
        factor = 1.0 + brightness  # brightness is negative
        r = int(r * factor)
        g = int(g * factor)
        b = int(b * factor)
    r = max(0, min(255, r))
    g = max(0, min(255, g))
    b = max(0, min(255, b))
    return f"{r:02X}{g:02X}{b:02X}"


def pick_text_color_for_brightness(base_hex: str, brightness: float,
                                   *, large_text: bool = False) -> str:
    """Like :func:`pick_text_color`, but computes the effective fill after
    brightness first so pale/tinted cards get readable dark text.
    """
    eff = effective_hex_after_brightness(base_hex, brightness)
    return pick_text_color(eff, large_text=large_text)


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


# ── Numeric abbreviation ────────────────────────────────────────────

_NUM_PATTERN = re.compile(
    r"""
    ^\s*
    (?P<prefix>[-+]?[\$€£¥₹]?)          # optional sign + currency symbol
    \s*
    (?P<number>                          # the numeric payload
        \d{1,3}(?:,\d{3})+(?:\.\d+)?     #   1,234,567.89
        | \d+(?:\.\d+)?                  #   1234567.89  or  12
    )
    \s*
    (?P<suffix>.*?)                      # trailing unit text (USD, %, etc.)
    \s*$
    """,
    re.VERBOSE,
)


def _format_short(v: float) -> str:
    """Format ``v`` with one decimal and strip a trailing ``.0``."""
    s = f"{v:.1f}"
    if s.endswith(".0"):
        s = s[:-2]
    return s


def abbreviate_number(text: str, *, min_magnitude: int = 10_000) -> str:
    """Return *text* with the numeric payload shortened to K/M/B/T when large.

    Leaves the input untouched when:
    - the string does not parse as a single number,
    - the absolute magnitude is below ``min_magnitude`` (default 10 000),
    - the trailing suffix already contains a magnitude word (``million``,
      ``billion``, ``trillion``) or an explicit K/M/B/T letter after the
      number — we assume the author has already formatted it.

    Currency prefixes (``$ € £ ¥ ₹``) and trailing units (``%``, ``USD``,
    ``per year``, ``index points``…) are preserved verbatim.
    """
    if text is None:
        return text
    s = str(text).strip()
    if not s:
        return text

    m = _NUM_PATTERN.match(s)
    if not m:
        return text

    prefix = m.group("prefix") or ""
    raw_num = m.group("number")
    suffix = (m.group("suffix") or "").strip()

    # Skip if suffix already carries a magnitude label
    low_suffix = suffix.lower()
    if any(w in low_suffix for w in ("million", "billion", "trillion", "thousand")):
        return text
    # Leading K/M/B/T right after the number (e.g. "5M", "2.3B") — already short
    if low_suffix[:1] in ("k", "m", "b", "t") and (
        len(low_suffix) == 1 or not low_suffix[1:2].isalpha()
    ):
        return text

    try:
        value = float(raw_num.replace(",", ""))
    except ValueError:
        return text

    a = abs(value)
    if a < min_magnitude:
        return text

    if a >= 1e12:
        body, unit = _format_short(value / 1e12), "T"
    elif a >= 1e9:
        body, unit = _format_short(value / 1e9), "B"
    elif a >= 1e6:
        body, unit = _format_short(value / 1e6), "M"
    elif a >= 1e3:
        body, unit = _format_short(value / 1e3), "K"
    else:
        return text

    # Reassemble: "<prefix><body><unit>[ <suffix>]"
    # If the suffix is a currency-unit word like "USD"/"AED"/"INR" keep it,
    # otherwise drop a lone unit suffix to avoid "$1.2B 000" style artifacts.
    out = f"{prefix}{body}{unit}"
    if suffix:
        # Drop redundant trailing ".00" style fragments; keep currency codes
        # or percent signs or other meaningful units.
        if suffix and not suffix.startswith(("USD", "EUR", "GBP", "INR", "AED",
                                              "CNY", "JPY", "CAD", "AUD", "%", "per")):
            return f"{out} {suffix}" if len(suffix) <= 16 else out
        return f"{out} {suffix}"
    return out
