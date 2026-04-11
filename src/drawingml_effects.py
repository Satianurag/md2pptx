"""DrawingML visual effects for python-pptx shapes.

Provides shadow, gradient, corner radius, and transparency helpers
using a mix of native python-pptx 1.0.2 API and direct XML injection
where the native API doesn't expose the feature.

All functions are verified working on python-pptx 1.0.2.
"""

from __future__ import annotations

import logging
from typing import Sequence

import lxml.etree as etree
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.shapes.base import BaseShape
from pptx.util import Emu

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# XML namespace
# ---------------------------------------------------------------------------
_DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NSMAP = {"a": _DML_NS}

# ---------------------------------------------------------------------------
# Shadow presets
# ---------------------------------------------------------------------------
# Values calibrated to match the sample PPTX outputs.
# blurRad / dist in EMU, dir in 60000ths of a degree, alpha in 1000ths of %.
SHADOW_PRESETS: dict[str, dict] = {
    "card": {
        "blurRad": "50800",   # ~4pt blur
        "dist": "38100",      # ~3pt distance
        "dir": "5400000",     # straight down (90°)
        "alpha": "40000",     # 40% opacity
    },
    "subtle": {
        "blurRad": "25400",   # ~2pt blur
        "dist": "12700",      # ~1pt distance
        "dir": "5400000",
        "alpha": "23000",     # 23% opacity
    },
    "medium": {
        "blurRad": "40000",
        "dist": "23000",
        "dir": "5400000",
        "alpha": "35000",
    },
    "strong": {
        "blurRad": "76200",   # ~6pt blur
        "dist": "50800",      # ~4pt distance
        "dir": "5400000",
        "alpha": "50000",     # 50% opacity
    },
}


def add_shadow(
    shape: BaseShape,
    preset: str = "card",
    *,
    color_hex: str = "000000",
    blur_rad: str | None = None,
    dist: str | None = None,
    direction: str | None = None,
    alpha: str | None = None,
) -> None:
    """Add an outer shadow to *shape* via XML injection.

    Parameters
    ----------
    shape : BaseShape
        Any non-GraphicsFrame shape (auto_shape, text_box, picture, freeform).
    preset : str
        One of ``"card"``, ``"subtle"``, ``"medium"``, ``"strong"``.
    color_hex : str
        Shadow colour as 6-char hex (default black ``"000000"``).
    blur_rad, dist, direction, alpha : str | None
        Override individual preset values (EMU strings).
    """
    try:
        sp = shape._element
        # GraphicsFrame (charts, tables) uses a different mechanism — skip.
        spPr = getattr(sp, "spPr", None)
        if spPr is None:
            log.debug("Shape %s has no spPr — skipping shadow", shape.name)
            return

        vals = dict(SHADOW_PRESETS.get(preset, SHADOW_PRESETS["card"]))
        if blur_rad is not None:
            vals["blurRad"] = blur_rad
        if dist is not None:
            vals["dist"] = dist
        if direction is not None:
            vals["dir"] = direction
        if alpha is not None:
            vals["alpha"] = alpha

        # Remove existing effectLst if present to avoid duplicates
        for existing in spPr.findall(f"{{{_DML_NS}}}effectLst"):
            spPr.remove(existing)

        xml = (
            f'<a:effectLst xmlns:a="{_DML_NS}">'
            f'<a:outerShdw blurRad="{vals["blurRad"]}" dist="{vals["dist"]}" '
            f'dir="{vals["dir"]}" rotWithShape="0">'
            f'<a:srgbClr val="{color_hex}">'
            f'<a:alpha val="{vals["alpha"]}"/>'
            f"</a:srgbClr>"
            f"</a:outerShdw>"
            f"</a:effectLst>"
        )
        spPr.append(etree.fromstring(xml))
    except Exception:
        log.warning("Failed to add shadow to shape %s", shape.name, exc_info=True)


# ---------------------------------------------------------------------------
# Gradient fill  (native python-pptx API)
# ---------------------------------------------------------------------------

def add_gradient(
    shape: BaseShape,
    stops: Sequence[tuple[float, RGBColor]],
    angle: float = 270.0,
) -> None:
    """Apply a linear gradient fill using **RGB colours**.

    Parameters
    ----------
    stops : sequence of (position, RGBColor)
        Each stop is ``(0.0–1.0 position, RGBColor)``.
        Must have at least 2 stops.
    angle : float
        Gradient angle in degrees (270 = top-to-bottom, 0 = left-to-right).
    """
    if len(stops) < 2:
        return
    try:
        fill = shape.fill
        fill.gradient()
        fill.gradient_angle = angle

        gs = fill.gradient_stops
        # Set first two stops (always present after .gradient())
        gs[0].position = stops[0][0]
        gs[0].color.rgb = stops[0][1]
        gs[1].position = stops[1][0]
        gs[1].color.rgb = stops[1][1]
        # python-pptx default gradient has 2 stops; additional stops require
        # XML manipulation — not implemented yet (2-stop covers 95% of cases).
    except Exception:
        log.warning("Failed to add gradient to shape %s", shape.name, exc_info=True)


def add_theme_gradient(
    shape: BaseShape,
    theme_color: MSO_THEME_COLOR = MSO_THEME_COLOR.ACCENT_1,
    angle: float = 270.0,
    brightness_delta: float = 0.4,
) -> None:
    """Apply a linear gradient using **theme/scheme colours**.

    Creates a 2-stop gradient from the theme colour to a lighter version,
    ensuring the shape automatically adapts to any Slide Master template.

    Parameters
    ----------
    theme_color : MSO_THEME_COLOR
        The accent colour to use (e.g. ``ACCENT_1``).
    angle : float
        Gradient angle in degrees.
    brightness_delta : float
        How much lighter the second stop should be (0.0–1.0).
    """
    try:
        fill = shape.fill
        fill.gradient()
        fill.gradient_angle = angle

        gs = fill.gradient_stops
        gs[0].position = 0.0
        gs[0].color.theme_color = theme_color
        gs[1].position = 1.0
        gs[1].color.theme_color = theme_color
        gs[1].color.brightness = brightness_delta
    except Exception:
        log.warning("Failed to add theme gradient to shape %s", shape.name, exc_info=True)


# ---------------------------------------------------------------------------
# Corner radius  (XML injection)
# ---------------------------------------------------------------------------

def set_corner_radius(shape: BaseShape, radius: int = 8000) -> None:
    """Adjust the corner radius of a ROUNDED_RECTANGLE shape.

    Parameters
    ----------
    radius : int
        Adjustment value (lower = less rounded).  Default 8000 gives a
        subtle, modern rounding.  PowerPoint default is ~16667.
    """
    try:
        sp = shape._element
        spPr = getattr(sp, "spPr", None)
        if spPr is None:
            return
        prstGeom = spPr.find(f"{{{_DML_NS}}}prstGeom")
        if prstGeom is None:
            return
        avLst = prstGeom.find(f"{{{_DML_NS}}}avLst")
        if avLst is None:
            avLst = etree.SubElement(prstGeom, f"{{{_DML_NS}}}avLst")
        # Clear existing guides
        for child in list(avLst):
            avLst.remove(child)
        gd = etree.SubElement(avLst, f"{{{_DML_NS}}}gd")
        gd.set("name", "adj")
        gd.set("fmla", f"val {radius}")
    except Exception:
        log.warning("Failed to set corner radius on shape %s", shape.name, exc_info=True)


# ---------------------------------------------------------------------------
# Transparency  (native python-pptx API)
# ---------------------------------------------------------------------------

def set_transparency(shape: BaseShape, percent: float = 20.0) -> None:
    """Set fill transparency on *shape* (0–100)."""
    try:
        shape.fill.transparency = percent / 100.0
    except Exception:
        log.warning("Failed to set transparency on shape %s", shape.name, exc_info=True)


# ---------------------------------------------------------------------------
# No-outline helper
# ---------------------------------------------------------------------------

def remove_outline(shape: BaseShape) -> None:
    """Remove the outline (border) from *shape*."""
    try:
        shape.line.fill.background()
    except Exception:
        log.warning("Failed to remove outline from shape %s", shape.name, exc_info=True)


# ---------------------------------------------------------------------------
# Convenience: shadow + gradient in one call
# ---------------------------------------------------------------------------

def style_card(
    shape: BaseShape,
    *,
    shadow_preset: str = "card",
    gradient_stops: Sequence[tuple[float, RGBColor]] | None = None,
    theme_color: MSO_THEME_COLOR | None = None,
    gradient_angle: float = 270.0,
    corner_radius: int | None = 8000,
    no_outline: bool = True,
) -> None:
    """Apply a full "card" style: shadow + gradient/solid + rounded corners.

    Supply *either* ``gradient_stops`` (RGB) or ``theme_color`` (scheme).
    If neither is given, only shadow + corner radius are applied.
    """
    add_shadow(shape, preset=shadow_preset)
    if gradient_stops is not None:
        add_gradient(shape, gradient_stops, angle=gradient_angle)
    elif theme_color is not None:
        add_theme_gradient(shape, theme_color, angle=gradient_angle)
    if corner_radius is not None:
        set_corner_radius(shape, corner_radius)
    if no_outline:
        remove_outline(shape)


def style_accent_bar(
    shape: BaseShape,
    *,
    gradient_stops: Sequence[tuple[float, RGBColor]] | None = None,
    theme_color: MSO_THEME_COLOR | None = None,
    angle: float = 0.0,
) -> None:
    """Style a thin accent/divider bar with gradient and no outline."""
    if gradient_stops is not None:
        add_gradient(shape, gradient_stops, angle=angle)
    elif theme_color is not None:
        add_theme_gradient(shape, theme_color, angle=angle, brightness_delta=0.3)
    remove_outline(shape)


def style_numbered_circle(
    shape: BaseShape,
    theme_color: MSO_THEME_COLOR = MSO_THEME_COLOR.ACCENT_1,
) -> None:
    """Style a small numbered circle (like the 01/02/03 in sample UAE output)."""
    add_theme_gradient(shape, theme_color, angle=270.0, brightness_delta=0.2)
    add_shadow(shape, preset="subtle")
    remove_outline(shape)
