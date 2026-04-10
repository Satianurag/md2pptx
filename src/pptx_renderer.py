from __future__ import annotations
import re
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData, XyChartData

from .schemas import (
    PresentationSpec, SlideSpec, SlideElement,
    TextContent, BulletContent, ChartContent, TableContent,
    ShapeContent, InfographicContent, SlideMasterInfo, ThemeColors,
)
from .slide_master import get_layout_for_slide_type, read_slide_master
from . import config

import logging
_log = logging.getLogger(__name__)


# ── Chart type mapping ──
CHART_TYPE_MAP = {
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "area": XL_CHART_TYPE.AREA,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}

ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}

# ── Shape type mapping ──
SHAPE_MAP = {
    "ROUNDED_RECTANGLE": MSO_SHAPE.ROUNDED_RECTANGLE,
    "RECTANGLE": MSO_SHAPE.RECTANGLE,
    "CHEVRON": MSO_SHAPE.CHEVRON,
    "RIGHT_ARROW": MSO_SHAPE.RIGHT_ARROW,
    "OVAL": MSO_SHAPE.OVAL,
    "PENTAGON": MSO_SHAPE.PENTAGON,
    "HEXAGON": MSO_SHAPE.HEXAGON,
    "DIAMOND": MSO_SHAPE.DIAMOND,
    "FLOWCHART_PROCESS": MSO_SHAPE.FLOWCHART_PROCESS,
    "FLOWCHART_DECISION": MSO_SHAPE.FLOWCHART_DECISION,
    "FLOWCHART_TERMINATOR": MSO_SHAPE.FLOWCHART_TERMINATOR,
    "ROUND_1_RECTANGLE": MSO_SHAPE.ROUND_1_RECTANGLE,
    "SNIP_1_RECTANGLE": MSO_SHAPE.SNIP_1_RECTANGLE,
}


def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert hex string like '1F77B4' to RGBColor."""
    hex_str = hex_str.lstrip('#')
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


# ── Theme color indices for shape fills (auto-inherit Slide Master palette) ──
_ACCENT_THEME_COLORS = [
    MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2,
    MSO_THEME_COLOR.ACCENT_3, MSO_THEME_COLOR.ACCENT_4,
    MSO_THEME_COLOR.ACCENT_5, MSO_THEME_COLOR.ACCENT_6,
]

# Fallback hex palette (only used when no template is loaded)
_FALLBACK_ACCENT_HEX = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47",
]


def _apply_accent_fill(shape, index: int, has_template: bool) -> None:
    """Apply an accent color fill to a shape, using theme colors when a template is loaded."""
    shape.fill.solid()
    if has_template:
        shape.fill.fore_color.theme_color = _ACCENT_THEME_COLORS[index % len(_ACCENT_THEME_COLORS)]
    else:
        shape.fill.fore_color.rgb = _hex_to_rgb(_FALLBACK_ACCENT_HEX[index % len(_FALLBACK_ACCENT_HEX)])


def _get_accent_hex(master_info: SlideMasterInfo | None, index: int) -> str:
    """Get accent hex color from theme or fallback."""
    if master_info:
        accents = master_info.theme_colors.accents()
        return accents[index % len(accents)]
    return _FALLBACK_ACCENT_HEX[index % len(_FALLBACK_ACCENT_HEX)]


def _remove_unused_placeholders(slide) -> None:
    """Remove all placeholder shapes that have no user-supplied text.
    Prevents ghost text from template placeholders bleeding through."""
    for shape in list(slide.placeholders):
        if shape.has_text_frame:
            text = shape.text_frame.text.strip() if shape.text_frame.text else ""
            if not text:
                try:
                    sp = shape._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass  # some placeholders can't be removed (e.g. slide number)


def _set_autofit(text_frame) -> None:
    """Set text frame to auto-shrink text to fit shape (PowerPoint renders on open)."""
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    text_frame.word_wrap = True


def _set_text_frame_text(text_frame, text: str, font_size=None, bold: bool | None = None,
                         alignment=None, color_rgb: RGBColor | None = None,
                         theme_color: MSO_THEME_COLOR | None = None) -> None:
    """Replace a text frame with a single formatted paragraph."""
    text_frame.clear()
    text_frame.word_wrap = True
    p = text_frame.paragraphs[0]
    p.text = text
    if font_size is not None:
        p.font.size = font_size
    if bold is not None:
        p.font.bold = bold
    if alignment is not None:
        p.alignment = alignment
    if theme_color is not None:
        p.font.color.theme_color = theme_color
    elif color_rgb is not None:
        p.font.color.rgb = color_rgb
    _set_autofit(text_frame)


def _populate_text_list(text_frame, items: list[str], font_size, prefix: str = "") -> None:
    """Populate a text frame with a concise multi-paragraph list."""
    text_frame.clear()
    text_frame.word_wrap = True
    for idx, item in enumerate(items):
        p = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
        p.text = f"{prefix}{item}" if prefix else item
        p.font.size = font_size
        p.alignment = PP_ALIGN.LEFT
        if idx > 0:
            p.space_before = Pt(8)
    _set_autofit(text_frame)


def _apply_light_surface_fill(shape, has_tpl: bool, brightness: float = 0.96,
                              fallback_hex: str = "F7F9FC") -> None:
    """Apply a subtle, low-noise surface fill for professional content containers."""
    shape.fill.solid()
    if has_tpl:
        shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        shape.fill.fore_color.brightness = brightness
    else:
        shape.fill.fore_color.rgb = _hex_to_rgb(fallback_hex)
    shape.line.fill.background()


def _remove_text_artifacts(slide) -> None:
    """Remove non-placeholder text shapes inherited from a layout to avoid ghost copy."""
    for shape in list(slide.shapes):
        if getattr(shape, "is_placeholder", False):
            continue
        if getattr(shape, "has_text_frame", False):
            text = shape.text_frame.text.strip() if shape.text_frame.text else ""
            if text:
                try:
                    sp = shape._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass


def _render_slide_title(slide, title: str, subtitle: str | None = None, has_tpl: bool = False) -> None:
    """Render title using the template title placeholder when available."""
    title_shape = slide.shapes.title if has_tpl else None
    if title_shape is not None and title_shape.has_text_frame:
        _set_text_frame_text(
            title_shape.text_frame,
            title,
            font_size=config.FONT_TITLE,
            bold=True,
            alignment=PP_ALIGN.LEFT,
        )
        return
    _add_title_bar(slide, title, subtitle, has_tpl)


def _apply_brand_card_fill(shape, index: int, has_tpl: bool) -> None:
    """Apply brand-aligned card background: ACCENT_1 with varying brightness.
    Keeps cards visually consistent instead of rainbow cycling."""
    shape.fill.solid()
    if has_tpl:
        shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        brightness = config.CARD_BRIGHTNESS_LEVELS[index % len(config.CARD_BRIGHTNESS_LEVELS)]
        shape.fill.fore_color.brightness = brightness
    else:
        shape.fill.fore_color.rgb = _hex_to_rgb("EDF2F9")


def _add_speaker_notes(slide, text: str) -> None:
    """Add speaker notes to a slide."""
    try:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = text[:2000]
    except Exception:
        pass  # notes slide creation can fail on some templates


def render_presentation(spec: PresentationSpec, output_path: str | Path) -> Path:
    """Render a full presentation from PresentationSpec to .pptx file."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    template_path = spec.template_path
    if template_path and Path(template_path).exists():
        prs = Presentation(str(template_path))
        # Remove existing slides from template (they're just examples)
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rId is None:
                # Fallback: try different attribute access
                slide_id = prs.slides._sldIdLst[0]
                rId = slide_id.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rId:
                prs.part.drop_rel(rId)
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
    else:
        prs = Presentation()

    master_info = read_slide_master(template_path) if template_path and Path(template_path).exists() else None

    for slide_spec in spec.slides:
        _render_slide(prs, slide_spec, master_info, deck_title=spec.title)

    prs.save(str(output_path))
    return output_path


def _get_slide_layout(prs: Presentation, slide_spec: SlideSpec, master_info: SlideMasterInfo | None):
    """Get the appropriate slide layout from the presentation."""
    if master_info:
        layout_info = get_layout_for_slide_type(master_info, slide_spec.slide_type)
        # Find the actual layout object by index
        master = prs.slide_masters[0]
        if layout_info.index < len(master.slide_layouts):
            return master.slide_layouts[layout_info.index]

    # Fallback: use layout by name matching or index
    for layout in prs.slide_layouts:
        if slide_spec.layout_name.lower() in layout.name.lower():
            return layout

    # Last fallback: blank-like layout (usually index 6 for blank in default)
    for layout in prs.slide_layouts:
        if "blank" in layout.name.lower():
            return layout

    return prs.slide_layouts[0]


def _render_slide(prs: Presentation, spec: SlideSpec, master_info: SlideMasterInfo | None,
                  deck_title: str = ""):
    """Render a single slide."""
    layout = _get_slide_layout(prs, spec, master_info)
    slide = prs.slides.add_slide(layout)
    has_tpl = master_info is not None

    # ── Cover slide ──
    if spec.slide_type == "cover":
        _render_cover(slide, spec, master_info, has_tpl)
        _add_speaker_notes(slide, f"Cover: {spec.title}")
        return

    # ── Thank you slide ──
    if spec.slide_type == "thank_you":
        _render_thank_you(slide, spec, master_info, has_tpl)
        _add_speaker_notes(slide, "Thank you slide — open for questions and discussion.")
        return

    # ── Section divider ──
    if spec.slide_type == "section_divider":
        _render_divider(slide, spec, has_tpl)
        _remove_unused_placeholders(slide)
        _add_speaker_notes(slide, f"Section: {spec.title}" + (f" — {spec.subtitle}" if spec.subtitle else ""))
        return

    # ── All other slides: populate title first, then add content, then clean placeholders ──
    if spec.title:
        _render_slide_title(slide, spec.title, spec.subtitle, has_tpl)

    for element in spec.elements:
        try:
            _render_element(slide, element, master_info, has_tpl, spec.slide_type)
        except Exception as e:
            _log.warning(f"Skipped element {element.element_type} on slide {spec.slide_number}: {e}")

    _remove_unused_placeholders(slide)

    # ── Slide furniture (footer, accent stripe) ──
    _add_slide_furniture(slide, spec, has_tpl, deck_title)

    # ── Speaker notes ──
    notes_text = spec.title or ""
    if spec.subtitle:
        notes_text += f"\n{spec.subtitle}"
    _add_speaker_notes(slide, notes_text)


# ── Slide type renderers ────────────────────────────────────────────

def _render_cover(slide, spec: SlideSpec, master_info: SlideMasterInfo | None = None,
                  has_tpl: bool = False):
    """Render cover slide using placeholders or fallback to shapes."""
    phs = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

    if phs:
        ph_list = sorted(phs.values(), key=lambda p: p.placeholder_format.idx)
        if len(ph_list) >= 1:
            ph_list[0].text = spec.title
            for para in ph_list[0].text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = config.FONT_TITLE
        if len(ph_list) >= 2 and spec.subtitle:
            ph_list[1].text = spec.subtitle
            for para in ph_list[1].text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = config.FONT_SUBTITLE
        # Remove remaining unused placeholders to prevent ghost text
        for ph in ph_list[2:]:
            if not ph.text.strip():
                try:
                    sp = ph._element
                    sp.getparent().remove(sp)
                except Exception:
                    ph.text = ""  # fallback: blank it out
        # Auto-fit title/subtitle text
        if len(ph_list) >= 1:
            _set_autofit(ph_list[0].text_frame)
        if len(ph_list) >= 2:
            _set_autofit(ph_list[1].text_frame)
    else:
        _add_textbox(slide, spec.title, config.MARGIN_LEFT, Emu(2500000),
                     config.CONTENT_WIDTH, Emu(800000),
                     font_size=config.FONT_TITLE, bold=True, alignment="center")
        if spec.subtitle:
            _add_textbox(slide, spec.subtitle, config.MARGIN_LEFT, Emu(3500000),
                         config.CONTENT_WIDTH, Emu(600000),
                         font_size=config.FONT_SUBTITLE, alignment="center")

    # Bottom accent bar on cover
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, config.SLIDE_HEIGHT - Emu(120000),
        config.SLIDE_WIDTH, Emu(120000)
    )
    accent.fill.solid()
    if has_tpl:
        accent.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        accent.fill.fore_color.rgb = _hex_to_rgb("4472C4")
    accent.line.fill.background()

    for element in spec.elements:
        _render_element(slide, element, master_info, has_tpl)


def _render_divider(slide, spec: SlideSpec, has_tpl: bool = False):
    """Render a section divider slide with title and optional subtitle."""
    phs = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
    if phs:
        ph_list = sorted(phs.values(), key=lambda p: p.placeholder_format.idx)
        if ph_list:
            ph_list[0].text = spec.title
            _set_autofit(ph_list[0].text_frame)
        # Fill subtitle into second placeholder or remove it
        if len(ph_list) >= 2:
            if spec.subtitle:
                ph_list[1].text = spec.subtitle
                _set_autofit(ph_list[1].text_frame)
            else:
                try:
                    sp = ph_list[1]._element
                    sp.getparent().remove(sp)
                except Exception:
                    ph_list[1].text = ""
        # Remove any remaining unused placeholders
        for ph in ph_list[2:]:
            try:
                sp = ph._element
                sp.getparent().remove(sp)
            except Exception:
                ph.text = ""
    else:
        _add_textbox(slide, spec.title, config.MARGIN_LEFT, Emu(2800000),
                     config.CONTENT_WIDTH, Emu(800000),
                     font_size=config.FONT_TITLE, bold=True, alignment="center")

    # Subtitle textbox (if no placeholders handled it)
    if spec.subtitle and not phs:
        _add_textbox(slide, spec.subtitle, config.MARGIN_LEFT, Emu(3600000),
                     config.CONTENT_WIDTH, Emu(500000),
                     font_size=config.FONT_SUBTITLE, alignment="center", color="666666")

    # Accent bar under title
    bar_y = Emu(4200000) if spec.subtitle else Emu(3700000)
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        config.MARGIN_LEFT + Emu(3000000), bar_y,
        Emu(5000000), Emu(40000)
    )
    bar.fill.solid()
    if has_tpl:
        bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        bar.fill.fore_color.rgb = _hex_to_rgb("4472C4")
    bar.line.fill.background()


def _render_thank_you(slide, spec: SlideSpec, master_info: SlideMasterInfo | None = None,
                      has_tpl: bool = False):
    """Render thank you slide — use EITHER placeholders OR textboxes, never both."""
    title_text = spec.title or "Thank You"
    subtitle_text = spec.subtitle or "Questions & Discussion"

    # Try placeholders first
    phs = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
    used_phs = set()

    if phs:
        ph_list = sorted(phs.values(), key=lambda p: p.placeholder_format.idx)
        # Fill title placeholder
        if len(ph_list) >= 1:
            ph_list[0].text = title_text
            _set_autofit(ph_list[0].text_frame)
            used_phs.add(ph_list[0].placeholder_format.idx)
        # Fill subtitle placeholder
        if len(ph_list) >= 2:
            ph_list[1].text = subtitle_text
            _set_autofit(ph_list[1].text_frame)
            used_phs.add(ph_list[1].placeholder_format.idx)
        # Remove ALL unused placeholders — this prevents ghost text
        for ph in ph_list:
            if ph.placeholder_format.idx not in used_phs:
                try:
                    sp = ph._element
                    sp.getparent().remove(sp)
                except Exception:
                    ph.text = ""
    else:
        # No placeholders at all — use manual textboxes
        _remove_text_artifacts(slide)
        _add_textbox(slide, title_text, config.MARGIN_LEFT + Emu(600000), Emu(2350000),
                     config.CONTENT_WIDTH - Emu(1200000), Emu(800000),
                     font_size=Pt(36), bold=True, alignment="center")
        _add_textbox(slide, subtitle_text, config.MARGIN_LEFT + Emu(600000), Emu(3200000),
                     config.CONTENT_WIDTH - Emu(1200000), Emu(400000),
                     font_size=config.FONT_SUBTITLE, alignment="center",
                     color="666666")


# ── Title bar ───────────────────────────────────────────────────────

def _add_title_bar(slide, title: str, subtitle: str | None = None, has_tpl: bool = False):
    """Add a title bar at the top of a content slide with accent underline."""
    # Title
    txBox = slide.shapes.add_textbox(
        config.MARGIN_LEFT, config.MARGIN_TOP,
        config.CONTENT_WIDTH, Emu(530000)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = config.FONT_TITLE
    p.font.bold = True
    p.alignment = PP_ALIGN.LEFT

    # Subtitle
    if subtitle:
        txBox2 = slide.shapes.add_textbox(
            config.MARGIN_LEFT, Emu(config.MARGIN_TOP + 550000),
            config.CONTENT_WIDTH, Emu(350000)
        )
        tf2 = txBox2.text_frame
        tf2.word_wrap = True
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.font.size = config.FONT_SUBTITLE
        p2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        p2.alignment = PP_ALIGN.LEFT

    # Accent line under title
    accent_y = Emu(config.MARGIN_TOP + (900000 if subtitle else 560000))
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, config.MARGIN_LEFT, accent_y,
        Emu(2000000), Emu(36000)
    )
    accent.fill.solid()
    if has_tpl:
        accent.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        accent.fill.fore_color.rgb = _hex_to_rgb("4472C4")
    accent.line.fill.background()


# ── Slide furniture (footer bar, accent stripe) ──────────────────────

def _add_slide_furniture(slide, spec: SlideSpec, has_tpl: bool, deck_title: str):
    """Add footer bar with deck title + slide number, and left accent stripe."""
    if has_tpl:
        return
    # ── Footer bar ──
    footer_h = Emu(300000)
    footer_y = config.SLIDE_HEIGHT - footer_h
    footer = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, footer_y, config.SLIDE_WIDTH, footer_h
    )
    footer.fill.solid()
    if has_tpl:
        footer.fill.fore_color.theme_color = MSO_THEME_COLOR.DARK_2
        footer.fill.fore_color.brightness = 0.85
    else:
        footer.fill.fore_color.rgb = _hex_to_rgb("F2F2F2")
    footer.line.fill.background()

    # Footer text — deck title
    if deck_title:
        ft = slide.shapes.add_textbox(
            config.MARGIN_LEFT, footer_y + Emu(60000),
            Emu(8000000), Emu(200000)
        )
        tf = ft.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = deck_title[:80]
        p.font.size = Pt(8)
        p.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        p.alignment = PP_ALIGN.LEFT

    # Slide number
    sn = slide.shapes.add_textbox(
        Emu(config.SLIDE_WIDTH - 1000000), footer_y + Emu(60000),
        Emu(700000), Emu(200000)
    )
    tf2 = sn.text_frame
    tf2.word_wrap = False
    p2 = tf2.paragraphs[0]
    p2.text = str(spec.slide_number)
    p2.font.size = Pt(9)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    p2.alignment = PP_ALIGN.RIGHT

    # ── Left accent stripe ──
    stripe = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 0, 0, Emu(60000), config.SLIDE_HEIGHT - footer_h
    )
    stripe.fill.solid()
    if has_tpl:
        stripe.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        stripe.fill.fore_color.rgb = _hex_to_rgb("4472C4")
    stripe.line.fill.background()


# ── Element renderer dispatch ───────────────────────────────────────

def _render_element(slide, element: SlideElement,
                    master_info: SlideMasterInfo | None = None,
                    has_tpl: bool = False,
                    slide_type: str = "content"):
    """Dispatch rendering based on element type."""
    pos = element.position
    content = element.content

    if isinstance(content, TextContent):
        _render_text(slide, pos, content)
    elif isinstance(content, BulletContent):
        _render_bullets(slide, pos, content, has_tpl, slide_type)
    elif isinstance(content, ChartContent):
        _render_chart(slide, pos, content, master_info, has_tpl)
    elif isinstance(content, TableContent):
        _render_table(slide, pos, content, master_info, has_tpl)
    elif isinstance(content, ShapeContent):
        _render_shape(slide, pos, content)
    elif isinstance(content, InfographicContent):
        _render_infographic(slide, pos, content, has_tpl)


# ── Text rendering ──────────────────────────────────────────────────

def _add_textbox(slide, text, left, top, width, height,
                 font_size=None, bold=False, italic=False,
                 color=None, alignment=None, autofit=True):
    """Helper to add a text box with formatting and optional auto-fit."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    if font_size:
        p.font.size = font_size
    if bold:
        p.font.bold = True
    if italic:
        p.font.italic = True
    if color:
        p.font.color.rgb = _hex_to_rgb(color) if isinstance(color, str) else color
    if alignment and alignment in ALIGN_MAP:
        p.alignment = ALIGN_MAP[alignment]
    if autofit:
        _set_autofit(tf)
    else:
        tf.auto_size = None
    return txBox


def _render_text(slide, pos, content: TextContent):
    """Render a text element."""
    _add_textbox(
        slide, content.text,
        pos.left, pos.top, pos.width, pos.height,
        font_size=Pt(content.font_size) if content.font_size else config.FONT_BODY,
        bold=content.bold, italic=content.italic,
        color=content.color, alignment=content.alignment,
    )


def _render_bullets(slide, pos, content: BulletContent, has_tpl: bool = False,
                    slide_type: str = "content"):
    """Render bullets with slide-type-specific, lower-noise layouts."""
    items = [item.strip() for item in content.items if item and item.strip()]
    if not items:
        return

    font_size = Pt(content.font_size) if content.font_size else config.FONT_BODY

    if slide_type == "agenda":
        _render_agenda_bullets(slide, pos, items, font_size, has_tpl)
        return

    if slide_type in ("executive_summary", "conclusion"):
        _render_summary_bullets(slide, pos, items, font_size, has_tpl)
        return

    _render_content_bullets(slide, pos, items, font_size, has_tpl)


def _render_agenda_bullets(slide, pos, items: list[str], font_size, has_tpl: bool) -> None:
    cols = 2 if len(items) > 4 else 1
    rows = (len(items) + cols - 1) // cols
    gap_h = Emu(180000)
    gap_v = Emu(140000)
    card_w = (pos.width - gap_h * (cols - 1)) // cols
    card_h = min((pos.height - gap_v * (rows - 1)) // max(rows, 1), Emu(720000))

    for idx, item in enumerate(items):
        col = idx % cols
        row = idx // cols
        x = pos.left + col * (card_w + gap_h)
        y = pos.top + row * (card_h + gap_v)

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
        _apply_light_surface_fill(card, has_tpl, brightness=0.98, fallback_hex="FBFCFE")

        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, x + Emu(90000), y + Emu(110000), Emu(220000), Emu(220000))
        badge.fill.solid()
        if has_tpl:
            badge.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        else:
            badge.fill.fore_color.rgb = _hex_to_rgb("4472C4")
        badge.line.fill.background()
        _set_text_frame_text(
            badge.text_frame,
            str(idx + 1),
            font_size=Pt(10),
            bold=True,
            alignment=PP_ALIGN.CENTER,
            color_rgb=RGBColor(0xFF, 0xFF, 0xFF),
        )

        tx = slide.shapes.add_textbox(x + Emu(380000), y + Emu(90000), card_w - Emu(500000), card_h - Emu(180000))
        _populate_text_list(tx.text_frame, [item], font_size)


def _render_summary_bullets(slide, pos, items: list[str], font_size, has_tpl: bool) -> None:
    cols = 2 if len(items) >= 4 else 1
    rows = (len(items) + cols - 1) // cols
    gap_h = Emu(180000)
    gap_v = Emu(160000)
    card_w = (pos.width - gap_h * (cols - 1)) // cols
    card_h = min((pos.height - gap_v * (rows - 1)) // max(rows, 1), Emu(1000000))

    for idx, item in enumerate(items):
        col = idx % cols
        row = idx // cols
        x = pos.left + col * (card_w + gap_h)
        y = pos.top + row * (card_h + gap_v)

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w, card_h)
        _apply_light_surface_fill(card, has_tpl, brightness=0.97, fallback_hex="F8FAFD")

        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, card_w, Emu(50000))
        accent.fill.solid()
        if has_tpl:
            accent.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        else:
            accent.fill.fore_color.rgb = _hex_to_rgb("4472C4")
        accent.line.fill.background()

        tx = slide.shapes.add_textbox(x + Emu(120000), y + Emu(120000), card_w - Emu(240000), card_h - Emu(180000))
        _populate_text_list(tx.text_frame, [item], font_size)


def _render_content_bullets(slide, pos, items: list[str], font_size, has_tpl: bool) -> None:
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pos.left, pos.top, pos.width, pos.height)
    _apply_light_surface_fill(panel, has_tpl, brightness=0.985, fallback_hex="FBFCFE")

    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, pos.left, pos.top, Emu(70000), pos.height)
    stripe.fill.solid()
    if has_tpl:
        stripe.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        stripe.fill.fore_color.rgb = _hex_to_rgb("4472C4")
    stripe.line.fill.background()

    inner_left = pos.left + Emu(180000)
    inner_top = pos.top + Emu(100000)
    inner_width = pos.width - Emu(260000)
    inner_height = pos.height - Emu(180000)

    split_columns = len(items) >= 5 and pos.width >= Emu(7000000)
    if split_columns:
        gap = Emu(180000)
        col_width = (inner_width - gap) // 2
        midpoint = (len(items) + 1) // 2
        left_box = slide.shapes.add_textbox(inner_left, inner_top, col_width, inner_height)
        right_box = slide.shapes.add_textbox(inner_left + col_width + gap, inner_top, col_width, inner_height)
        _populate_text_list(left_box.text_frame, items[:midpoint], font_size, prefix="• ")
        _populate_text_list(right_box.text_frame, items[midpoint:], font_size, prefix="• ")
        return

    tx = slide.shapes.add_textbox(inner_left, inner_top, inner_width, inner_height)
    _populate_text_list(tx.text_frame, items, font_size, prefix="• ")


# ── Chart rendering ─────────────────────────────────────────────────

def _render_chart(slide, pos, content: ChartContent,
                  master_info: SlideMasterInfo | None = None, has_tpl: bool = False):
    """Render a native PowerPoint chart."""
    chart_type = CHART_TYPE_MAP.get(content.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    if content.chart_type == "scatter":
        chart_data = XyChartData()
        for series in content.series:
            s = chart_data.add_series(series.name)
            for i, val in enumerate(series.values):
                x = float(i)
                s.add_data_point(x, val)
    else:
        chart_data = CategoryChartData()
        chart_data.categories = content.categories
        for series in content.series:
            chart_data.add_series(series.name, tuple(series.values))

    graphic_frame = slide.shapes.add_chart(
        chart_type, pos.left, pos.top, pos.width, pos.height, chart_data
    )
    chart = graphic_frame.chart

    # Title
    if content.title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = content.title
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(11)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True

    # Legend — always show for multi-series, and for pie/doughnut
    if len(content.series) > 1 or content.chart_type in ("pie", "doughnut"):
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        try:
            chart.legend.font.size = Pt(9)
        except Exception:
            pass

    # Data labels only when they help readability rather than creating clutter.
    try:
        plot = chart.plots[0]
        show_data_labels = (
            content.chart_type in ("pie", "doughnut")
            or (len(content.series) == 1 and len(content.categories) <= 6)
        )
        plot.has_data_labels = show_data_labels
        if show_data_labels:
            data_labels = plot.data_labels
            data_labels.font.size = Pt(9)
            if content.chart_type in ("pie", "doughnut"):
                data_labels.number_format = '0%'
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif content.chart_type in ("bar", "column"):
                data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            elif content.chart_type == "line":
                data_labels.position = XL_LABEL_POSITION.ABOVE
    except Exception:
        pass  # some chart types don't support all label positions

    # Axis formatting (skip for pie/doughnut which have no axes)
    if content.chart_type not in ("pie", "doughnut"):
        try:
            # Category axis
            cat_axis = chart.category_axis
            cat_axis.has_major_gridlines = False
            cat_axis.tick_labels.font.size = Pt(9)
            # Value axis
            val_axis = chart.value_axis
            val_axis.has_major_gridlines = True
            val_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
            val_axis.has_minor_gridlines = False
            val_axis.tick_labels.font.size = Pt(9)
        except Exception:
            pass  # scatter charts may have different axis structure

    # Color series using theme colors (full accent cycling for charts)
    for i, series in enumerate(chart.series):
        fill = series.format.fill
        fill.solid()
        if has_tpl:
            fill.fore_color.theme_color = _ACCENT_THEME_COLORS[i % len(_ACCENT_THEME_COLORS)]
        else:
            fill.fore_color.rgb = _hex_to_rgb(_FALLBACK_ACCENT_HEX[i % len(_FALLBACK_ACCENT_HEX)])


# ── Table rendering ─────────────────────────────────────────────────

def _render_table(slide, pos, content: TableContent,
                  master_info: SlideMasterInfo | None = None,
                  has_tpl: bool = False):
    """Render a formatted table."""
    rows = len(content.rows) + 1  # +1 for header
    cols = len(content.headers)
    if cols == 0:
        return

    table_shape = slide.shapes.add_table(rows, cols, pos.left, pos.top, pos.width, pos.height)
    table = table_shape.table

    # Calculate column widths
    if content.col_widths:
        total = sum(content.col_widths)
        for i, w in enumerate(content.col_widths):
            if i < cols:
                table.columns[i].width = int(pos.width * w / total)
    else:
        col_width = pos.width // cols
        for i in range(cols):
            table.columns[i].width = col_width

    # Resolve header color from theme
    header_hex = _get_accent_hex(master_info, 0) if master_info else "1F4E79"
    # Derive a lighter shade for alternating rows
    alt_row_hex = _get_accent_hex(master_info, 0) if master_info else "D6E4F0"

    # Header row
    for col_idx, header in enumerate(content.headers):
        cell = table.cell(0, col_idx)
        cell.text = header
        para = cell.text_frame.paragraphs[0]
        para.font.bold = True
        para.font.size = Pt(10)
        para.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
        cell.fill.solid()
        if has_tpl:
            cell.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        else:
            cell.fill.fore_color.rgb = _hex_to_rgb(header_hex)
        para.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # Cell padding
        cell.margin_left = Emu(70000)
        cell.margin_right = Emu(70000)
        cell.margin_top = Emu(40000)
        cell.margin_bottom = Emu(40000)

    # Data rows
    for row_idx, row_data in enumerate(content.rows):
        for col_idx, value in enumerate(row_data):
            if col_idx < cols:
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(value)
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(10)
                para.alignment = PP_ALIGN.LEFT if col_idx == 0 else PP_ALIGN.CENTER
                # Ultra-light alternating row shading (0.92 brightness)
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    if has_tpl:
                        cell.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
                        cell.fill.fore_color.brightness = 0.92
                    else:
                        cell.fill.fore_color.rgb = _hex_to_rgb("EDF2F9")
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                # Cell padding
                cell.margin_left = Emu(70000)
                cell.margin_right = Emu(70000)
                cell.margin_top = Emu(40000)
                cell.margin_bottom = Emu(40000)
                # Auto-fit cell text
                _set_autofit(cell.text_frame)


# ── Shape rendering ─────────────────────────────────────────────────

def _render_shape(slide, pos, content: ShapeContent):
    """Render an auto shape."""
    shape_type = SHAPE_MAP.get(content.shape_type, MSO_SHAPE.ROUNDED_RECTANGLE)
    shape = slide.shapes.add_shape(shape_type, pos.left, pos.top, pos.width, pos.height)

    if content.text:
        shape.text = content.text
        tf = shape.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        for para in tf.paragraphs:
            para.font.size = Pt(content.font_size) if content.font_size else config.FONT_BODY
            if content.bold:
                para.font.bold = True
            para.alignment = PP_ALIGN.CENTER
        shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    if content.fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _hex_to_rgb(content.fill_color)
    else:
        shape.fill.background()  # transparent

    if content.line_color:
        shape.line.color.rgb = _hex_to_rgb(content.line_color)
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()  # no border


# ── Infographic rendering ───────────────────────────────────────────

def _render_infographic(slide, pos, content: InfographicContent,
                       has_tpl: bool = False):
    """Render an infographic based on type."""
    if content.infographic_type == "process_flow":
        _render_process_flow(slide, pos, content.items, has_tpl)
    elif content.infographic_type == "timeline":
        _render_timeline(slide, pos, content.items, has_tpl)
    elif content.infographic_type == "comparison":
        _render_comparison(slide, pos, content.items, has_tpl)
    elif content.infographic_type == "kpi_cards":
        _render_kpi_cards(slide, pos, content.items, has_tpl)
    elif content.infographic_type == "hierarchy":
        _render_hierarchy(slide, pos, content.items, has_tpl)


def _render_process_flow(slide, pos, items, has_tpl: bool = False):
    """Render a process flow with rounded rectangles, step number overlays, and arrow connectors."""
    n = len(items)
    if n == 0:
        return

    # Brand-aligned: cycle through only 2-3 colors
    _FLOW_ACCENTS = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_6, MSO_THEME_COLOR.ACCENT_2]
    _FLOW_FALLBACK = ["4472C4", "70AD47", "ED7D31"]

    arrow_gap = Emu(200000)  # space for arrow text between boxes
    usable_w = pos.width - arrow_gap * max(n - 1, 0)
    item_width = min(usable_w // n, Emu(2800000))  # cap width at ~3 inches

    # Center the flow horizontally
    total_w = item_width * n + arrow_gap * max(n - 1, 0)
    x_offset = pos.left + (pos.width - total_w) // 2

    item_height = Emu(min(pos.height, 1200000))
    y_center = pos.top + (pos.height - item_height) // 2
    step_circle_size = Emu(240000)

    for i, item in enumerate(items):
        x = x_offset + i * (item_width + arrow_gap)

        # ── Rounded rectangle (not chevron — readable at any width) ──
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, x, y_center, item_width, item_height
        )
        shape.fill.solid()
        if has_tpl:
            shape.fill.fore_color.theme_color = _FLOW_ACCENTS[i % len(_FLOW_ACCENTS)]
        else:
            shape.fill.fore_color.rgb = _hex_to_rgb(_FLOW_FALLBACK[i % len(_FLOW_FALLBACK)])
        shape.line.fill.background()

        # ── Step number circle overlay (top-left of rectangle) ──
        sc_x = x - step_circle_size // 3
        sc_y = y_center - step_circle_size // 3
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, sc_x, sc_y, step_circle_size, step_circle_size)
        circle.fill.solid()
        circle.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        circle.line.fill.background()
        ctf = circle.text_frame
        ctf.word_wrap = False
        cp = ctf.paragraphs[0]
        cp.text = str(i + 1)
        cp.font.size = Pt(11)
        cp.font.bold = True
        if has_tpl:
            cp.font.color.theme_color = _FLOW_ACCENTS[i % len(_FLOW_ACCENTS)]
        else:
            cp.font.color.rgb = _hex_to_rgb(_FLOW_FALLBACK[i % len(_FLOW_FALLBACK)])
        cp.alignment = PP_ALIGN.CENTER

        # ── Title + description inside the rectangle ──
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(60000)
        tf.margin_right = Emu(60000)
        tf.margin_top = Emu(80000)
        _set_autofit(tf)

        p = tf.paragraphs[0]
        p.text = item.title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.CENTER

        if item.description:
            p2 = tf.add_paragraph()
            p2.text = item.description[:100]
            p2.font.size = Pt(10)
            p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(6)

        # ── Arrow connector (clean text '→' instead of shape) ──
        if i < n - 1:
            ax = x + item_width
            ay = y_center + item_height // 2 - Emu(120000)
            _add_textbox(slide, "→", ax, ay, arrow_gap, Emu(240000),
                         font_size=Pt(20), bold=True, alignment="center",
                         color="999999", autofit=False)


def _render_timeline(slide, pos, items, has_tpl: bool = False):
    """Render a horizontal timeline with alternating above/below labels."""
    n = len(items)
    if n == 0:
        return

    # Horizontal line
    line_y = pos.top + pos.height // 2
    connector = slide.shapes.add_connector(
        1, pos.left, line_y, pos.left + pos.width, line_y
    )
    if has_tpl:
        connector.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
    else:
        connector.line.color.rgb = _hex_to_rgb("2E75B6")
    connector.line.width = Pt(2)

    # Nodes
    node_gap = pos.width // max(n, 1)
    circle_size = Emu(250000)

    for i, item in enumerate(items):
        cx = pos.left + i * node_gap + node_gap // 2
        # Circle marker (uniform ACCENT_1 for brand consistency)
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, cx - circle_size // 2, line_y - circle_size // 2,
            circle_size, circle_size
        )
        circle.fill.solid()
        if has_tpl:
            circle.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        else:
            circle.fill.fore_color.rgb = _hex_to_rgb("4472C4")
        circle.line.fill.background()

        # Step number inside circle
        tf = circle.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.text = str(i + 1)
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.CENTER

        # Connecting line segment from circle to label
        is_above = (i % 2 == 0)
        label_y = line_y - Emu(600000) if is_above else line_y + Emu(350000)
        conn_start_y = line_y - circle_size // 2 if is_above else line_y + circle_size // 2
        conn_end_y = label_y + Emu(400000) if is_above else label_y
        try:
            seg = slide.shapes.add_connector(1, cx, conn_start_y, cx, conn_end_y)
            if has_tpl:
                seg.line.color.theme_color = MSO_THEME_COLOR.ACCENT_1
            else:
                seg.line.color.rgb = _hex_to_rgb("4472C4")
            seg.line.width = Pt(1)
            seg.line.dash_style = 2  # dash
        except Exception:
            pass

        # Label above or below (alternate)
        label_w = Emu(min(node_gap - 50000, 1800000))
        _add_textbox(slide, item.title, cx - label_w // 2, label_y,
                     label_w, Emu(400000),
                     font_size=Pt(9), bold=True, alignment="center")

        if item.description:
            desc_y = label_y + Emu(300000)
            _add_textbox(slide, item.description[:100], cx - label_w // 2, desc_y,
                         label_w, Emu(300000),
                         font_size=Pt(8), alignment="center")


def _render_comparison(slide, pos, items, has_tpl: bool = False):
    """Render side-by-side comparison cards with brand-aligned colors."""
    _CMP_ACCENTS = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_6, MSO_THEME_COLOR.DARK_2]
    _CMP_FALLBACK = ["4472C4", "70AD47", "44546A"]
    n = len(items)
    if n == 0:
        return
    gap = Emu(120000)
    card_width = (pos.width - gap * (n - 1)) // n
    card_height = pos.height

    for i, item in enumerate(items):
        x = pos.left + i * (card_width + gap)

        # Card background (brand-aligned: 3-color rotation)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, x, pos.top, card_width, card_height
        )
        card.fill.solid()
        if has_tpl:
            card.fill.fore_color.theme_color = _CMP_ACCENTS[i % len(_CMP_ACCENTS)]
        else:
            card.fill.fore_color.rgb = _hex_to_rgb(_CMP_FALLBACK[i % len(_CMP_FALLBACK)])
        card.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(60000)
        tf.margin_right = Emu(60000)
        tf.margin_top = Emu(50000)
        _set_autofit(tf)

        p = tf.paragraphs[0]
        p.text = item.title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.CENTER

        if item.description:
            p2 = tf.add_paragraph()
            p2.text = item.description[:120]
            p2.font.size = Pt(9)
            p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p2.alignment = PP_ALIGN.CENTER
            p2.space_before = Pt(8)

        if item.value:
            p3 = tf.add_paragraph()
            p3.text = item.value
            p3.font.size = Pt(18)
            p3.font.bold = True
            p3.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p3.alignment = PP_ALIGN.CENTER
            p3.space_before = Pt(12)


def _render_kpi_cards(slide, pos, items, has_tpl: bool = False):
    """Render KPI metric cards with value, title, and vertical centering.
    Uses 3-color brand rotation: ACCENT_1 + ACCENT_6 + DARK_2."""
    _KPI_ACCENTS = [MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_6, MSO_THEME_COLOR.DARK_2]
    _KPI_FALLBACK = ["4472C4", "70AD47", "44546A"]
    n = len(items)
    if n == 0:
        return
    gap = Emu(120000)
    card_width = (pos.width - gap * (n - 1)) // n
    card_height = min(pos.height, Emu(1800000))
    y = pos.top + (pos.height - card_height) // 2

    for i, item in enumerate(items):
        x = pos.left + i * (card_width + gap)

        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_width, card_height
        )
        card.fill.solid()
        if has_tpl:
            card.fill.fore_color.theme_color = _KPI_ACCENTS[i % len(_KPI_ACCENTS)]
        else:
            card.fill.fore_color.rgb = _hex_to_rgb(_KPI_FALLBACK[i % len(_KPI_FALLBACK)])
        card.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_top = Emu(100000)
        tf.margin_left = Emu(50000)
        tf.margin_right = Emu(50000)
        _set_autofit(tf)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].space_before = Pt(16)

        # Big value
        p = tf.paragraphs[0]
        p.text = item.value or ""
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.CENTER

        # Label below
        p2 = tf.add_paragraph()
        p2.text = item.title
        p2.font.size = Pt(11)
        p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(8)

        # Description if present
        if item.description:
            p3 = tf.add_paragraph()
            p3.text = item.description[:80]
            p3.font.size = Pt(8)
            p3.font.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
            p3.alignment = PP_ALIGN.CENTER


def _render_hierarchy(slide, pos, items, has_tpl: bool = False):
    """Render a hierarchy with indented boxes and connectors."""
    n = len(items)
    if n == 0:
        return
    item_height = min(pos.height // n - Emu(50000), Emu(800000))

    for i, item in enumerate(items):
        indent = Emu(i * 200000)
        y = pos.top + i * (item_height + Emu(60000))
        w = pos.width - indent

        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, pos.left + indent, y, w, item_height
        )
        _apply_accent_fill(shape, i, has_tpl)
        shape.line.fill.background()

        # Vertical connector to next level
        if i < n - 1:
            cx = pos.left + indent + Emu(100000)
            cy = y + item_height
            next_y = pos.top + (i + 1) * (item_height + Emu(60000))
            connector = slide.shapes.add_connector(
                1, cx, cy, cx, next_y
            )
            connector.line.color.rgb = RGBColor(0x99, 0x99, 0x99)
            connector.line.width = Pt(1)

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(60000)
        tf.margin_right = Emu(60000)
        _set_autofit(tf)
        p = tf.paragraphs[0]
        p.text = item.title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.alignment = PP_ALIGN.LEFT

        if item.description:
            p2 = tf.add_paragraph()
            p2.text = item.description[:100]
            p2.font.size = Pt(9)
            p2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
