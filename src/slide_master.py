from __future__ import annotations
import zipfile
import logging
from pathlib import Path
from lxml import etree
from pptx import Presentation
from .schemas import SlideMasterInfo, LayoutInfo, PlaceholderInfo, ThemeColors
from . import config

logger = logging.getLogger(__name__)

_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS = {"a": _DRAWINGML_NS}

_COLOR_SLOTS = (
    "dk1", "lt1", "dk2", "lt2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
)


def _extract_theme_colors(template_path: Path) -> ThemeColors:
    """Parse the PPTX theme XML and return all 12 color-scheme slots."""
    colors: dict[str, str] = {}
    try:
        with zipfile.ZipFile(str(template_path), "r") as z:
            theme_files = sorted(
                f for f in z.namelist()
                if f.startswith("ppt/theme/theme") and f.endswith(".xml")
            )
            if not theme_files:
                return ThemeColors()
            theme_xml = z.read(theme_files[0])

        root = etree.fromstring(theme_xml)
        clr_scheme = root.find(".//a:clrScheme", _NS)
        if clr_scheme is None:
            return ThemeColors()

        for child in clr_scheme:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag not in _COLOR_SLOTS:
                continue
            srgb = child.find("a:srgbClr", _NS)
            sys_clr = child.find("a:sysClr", _NS)
            if srgb is not None:
                colors[tag] = srgb.get("val", "000000")
            elif sys_clr is not None:
                colors[tag] = sys_clr.get("lastClr", "000000")

    except Exception as e:
        logger.warning(f"Failed to extract theme colors: {e}")

    return ThemeColors(**colors) if colors else ThemeColors()


def _categorize_layout(name: str, placeholders: list[PlaceholderInfo]) -> str:
    """Categorize a layout based on its name and placeholder structure."""
    name_lower = name.lower().strip()

    if "cover" in name_lower or "title company" in name_lower:
        return "cover"
    if "divider" in name_lower or "section" in name_lower:
        return "divider"
    if "thank" in name_lower:
        return "thank_you"
    if name_lower == "blank" or (not placeholders and "blank" in name_lower):
        return "blank"
    if "title only" in name_lower:
        return "title_only"

    # Categorize by placeholder count and types
    has_title = any(p.ph_type == "TITLE" for p in placeholders)
    body_count = sum(1 for p in placeholders if p.ph_type in ("BODY", "OBJECT"))

    if not placeholders:
        return "blank"
    if has_title and body_count >= 2:
        return "two_content"
    if has_title and body_count == 1:
        return "title_content"
    if has_title and body_count == 0:
        return "title_only"

    return "other"


def read_slide_master(template_path: str | Path) -> SlideMasterInfo:
    """Read a .pptx template and extract layouts, placeholders, and theme colors."""
    template_path = Path(template_path)
    prs = Presentation(str(template_path))

    theme_colors = _extract_theme_colors(template_path)
    logger.info(f"Theme accents: {theme_colors.accents()}")

    layouts: list[LayoutInfo] = []
    master = prs.slide_masters[0]

    for idx, layout in enumerate(master.slide_layouts):
        placeholders: list[PlaceholderInfo] = []
        for ph in layout.placeholders:
            fmt = ph.placeholder_format
            placeholders.append(PlaceholderInfo(
                idx=fmt.idx,
                name=ph.name,
                ph_type=str(fmt.type).split(".")[-1].split("(")[0].strip() if fmt.type else None,
                left=ph.left or 0,
                top=ph.top or 0,
                width=ph.width or 0,
                height=ph.height or 0,
            ))

        category = _categorize_layout(layout.name, placeholders)
        layouts.append(LayoutInfo(
            index=idx,
            name=layout.name,
            category=category,
            placeholders=placeholders,
        ))
        logger.debug(f"Layout [{idx}] '{layout.name}' → {category} ({len(placeholders)} phs)")

    return SlideMasterInfo(
        template_path=str(template_path),
        slide_width=prs.slide_width,
        slide_height=prs.slide_height,
        layouts=layouts,
        theme_colors=theme_colors,
    )


def find_layout_by_category(master_info: SlideMasterInfo, category: str) -> LayoutInfo | None:
    """Find the first layout matching a given category."""
    for layout in master_info.layouts:
        if layout.category == category:
            return layout
    return None


def get_layout_for_slide_type(master_info: SlideMasterInfo, slide_type: str) -> LayoutInfo:
    """Map a slide_type from SlideSpec to the best available layout."""
    category_map: dict[str, list[str]] = {
        "cover": ["cover"],
        "section_divider": ["divider"],
        "thank_you": ["thank_you"],
        "agenda": ["title_only", "title_content", "blank"],
        "executive_summary": ["title_only", "title_content", "blank"],
        "content": ["title_only", "title_content", "blank"],
        "chart": ["title_only", "blank"],
        "table": ["title_only", "blank"],
        "infographic": ["title_only", "blank"],
        "mixed": ["title_only", "two_content", "blank"],
        "conclusion": ["title_only", "title_content", "blank"],
    }

    candidates = category_map.get(slide_type, ["blank"])
    for cat in candidates:
        layout = find_layout_by_category(master_info, cat)
        if layout is not None:
            return layout

    # Fallback to blank, then first available
    blank = find_layout_by_category(master_info, "blank")
    return blank if blank is not None else master_info.layouts[0]


def auto_detect_template(md_filename: str) -> Path | None:
    """Find the best matching template using fuzzy word-overlap scoring."""
    import re
    templates_dir = config.TEMPLATES_DIR
    if not templates_dir.exists():
        return None

    # Normalize filename into keyword set
    md_stem = Path(md_filename).stem.lower()
    md_words = set(re.findall(r'[a-z]{3,}', md_stem))

    best_match = None
    best_score = 0

    for tpl in templates_dir.glob("*.pptx"):
        tpl_stem = tpl.stem.lower().replace("template_", "")
        tpl_words = set(re.findall(r'[a-z]{3,}', tpl_stem))

        # Word overlap score
        overlap = len(md_words & tpl_words)
        # Substring bonus
        if tpl_stem in md_stem or md_stem in tpl_stem:
            overlap += 5

        if overlap > best_score:
            best_score = overlap
            best_match = tpl

    if best_match:
        logger.info(f"Template match: {best_match.name} (score={best_score})")
    else:
        templates = list(templates_dir.glob("*.pptx"))
        if templates:
            best_match = templates[0]
            logger.info(f"No keyword match, using first template: {best_match.name}")

    return best_match
