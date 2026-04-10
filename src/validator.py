from __future__ import annotations
import logging
from .schemas import (
    PresentationSpec, SlideSpec, SlideElement,
    TextContent, BulletContent, ChartContent, TableContent,
    InfographicContent, Position,
)
from . import config

logger = logging.getLogger(__name__)

SLIDE_W = config.SLIDE_WIDTH
SLIDE_H = config.SLIDE_HEIGHT


class ValidationResult:
    def __init__(self):
        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.fixes_applied: list[str] = []

    @property
    def passed(self) -> bool:
        return len(self.errors) == 0

    def __repr__(self) -> str:
        return (
            f"ValidationResult(passed={self.passed}, "
            f"errors={len(self.errors)}, warnings={len(self.warnings)}, "
            f"fixes={len(self.fixes_applied)})"
        )


def validate_and_fix(spec: PresentationSpec) -> ValidationResult:
    """Validate a PresentationSpec and apply rule-based auto-fixes. Mutates spec in place."""
    result = ValidationResult()

    _check_slide_count(spec, result)
    _check_slide_flow(spec, result)

    for slide in spec.slides:
        _check_slide(slide, result)

    if result.errors:
        logger.warning(f"Validation: {len(result.errors)} errors, {len(result.warnings)} warnings")
    else:
        logger.info(f"Validation passed. {len(result.warnings)} warnings, {len(result.fixes_applied)} fixes applied")

    return result


# ── Top-level checks ──

def _check_slide_count(spec: PresentationSpec, result: ValidationResult) -> None:
    n = len(spec.slides)
    if n < config.MIN_SLIDES:
        result.warnings.append(f"Only {n} slides, minimum is {config.MIN_SLIDES}")
    elif n > config.MAX_SLIDES:
        result.warnings.append(f"{n} slides exceeds maximum {config.MAX_SLIDES}")


def _check_slide_flow(spec: PresentationSpec, result: ValidationResult) -> None:
    if not spec.slides:
        result.errors.append("No slides in presentation")
        return

    # First slide should be cover
    if spec.slides[0].slide_type != "cover":
        result.warnings.append(f"First slide is '{spec.slides[0].slide_type}', expected 'cover'")

    # Last slide should be thank_you
    if spec.slides[-1].slide_type != "thank_you":
        result.warnings.append(f"Last slide is '{spec.slides[-1].slide_type}', expected 'thank_you'")


# ── Per-slide checks ──

MAX_ELEMENTS_PER_SLIDE = 4
MAX_TOTAL_TEXT_CHARS = 1200


def _check_slide(slide: SlideSpec, result: ValidationResult) -> None:
    # Check title
    if not slide.title and slide.slide_type not in ("thank_you",):
        result.warnings.append(f"Slide {slide.slide_number}: missing title")

    # Content density: cap element count
    if len(slide.elements) > MAX_ELEMENTS_PER_SLIDE:
        slide.elements = slide.elements[:MAX_ELEMENTS_PER_SLIDE]
        result.fixes_applied.append(
            f"Slide {slide.slide_number}: trimmed to {MAX_ELEMENTS_PER_SLIDE} elements (density)"
        )

    # Content density: estimate total text characters
    total_chars = 0
    for element in slide.elements:
        c = element.content
        if isinstance(c, TextContent):
            total_chars += len(c.text)
        elif isinstance(c, BulletContent):
            total_chars += sum(len(b) for b in c.items)
        elif isinstance(c, InfographicContent):
            total_chars += sum(len(it.title) + len(it.description or "") for it in c.items)

    if total_chars > MAX_TOTAL_TEXT_CHARS:
        # Auto-reduce: trim bullets and text
        for element in slide.elements:
            c = element.content
            if isinstance(c, BulletContent) and len(c.items) > 4:
                c.items = c.items[:4]
                result.fixes_applied.append(
                    f"Slide {slide.slide_number}: reduced bullets for density"
                )
            elif isinstance(c, TextContent) and len(c.text) > 300:
                c.text = c.text[:297] + "..."
                result.fixes_applied.append(
                    f"Slide {slide.slide_number}: trimmed text for density"
                )

    has_major_visual = any(_is_visual_element(element) for element in slide.elements)
    if has_major_visual and total_chars > 700:
        for element in slide.elements:
            c = element.content
            if isinstance(c, BulletContent) and len(c.items) > 4:
                c.items = c.items[:4]
                result.fixes_applied.append(
                    f"Slide {slide.slide_number}: reduced supporting bullets around major visual"
                )
            elif isinstance(c, TextContent) and len(c.text) > 220:
                c.text = c.text[:217].rstrip() + "..."
                result.fixes_applied.append(
                    f"Slide {slide.slide_number}: reduced supporting text around major visual"
                )

    _prune_visual_clutter(slide, result)

    _resolve_overlaps(slide, result)

    for element in slide.elements:
        _check_element_bounds(slide.slide_number, element, result)
        _check_element_content(slide.slide_number, element, result)


def _is_visual_element(element: SlideElement) -> bool:
    return element.element_type in ("chart", "table", "infographic")


def _prune_visual_clutter(slide: SlideSpec, result: ValidationResult) -> None:
    """Keep slide compositions intentionally simple and readable."""
    if slide.slide_type == "mixed" and len(slide.elements) > 2:
        slide.elements = slide.elements[:2]
        result.fixes_applied.append(
            f"Slide {slide.slide_number}: trimmed mixed slide to two elements"
        )
        return

    visuals_seen = 0
    pruned: list[SlideElement] = []
    for element in slide.elements:
        if _is_visual_element(element):
            visuals_seen += 1
            if visuals_seen > 1:
                result.fixes_applied.append(
                    f"Slide {slide.slide_number}: removed extra visual element '{element.element_type}'"
                )
                continue
        pruned.append(element)
    slide.elements = pruned


def _resolve_overlaps(slide: SlideSpec, result: ValidationResult) -> None:
    """Drop materially overlapping elements, preserving the higher-priority one."""
    resolved: list[SlideElement] = []
    for element in slide.elements:
        conflicting_idx = next(
            (
                idx for idx, kept in enumerate(resolved)
                if _overlap_ratio(element.position, kept.position) > 0.14
            ),
            None,
        )
        if conflicting_idx is None:
            resolved.append(element)
            continue

        kept = resolved[conflicting_idx]
        if _element_priority(element) > _element_priority(kept):
            resolved[conflicting_idx] = element
            result.fixes_applied.append(
                f"Slide {slide.slide_number}: replaced overlapping '{kept.element_type}' with '{element.element_type}'"
            )
        else:
            result.fixes_applied.append(
                f"Slide {slide.slide_number}: removed overlapping '{element.element_type}'"
            )
    slide.elements = resolved


def _overlap_ratio(a: Position, b: Position) -> float:
    overlap_left = max(a.left, b.left)
    overlap_top = max(a.top, b.top)
    overlap_right = min(a.left + a.width, b.left + b.width)
    overlap_bottom = min(a.top + a.height, b.top + b.height)

    if overlap_right <= overlap_left or overlap_bottom <= overlap_top:
        return 0.0

    overlap_area = (overlap_right - overlap_left) * (overlap_bottom - overlap_top)
    smaller_area = min(a.width * a.height, b.width * b.height)
    return overlap_area / smaller_area if smaller_area else 0.0


def _element_priority(element: SlideElement) -> int:
    if element.element_type in ("chart", "table"):
        return 3
    if element.element_type == "infographic":
        return 2
    if element.element_type in ("bullets", "text"):
        return 1
    return 0


def _check_element_bounds(slide_num: int, element: SlideElement, result: ValidationResult) -> None:
    """Check that element is within slide boundaries and fix if needed."""
    pos = element.position

    fixed = False

    # Clamp left
    if pos.left < 0:
        pos.left = int(config.MARGIN_LEFT)
        fixed = True

    # Clamp top
    if pos.top < 0:
        pos.top = int(config.MARGIN_TOP)
        fixed = True

    # Ensure width doesn't overflow right edge
    right_edge = pos.left + pos.width
    if right_edge > SLIDE_W:
        pos.width = SLIDE_W - pos.left - int(config.MARGIN_RIGHT)
        fixed = True

    # Ensure height doesn't overflow bottom edge
    bottom_edge = pos.top + pos.height
    if bottom_edge > SLIDE_H:
        pos.height = SLIDE_H - pos.top - int(config.MARGIN_BOTTOM)
        fixed = True

    # Ensure minimum dimensions
    if pos.width < 100000:
        pos.width = 100000
        fixed = True
    if pos.height < 100000:
        pos.height = 100000
        fixed = True

    if fixed:
        result.fixes_applied.append(f"Slide {slide_num}: fixed element bounds for {element.element_type}")


def _check_element_content(slide_num: int, element: SlideElement, result: ValidationResult) -> None:
    """Check and fix element content issues."""
    content = element.content

    if isinstance(content, BulletContent):
        _fix_bullets(slide_num, content, result)
    elif isinstance(content, ChartContent):
        _fix_chart(slide_num, content, result)
    elif isinstance(content, TableContent):
        _fix_table(slide_num, content, result)
    elif isinstance(content, TextContent):
        _fix_text(slide_num, content, result)
    elif isinstance(content, InfographicContent):
        _fix_infographic(slide_num, content, result)


def _fix_bullets(slide_num: int, content: BulletContent, result: ValidationResult) -> None:
    if not content.items:
        result.warnings.append(f"Slide {slide_num}: empty bullet list")
        return

    # Enforce max bullets
    if len(content.items) > config.MAX_BULLETS_PER_SLIDE:
        content.items = content.items[:config.MAX_BULLETS_PER_SLIDE]
        result.fixes_applied.append(f"Slide {slide_num}: trimmed bullets to {config.MAX_BULLETS_PER_SLIDE}")

    # Enforce max chars per bullet (smart word-boundary truncation)
    for i, item in enumerate(content.items):
        if len(item) > config.MAX_CHARS_PER_BULLET:
            truncated = item[:config.MAX_CHARS_PER_BULLET]
            last_space = truncated.rfind(' ')
            if last_space > config.MAX_CHARS_PER_BULLET * 0.6:
                truncated = truncated[:last_space]
            content.items[i] = truncated.rstrip('.,;:- ') + '…'
            result.fixes_applied.append(f"Slide {slide_num}: truncated bullet {i + 1}")


def _fix_chart(slide_num: int, content: ChartContent, result: ValidationResult) -> None:
    if not content.categories:
        result.warnings.append(f"Slide {slide_num}: chart has no categories (will skip)")
        return
    if not content.series:
        result.warnings.append(f"Slide {slide_num}: chart has no series data (will skip)")
        return
    if content.series:
        # Ensure all series have same length as categories
        n_cats = len(content.categories)
        for s in content.series:
            if len(s.values) < n_cats:
                s.values.extend([0.0] * (n_cats - len(s.values)))
                result.fixes_applied.append(f"Slide {slide_num}: padded series '{s.name}' values")
            elif len(s.values) > n_cats:
                s.values = s.values[:n_cats]
                result.fixes_applied.append(f"Slide {slide_num}: trimmed series '{s.name}' values")


def _fix_table(slide_num: int, content: TableContent, result: ValidationResult) -> None:
    if not content.headers:
        result.warnings.append(f"Slide {slide_num}: table has no headers (will skip)")
        return
    if not content.rows:
        result.warnings.append(f"Slide {slide_num}: table has no data rows")
        return

    n_cols = len(content.headers)
    for i, row in enumerate(content.rows):
        if len(row) < n_cols:
            content.rows[i] = row + [""] * (n_cols - len(row))
            result.fixes_applied.append(f"Slide {slide_num}: padded table row {i + 1}")
        elif len(row) > n_cols:
            content.rows[i] = row[:n_cols]
            result.fixes_applied.append(f"Slide {slide_num}: trimmed table row {i + 1}")


def _fix_text(slide_num: int, content: TextContent, result: ValidationResult) -> None:
    if not content.text:
        result.warnings.append(f"Slide {slide_num}: empty text element")
    elif len(content.text) > 600:
        truncated = content.text[:600]
        last_space = truncated.rfind(' ')
        if last_space > 360:
            truncated = truncated[:last_space]
        content.text = truncated.rstrip('.,;:- ') + '…'
        result.fixes_applied.append(f"Slide {slide_num}: truncated long text")


def _fix_infographic(slide_num: int, content: InfographicContent, result: ValidationResult) -> None:
    if not content.items:
        result.warnings.append(f"Slide {slide_num}: infographic has no items")
    elif len(content.items) > 6:
        content.items = content.items[:6]
        result.fixes_applied.append(f"Slide {slide_num}: trimmed infographic to 6 items")
