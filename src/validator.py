from __future__ import annotations
import re
import logging
from .schemas import (
    PresentationSpec, SlideSpec, SlideElement,
    TextContent, BulletContent, ChartContent, TableContent,
    InfographicContent, Position, SlideMasterInfo,
)
from . import config

# Optional profile import for type hints
try:
    from .content_profiler import ContentProfile
except ImportError:
    ContentProfile = None  # type: ignore

logger = logging.getLogger(__name__)

# Default slide dims — overridden dynamically in validate_and_fix
_SLIDE_W = config.SLIDE_WIDTH
_SLIDE_H = config.SLIDE_HEIGHT


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


def validate_and_fix(
    spec: PresentationSpec,
    content_profile=None,
    slide_width: int | None = None,
    slide_height: int | None = None,
    master_info: SlideMasterInfo | None = None,
) -> ValidationResult:
    """Validate a PresentationSpec and apply rule-based auto-fixes. Mutates spec in place."""
    global _SLIDE_W, _SLIDE_H
    if slide_width:
        _SLIDE_W = slide_width
    if slide_height:
        _SLIDE_H = slide_height

    result = ValidationResult()

    _check_slide_count(spec, result)
    _check_slide_flow(spec, result, master_info)
    _enforce_narrative_arc(spec, result)
    _check_visual_ratio(spec, result, content_profile)

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
    if n > config.MAX_SLIDES:
        # Enforce: trim excess content slides using importance-based selection
        cover = [s for s in spec.slides if s.slide_type == "cover"][:1]
        thank = [s for s in spec.slides if s.slide_type == "thank_you"][-1:]
        middle = [s for s in spec.slides if s.slide_type not in ("cover", "thank_you")]
        allowed = config.MAX_SLIDES - len(cover) - len(thank)
        if len(middle) > allowed:
            # Sort by importance_score (descending), keep top N, then re-sort by position
            middle.sort(key=lambda s: s.importance_score, reverse=True)
            kept = middle[:allowed]
            dropped = middle[allowed:]
            # Preserve original ordering
            kept.sort(key=lambda s: s.slide_number)
            dropped_types = [f"{s.slide_type}({s.slide_number})" for s in dropped]
            result.fixes_applied.append(
                f"Trimmed to {config.MAX_SLIDES} slides (was {n}), dropped: {', '.join(dropped_types)}"
            )
            middle = kept
        spec.slides = cover + middle + thank
        # Renumber
        for i, s in enumerate(spec.slides):
            s.slide_number = i + 1
    elif n < config.MIN_SLIDES:
        result.warnings.append(f"Only {n} slides, minimum is {config.MIN_SLIDES}")


def _enforce_narrative_arc(spec: PresentationSpec, result: ValidationResult) -> None:
    """Warn on narrative arc violations — Introduction → Body → Conclusion flow.

    This is advisory (warnings only) to preserve the planner's decisions while
    flagging issues for debugging.
    """
    slides = spec.slides
    if len(slides) < 4:
        return

    # Check: executive_summary should come before body content
    exec_idx = next(
        (i for i, s in enumerate(slides) if s.slide_type == "executive_summary"), -1
    )
    first_content_idx = next(
        (i for i, s in enumerate(slides)
         if s.slide_type in ("content", "chart", "table", "infographic", "mixed")), -1
    )
    if exec_idx > 0 and first_content_idx > 0 and exec_idx > first_content_idx:
        result.warnings.append(
            f"Narrative arc: executive_summary (slide {exec_idx + 1}) "
            f"appears after first content slide (slide {first_content_idx + 1})"
        )

    # Check: conclusion should come after all content
    conclusion_idx = next(
        (i for i, s in enumerate(slides) if s.slide_type == "conclusion"), -1
    )
    if conclusion_idx > 0:
        content_after = [
            i for i, s in enumerate(slides)
            if i > conclusion_idx
            and s.slide_type in ("content", "chart", "table", "infographic", "mixed")
        ]
        if content_after:
            result.warnings.append(
                f"Narrative arc: content slides {content_after} appear after conclusion "
                f"(slide {conclusion_idx + 1})"
            )


def _check_visual_ratio(spec: PresentationSpec, result: ValidationResult, profile=None) -> None:
    """Check that the deck meets the recommended visual ratio from the content profile."""
    content_slides = [
        s for s in spec.slides
        if s.slide_type not in ("cover", "agenda", "section_divider", "thank_you")
    ]
    if not content_slides:
        return

    visual_slides = sum(
        1 for s in content_slides
        if any(el.element_type in ("chart", "table", "infographic") for el in s.elements)
    )
    ratio = visual_slides / len(content_slides)

    target = 0.5  # default: 50% visual
    if profile and hasattr(profile, 'recommended_visual_ratio'):
        target = profile.recommended_visual_ratio

    if ratio < target * 0.6:  # warn if below 60% of target
        result.warnings.append(
            f"Visual ratio {ratio:.0%} is below recommended {target:.0%} "
            f"({visual_slides}/{len(content_slides)} content slides have visuals)"
        )


def _check_slide_flow(spec: PresentationSpec, result: ValidationResult, master_info: SlideMasterInfo | None = None) -> None:
    if not spec.slides:
        result.errors.append("No slides in presentation")
        return

    _enforce_slide_ordering(spec, result, master_info)


def _enforce_slide_ordering(spec: PresentationSpec, result: ValidationResult, master_info: SlideMasterInfo | None = None) -> None:
    """Hard enforcement of structural slide ordering (Brief §3.10).

    Rules applied (mutates ``spec.slides`` in place):
    1. Cover must be slide 1 — move if misplaced, create if missing.
    2. Thank you must be the last slide — move if misplaced, create if missing.
       When a template is used (``master_info`` is not None), thank_you is
       **never** auto-created because the renderer adds the template's closing
       slide as a fixed bookend regardless of its type.
    3. Remove duplicate structural slides (max 1 cover, 1 thank_you, 1 agenda).
    4. Section dividers cannot be adjacent.
    5. Renumber all slides after reordering.
    """
    slides = spec.slides

    # --- Deduplicate structural slides ---
    seen_types: dict[str, int] = {}
    deduped: list[SlideSpec] = []
    for s in slides:
        if s.slide_type in ("cover", "thank_you", "agenda"):
            count = seen_types.get(s.slide_type, 0)
            if count > 0:
                result.fixes_applied.append(
                    f"Removed duplicate '{s.slide_type}' slide (was slide {s.slide_number})"
                )
                continue
            seen_types[s.slide_type] = count + 1
        deduped.append(s)
    slides = deduped

    # --- Remove adjacent section dividers ---
    cleaned: list[SlideSpec] = []
    for s in slides:
        if s.slide_type == "section_divider" and cleaned and cleaned[-1].slide_type == "section_divider":
            result.fixes_applied.append(
                f"Removed adjacent section_divider (slide {s.slide_number})"
            )
            continue
        cleaned.append(s)
    slides = cleaned

    # --- Remove empty content slides (no elements, not structural) ---
    non_empty: list[SlideSpec] = []
    for s in slides:
        if s.slide_type in ("cover", "thank_you", "section_divider", "agenda"):
            non_empty.append(s)
            continue
        if s.elements:
            non_empty.append(s)
        else:
            result.fixes_applied.append(
                f"Removed empty slide {s.slide_number} (type={s.slide_type}, no content)"
            )
    slides = non_empty

    # --- Ensure cover is first ---
    cover_slides = [s for s in slides if s.slide_type == "cover"]
    non_cover = [s for s in slides if s.slide_type != "cover"]
    if cover_slides:
        if slides[0].slide_type != "cover":
            result.fixes_applied.append("Moved cover slide to position 1")
        slides = cover_slides[:1] + non_cover
    else:
        result.warnings.append("No cover slide found in presentation")

    # --- Handle thank_you slide ---
    if master_info:
        # Template provides the closing slide — strip any generated thank_you
        ty_in_plan = [s for s in slides if s.slide_type == "thank_you"]
        if ty_in_plan:
            slides = [s for s in slides if s.slide_type != "thank_you"]
            result.fixes_applied.append(
                "Removed thank_you slide — template provides the closing slide"
            )
    else:
        # No template — ensure thank_you is present and last
        ty_slides = [s for s in slides if s.slide_type == "thank_you"]
        non_ty = [s for s in slides if s.slide_type != "thank_you"]
        if ty_slides:
            if slides[-1].slide_type != "thank_you":
                result.fixes_applied.append("Moved thank_you slide to last position")
            slides = non_ty + ty_slides[-1:]
        else:
            ty = SlideSpec(
                slide_number=len(slides) + 1,
                slide_type="thank_you",
                layout_name=config.LAYOUT_THANK_YOU,
                title="Thank You",
            )
            slides.append(ty)
            result.fixes_applied.append("Appended missing thank_you slide")

    # --- Renumber ---
    for i, s in enumerate(slides):
        s.slide_number = i + 1

    spec.slides = slides


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
    if right_edge > _SLIDE_W:
        pos.width = _SLIDE_W - pos.left - int(config.MARGIN_RIGHT)
        fixed = True

    # Ensure height doesn't overflow bottom edge
    bottom_edge = pos.top + pos.height
    if bottom_edge > _SLIDE_H:
        pos.height = _SLIDE_H - pos.top - int(config.MARGIN_BOTTOM)
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


def _smart_truncate_validator(text: str, max_chars: int) -> str:
    """Truncate text preferring sentence boundaries, no trailing ellipsis."""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    candidate = text[:max_chars]
    last_sentence_end = -1
    for m in re.finditer(r'[.!?](?:\s|$)', candidate):
        pos = m.start() + 1
        if pos <= max_chars:
            last_sentence_end = pos
    if last_sentence_end > max_chars * 0.25:
        return text[:last_sentence_end].strip()
    last_space = candidate.rfind(' ')
    if last_space > max_chars * 0.4:
        candidate = candidate[:last_space]
    return candidate.rstrip('.,;:- ')


def _fix_bullets(slide_num: int, content: BulletContent, result: ValidationResult) -> None:
    # Filter out empty bullets first
    content.items = [b for b in content.items if b and b.strip()]
    if not content.items:
        result.warnings.append(f"Slide {slide_num}: empty bullet list")
        return

    # Enforce max bullets
    if len(content.items) > config.MAX_BULLETS_PER_SLIDE:
        content.items = content.items[:config.MAX_BULLETS_PER_SLIDE]
        result.fixes_applied.append(f"Slide {slide_num}: trimmed bullets to {config.MAX_BULLETS_PER_SLIDE}")

    # Enforce max chars per bullet (sentence-boundary truncation)
    for i, item in enumerate(content.items):
        if len(item) > config.MAX_CHARS_PER_BULLET:
            content.items[i] = _smart_truncate_validator(item, config.MAX_CHARS_PER_BULLET)
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
    if not content.text or not content.text.strip():
        result.warnings.append(f"Slide {slide_num}: empty text element")
    elif len(content.text) > 600:
        content.text = _smart_truncate_validator(content.text, 600)
        result.fixes_applied.append(f"Slide {slide_num}: truncated long text")


def _fix_infographic(slide_num: int, content: InfographicContent, result: ValidationResult) -> None:
    # Filter out items with empty titles
    content.items = [it for it in content.items if it.title and it.title.strip()]
    if not content.items:
        result.warnings.append(f"Slide {slide_num}: infographic has no items")
    elif len(content.items) > 6:
        content.items = content.items[:6]
        result.fixes_applied.append(f"Slide {slide_num}: trimmed infographic to 6 items")
