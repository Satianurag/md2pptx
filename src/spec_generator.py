from __future__ import annotations
import logging
import math
import re
from typing import Optional
from .schemas import (
    ContentTree, ContentSection, SlidePlan, SlidePlanItem,
    SlideMasterInfo, PresentationSpec, SlideSpec, SlideElement,
    Position, TextContent, BulletContent, ChartContent, ChartSeries,
    TableContent, ShapeContent, InfographicContent, InfographicItem,
    DataTable, KeyMetric, DeckContent, SlideContent,
)
from .content_profiler import ContentProfile
from .grid_system import Grid
from . import config

logger = logging.getLogger(__name__)


# Code detection patterns for JetBrains Mono font application
_CODE_PATTERNS = [
    re.compile(r'^\s*[\w\-]+:\s*\S+', re.MULTILINE),  # YAML-like key: value
    re.compile(r'[{\[\]}]'),  # JSON brackets
    re.compile(r'(function|class|def|const|let|var)\s+'),  # Code keywords
    re.compile(r'[;{}]\s*$', re.MULTILINE),  # Statement endings
    re.compile(r'\`[^`]+\`'),  # Inline code backticks
]


def _is_code_content(text: str) -> bool:
    """Detect if text appears to be code/technical content for monospace font."""
    if not text:
        return False
    code_indicators = sum(1 for p in _CODE_PATTERNS if p.search(text))
    return code_indicators >= 2 or text.count('`') >= 2


def _create_text_content(
    text: str,
    font_size: int | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    alignment: str | None = None,
) -> TextContent:
    """Create TextContent with appropriate font based on content type."""
    is_code = _is_code_content(text)
    return TextContent(
        text=text,
        font_size=font_size,
        font_name=config.FONT_NAME_MONO if is_code else config.FONT_NAME_PRIMARY,
        bold=bold,
        italic=italic,
        color=color,
        alignment=alignment,
        is_code=is_code,
    )


def _smart_truncate(text: str, max_chars: int = 200) -> str:
    """Truncate text intelligently, preferring sentence boundaries over mid-word cuts.

    Strategy:
    1. If text fits, return as-is.
    2. Try to end at the last sentence boundary (. ! ?) within the limit.
    3. Fall back to word-boundary cut only when no sentence boundary exists.
    """
    text = text.strip()
    if len(text) <= max_chars:
        return text

    # Try sentence-boundary truncation: find last sentence-ending punctuation
    # that occurs before max_chars
    candidate = text[:max_chars]
    last_sentence_end = -1
    for m in re.finditer(r'[.!?](?:\s|$)', candidate):
        pos = m.start() + 1  # include the punctuation
        if pos <= max_chars:
            last_sentence_end = pos

    # Use sentence boundary if it preserves at least 25% of allowed length
    if last_sentence_end > max_chars * 0.25:
        return text[:last_sentence_end].strip()

    # Fall back to word-boundary cut — never cut mid-word
    last_space = candidate.rfind(' ')
    if last_space > max_chars * 0.4:
        candidate = candidate[:last_space]
    return candidate.rstrip('.,;:- ')


# Module-level grid instance — set at the start of generate_presentation_spec()
_grid: Grid = Grid.default()


def generate_presentation_spec(
    content_tree: ContentTree,
    slide_plan: SlidePlan,
    master_info: SlideMasterInfo | None = None,
    template_path: str = "",
    content_profile: ContentProfile | None = None,
    deck_content: DeckContent | None = None,
) -> PresentationSpec:
    """Convert a SlidePlan + ContentTree + AI-written DeckContent into a full PresentationSpec."""
    global _grid
    _grid = Grid.from_template(master_info) if master_info else Grid.default()
    logger.info(f"Grid: {_grid.slide_w}x{_grid.slide_h}, content_w={_grid.content_w}, content_h={_grid.content_h}")

    # Build slide_number → SlideContent lookup
    content_map: dict[int, SlideContent] = {}
    if deck_content:
        for sc in deck_content.slides:
            content_map[sc.slide_number] = sc

    slides: list[SlideSpec] = []

    for plan_item in slide_plan.slides:
        slide_content = content_map.get(plan_item.slide_number)
        slide_spec = _generate_slide_spec(content_tree, plan_item, slide_plan, slide_content)
        slide_spec.importance_score = plan_item.importance_score
        slides.append(slide_spec)

    _ensure_visual_coverage(slides, content_tree, content_profile, slide_plan.target_slide_count)

    return PresentationSpec(
        title=content_tree.title,
        subtitle=content_tree.subtitle,
        slides=slides,
        template_path=template_path,
        target_slide_count=len(slides),
    )


def _ensure_visual_coverage(
    slides: list[SlideSpec],
    tree: ContentTree,
    profile: ContentProfile | None = None,
    target_count: int = 15,
) -> None:
    """Ensure the deck contains dedicated visual slides without cluttering existing ones.

    When a *profile* is available, use its ranked tables/metrics for smarter selection.
    Respects *target_count* so we don't bloat the deck beyond what the user asked for.
    """
    # Build ordered table list (profile-ranked first, then fallback)
    ordered_tables: list[DataTable] = []
    if profile and profile.best_tables:
        ordered_tables = [st.table for st in profile.best_tables]
    if not ordered_tables:
        ordered_tables = list(tree.all_tables)

    if not ordered_tables:
        _renumber_slides(slides)
        return

    if not _deck_has_element(slides, "chart"):
        for table in ordered_tables:
            if not (table.rows and table.headers and len(table.headers) >= 2):
                continue
            categories, series = _extract_chart_data(table)
            if categories and series:
                if _insert_support_slide(
                    slides,
                    _build_support_chart_slide(table, categories, series),
                    target_count,
                ):
                    logger.info("Added dedicated chart slide for deck-level visual coverage")
                break

    if not _deck_has_element(slides, "table"):
        best_table = None
        for table in ordered_tables:
            if table.rows and table.headers and len(table.headers) >= 2:
                if _should_render_as_table(table) or best_table is None:
                    best_table = table
                    if _should_render_as_table(table):
                        break
        if best_table and _insert_support_slide(slides, _build_support_table_slide(best_table), target_count):
            logger.info("Added dedicated table slide for deck-level visual coverage")

    _renumber_slides(slides)


def _deck_has_element(slides: list[SlideSpec], element_type: str) -> bool:
    return any(el.element_type == element_type for slide in slides for el in slide.elements)


def _insert_support_slide(slides: list[SlideSpec], support_slide: SlideSpec, target_count: int = 15) -> bool:
    """Insert a new support slide before closing slides, or replace a low-value content slide."""
    max_allowed = min(target_count, config.MAX_SLIDES)
    insert_at = next(
        (idx for idx, slide in enumerate(slides) if slide.slide_type in ("conclusion", "thank_you")),
        len(slides),
    )
    if len(slides) < max_allowed:
        slides.insert(insert_at, support_slide)
        return True

    for idx, slide in enumerate(slides):
        has_visual = any(el.element_type in ("chart", "table", "infographic") for el in slide.elements)
        if slide.slide_type == "content" and slide.elements and not has_visual:
            support_slide.slide_number = slide.slide_number
            slides[idx] = support_slide
            return True
    return False


def _renumber_slides(slides: list[SlideSpec]) -> None:
    for idx, slide in enumerate(slides, start=1):
        slide.slide_number = idx


def _build_support_chart_slide(
    table: DataTable,
    categories: list[str],
    series: list[ChartSeries],
) -> SlideSpec:
    title = _smart_truncate(table.title or "Data Highlights", 60)
    return SlideSpec(
        slide_number=0,
        slide_type="chart",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        elements=[
            SlideElement(
                element_type="chart",
                position=_grid.chart(),
                content=ChartContent(
                    chart_type=_auto_detect_chart_type(table),
                    title=table.title or title,
                    categories=categories,
                    series=series,
                ),
            )
        ],
    )


def _build_support_table_slide(table: DataTable) -> SlideSpec:
    title = _smart_truncate(table.title or "Detailed Data Table", 60)
    headers = table.headers[:6]
    rows = [row[:6] for row in table.rows[:8]]
    return SlideSpec(
        slide_number=0,
        slide_type="table",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        elements=[
            SlideElement(
                element_type="table",
                position=_grid.table(),
                content=TableContent(headers=headers, rows=rows),
            )
        ],
    )


def _resolve_title(plan_item: SlidePlanItem) -> str:
    """Pick the best title: action_title > title."""
    return plan_item.action_title or plan_item.title


def _generate_slide_spec(
    content_tree: ContentTree,
    plan_item: SlidePlanItem,
    slide_plan: SlidePlan | None = None,
    slide_content: SlideContent | None = None,
) -> SlideSpec:
    """Generate a single SlideSpec from a SlidePlanItem, using AI-written SlideContent."""

    # Use AI-written title when available
    title = (slide_content.title if slide_content and slide_content.title
             else plan_item.action_title or plan_item.title)

    if plan_item.slide_type == "cover":
        return _build_cover_slide(content_tree, plan_item, slide_content)

    if plan_item.slide_type == "thank_you":
        return _build_thank_you_slide(plan_item, slide_content)

    if plan_item.slide_type == "section_divider":
        return _build_divider_slide(plan_item, slide_content)

    # For content slides, gather relevant source content
    source_sections = _find_source_sections(content_tree, plan_item.content_source)

    # Route to the appropriate builder based on visualization_hint
    if plan_item.slide_type == "agenda":
        return _build_agenda_slide(plan_item, content_tree, slide_content)

    if plan_item.slide_type == "executive_summary":
        return _build_exec_summary_slide(plan_item, content_tree, slide_content)

    if plan_item.slide_type == "conclusion":
        return _build_conclusion_slide(plan_item, source_sections, content_tree, slide_plan, slide_content)

    if plan_item.visualization_hint == "chart" and source_sections:
        return _build_chart_slide(plan_item, source_sections, slide_content)

    if plan_item.visualization_hint == "table" and source_sections:
        return _build_table_slide(plan_item, source_sections, slide_content)

    if plan_item.visualization_hint == "kpi" and source_sections:
        return _build_kpi_slide(plan_item, source_sections, content_tree, slide_content)

    if plan_item.visualization_hint == "infographic" and source_sections:
        return _build_infographic_slide(plan_item, source_sections, slide_content)

    if plan_item.visualization_hint == "mixed" and source_sections:
        return _build_mixed_slide(plan_item, source_sections, content_tree, slide_content)

    # Default: bullets slide
    return _build_bullets_slide(plan_item, source_sections, slide_content)


def _build_cover_slide(tree: ContentTree, plan: SlidePlanItem, sc: SlideContent | None = None) -> SlideSpec:
    title = sc.title if sc and sc.title else (tree.title or plan.title)
    subtitle = sc.subtitle if sc and sc.subtitle else (tree.subtitle or plan.subtitle or "")
    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="cover",
        layout_name="cover",
        title=title,
        subtitle=subtitle,
        elements=[],
    )


def _build_thank_you_slide(plan: SlidePlanItem, sc: SlideContent | None = None) -> SlideSpec:
    title = sc.title if sc and sc.title else (plan.title or "Thank You")
    subtitle = sc.subtitle if sc and sc.subtitle else (plan.subtitle or "Questions & Discussion")
    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="thank_you",
        layout_name="thank_you",
        title=title,
        subtitle=subtitle,
        elements=[],
    )


def _build_divider_slide(plan: SlidePlanItem, sc: SlideContent | None = None) -> SlideSpec:
    key_msg = sc.key_takeaway if sc and sc.key_takeaway else (plan.key_message or "")
    subtitle = sc.subtitle if sc and sc.subtitle else (plan.subtitle or key_msg)
    elements = []

    if key_msg.strip():
        elements.append(SlideElement(
            element_type="text",
            position=_grid.full(),
            content=_create_text_content(
                text=key_msg,
                font_size=14,
                italic=True,
            ),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="section_divider",
        layout_name="divider",
        title=sc.title if sc and sc.title else plan.title,
        subtitle=subtitle,
        elements=elements,
    )


def _find_source_sections(
    tree: ContentTree,
    content_source: list[str],
) -> list[ContentSection]:
    """Find sections matching the content_source headings.

    Uses exact match first, then substring/partial match as fallback
    to avoid empty slides when LLM-generated headings don't match exactly.
    """
    if not content_source:
        return []

    found = []
    source_lower = [s.lower().strip() for s in content_source]
    found_headings: set[str] = set()

    def search_exact(sections: list[ContentSection]) -> None:
        for sec in sections:
            h = sec.heading.lower().strip()
            if h in source_lower:
                found.append(sec)
                found_headings.add(h)
            search_exact(sec.subsections)

    search_exact(tree.sections)

    # Fallback: partial / substring matching for unmatched sources
    unmatched = [s for s in source_lower if s not in found_headings]
    if unmatched:
        def search_partial(sections: list[ContentSection]) -> None:
            for sec in sections:
                h = sec.heading.lower().strip()
                if h in found_headings:
                    continue
                for src in unmatched:
                    if src in h or h in src:
                        found.append(sec)
                        found_headings.add(h)
                        break
                    # Word overlap: if ≥50% of words match
                    src_words = set(src.split())
                    h_words = set(h.split())
                    if src_words and h_words:
                        overlap = len(src_words & h_words) / min(len(src_words), len(h_words))
                        if overlap >= 0.5:
                            found.append(sec)
                            found_headings.add(h)
                            break
                search_partial(sec.subsections)

        search_partial(tree.sections)

    return found


def _parse_number(s: str) -> float:
    """Parse a string into a float, handling currency/percent/comma formats.

    Returns ``math.nan`` for non-numeric values like 'N/A', empty strings,
    or purely textual content — so callers can distinguish missing data
    from a genuine zero.
    """
    s = str(s).strip()
    if not s or s.upper() in ('N/A', 'NA', '-', '—', 'N.A.', 'N.A', 'NULL', 'NONE', 'TBD', 'TBC'):
        return math.nan
    s = re.sub(r'[,$€£%]', '', s)
    s = s.replace(' ', '')
    try:
        # Handle B/M/K suffixes
        if s.upper().endswith('B'):
            return float(s[:-1]) * 1_000_000_000
        elif s.upper().endswith('M'):
            return float(s[:-1]) * 1_000_000
        elif s.upper().endswith('K'):
            return float(s[:-1]) * 1_000
        elif s.upper().endswith('T'):
            return float(s[:-1]) * 1_000_000_000_000
        return float(s)
    except (ValueError, IndexError):
        return math.nan


def _is_numeric_column(table: DataTable, col_idx: int) -> bool:
    """Return True if > 40% of non-empty cells in *col_idx* are numeric."""
    total = 0
    numeric = 0
    for row in table.rows:
        if col_idx < len(row):
            cell = str(row[col_idx]).strip()
            if not cell:
                continue
            total += 1
            val = _parse_number(cell)
            if not math.isnan(val):
                numeric += 1
    if total == 0:
        return False
    return numeric / total > 0.40


def _should_render_as_table(table: DataTable) -> bool:
    """Return True when a DataTable is better shown as a table than a chart.

    Heuristic: render as table when the data is mostly textual, has many
    columns, or has very few rows where a chart adds no value.
    """
    if not table.headers or not table.rows:
        return False

    n_cols = len(table.headers)
    n_rows = len(table.rows)

    # Count how many data columns are predominantly numeric
    numeric_cols = 0
    for col_idx in range(1, n_cols):  # skip first (category) column
        if _is_numeric_column(table, col_idx):
            numeric_cols += 1

    text_cols = max(n_cols - 1 - numeric_cols, 0)

    # Mostly text columns → table
    if text_cols > numeric_cols:
        return True

    # Many columns (≥5) → chart would be unreadable
    if n_cols >= 5 and numeric_cols <= 2:
        return True

    # Very few rows (≤2) → chart adds no value
    if n_rows <= 2:
        return True

    return False


## _try_rule_based_generation and _detect_infographic_pattern removed — AI content writer handles all content decisions


def _build_agenda_slide(plan: SlidePlanItem, tree: ContentTree, sc: SlideContent | None = None) -> SlideSpec:
    """Build agenda slide using AI-written content."""
    # AI-written bullets preferred
    if sc and sc.bullets:
        items = sc.bullets
    else:
        items = []
        for sec in tree.sections[:10]:
            heading = sec.heading.strip()
            if heading.lower() not in ("executive summary", "table of contents", "references", "references and source documentation"):
                items.append(heading)

    elements = []
    if items:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.full(),
            content=BulletContent(items=items[:8], font_size=14),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="agenda",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=sc.title if sc and sc.title else (plan.title or "Agenda"),
        elements=elements,
    )


def _build_exec_summary_slide(plan: SlidePlanItem, tree: ContentTree, sc: SlideContent | None = None) -> SlideSpec:
    """Build executive summary slide using AI-written content."""
    elements = []

    # AI-written bullets preferred
    if sc and sc.bullets:
        key_points = sc.bullets
    else:
        summary = tree.executive_summary or ""
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', summary) if s.strip()]
        key_points = sentences[:5]

    if key_points:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.full(),
            content=BulletContent(items=key_points[:6], font_size=13),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="executive_summary",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=sc.title if sc and sc.title else (plan.title or "Executive Summary"),
        elements=elements,
    )


def _build_bullets_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build a bullet list slide using AI-written content."""
    # AI-written bullets preferred
    if sc and sc.bullets:
        bullets = sc.bullets
    else:
        bullets = []
        for sec in sections:
            if sec.bullets:
                bullets.extend(sec.bullets)
            elif sec.text:
                sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
                bullets.extend(sentences)
            for sub in sec.subsections:
                if sub.bullets:
                    bullets.extend(sub.bullets)

    bullets = bullets[:config.MAX_BULLETS_PER_SLIDE]

    elements = []
    title = sc.title if sc and sc.title else (plan.action_title or plan.title)
    key_msg = sc.key_takeaway if sc and sc.key_takeaway else (plan.key_message.strip() if plan.key_message else "")

    if key_msg and bullets and len(bullets) >= 3:
        pos_side, pos_main = _grid.sidebar_main()
        elements.append(SlideElement(
            element_type="text",
            position=pos_side,
            content=_create_text_content(
                text=key_msg,
                font_size=13,
                italic=True,
            ),
        ))
        elements.append(SlideElement(
            element_type="bullets",
            position=pos_main,
            content=BulletContent(items=bullets, font_size=14),
        ))
    elif bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.full(),
            content=BulletContent(items=bullets, font_size=14),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type=plan.slide_type,
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )


def _auto_detect_chart_type(table: DataTable) -> str:
    """Automatically select the best chart type based on data structure."""
    n_rows = len(table.rows)
    n_cols = len(table.headers) - 1  # exclude category column

    if n_cols <= 0:
        return "column"

    # Single series with few categories → pie/doughnut
    if n_cols == 1 and 2 <= n_rows <= 8:
        values = [_parse_number(row[1] if len(row) > 1 else "0") for row in table.rows]
        all_positive = all(v >= 0 for v in values)
        if all_positive and sum(values) > 0:
            return "pie" if n_rows <= 5 else "doughnut"

    # Time-series data (years, quarters) → line or area
    first_col_values = [str(row[0]) if row else "" for row in table.rows]
    year_pattern = sum(1 for v in first_col_values if re.match(r'^(19|20)\d{2}', v))
    quarter_pattern = sum(1 for v in first_col_values if re.match(r'^Q[1-4]', v, re.I))
    if year_pattern >= 3 or quarter_pattern >= 3:
        return "line" if n_cols <= 3 else "area"

    # Many categories → bar (horizontal better for long labels)
    if n_rows > 6:
        return "bar"

    # Default → column
    return "column"


def _extract_chart_series(table: DataTable, chart_type: str) -> list[ChartSeries]:
    """Extract chart series from table data, adapted for the chart type."""
    categories, series_list = _extract_chart_data(table)
    return series_list


def _build_chart_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build a chart slide from table data, using AI-written insight as chart title."""
    table = None
    for sec in sections:
        for t in sec.tables:
            if t.rows and t.headers and len(t.headers) >= 2:
                table = t
                break
        if table:
            break
        for sub in sec.subsections:
            for t in sub.tables:
                if t.rows and t.headers and len(t.headers) >= 2:
                    table = t
                    break
            if table:
                break

    if not table:
        # No table data — build as bullets slide instead
        return _build_bullets_slide(plan, sections, sc)

    chart_type = plan.chart_type_hint or _auto_detect_chart_type(table)
    categories, series_list = _extract_chart_data(table)

    if not categories or not series_list:
        return _build_bullets_slide(plan, sections, sc)

    # Use AI-written chart insight as the chart title
    chart_title = sc.chart_insight if sc and sc.chart_insight else (table.title or plan.title)
    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    elements = [
        SlideElement(
            element_type="chart",
            position=_grid.chart(),
            content=ChartContent(
                chart_type=chart_type,
                title=chart_title,
                categories=categories,
                series=series_list,
            ),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="chart",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )


def _detect_column_unit(table: DataTable, col_idx: int) -> str:
    """Detect the dominant unit type of a numeric column: 'pct', 'currency', or 'raw'."""
    pct_count = 0
    cur_count = 0
    total = 0
    for row in table.rows:
        if col_idx < len(row):
            cell = str(row[col_idx]).strip()
            if not cell:
                continue
            total += 1
            if '%' in cell:
                pct_count += 1
            elif any(c in cell for c in ('$', '€', '£')):
                cur_count += 1
    if total == 0:
        return "raw"
    if pct_count / total > 0.4:
        return "pct"
    if cur_count / total > 0.4:
        return "currency"
    return "raw"


def _extract_chart_data(table: DataTable) -> tuple[list[str], list[ChartSeries]]:
    """Extract chart-friendly data from a DataTable.

    Improvements over the naive 'first-col = categories, rest = series':
    1. Skips text-heavy columns that aren't chartable (e.g. 'Bank', 'Year').
    2. Handles N/A / missing values — rows where ALL numeric cells are NaN
       are dropped; remaining NaN values become 0.0.
    3. Separates columns by unit type (%, $, raw) and only charts the
       dominant unit group to prevent misleading mixed-axis visuals.
    """
    if len(table.headers) < 2 or not table.rows:
        return [], []

    # Identify which columns (index ≥ 1) are genuinely numeric
    numeric_col_indices: list[int] = []
    for col_idx in range(1, len(table.headers)):
        if _is_numeric_column(table, col_idx):
            numeric_col_indices.append(col_idx)

    if not numeric_col_indices:
        return [], []

    # Group columns by unit type and pick the largest group
    unit_groups: dict[str, list[int]] = {}
    for ci in numeric_col_indices:
        unit = _detect_column_unit(table, ci)
        unit_groups.setdefault(unit, []).append(ci)

    # Select the dominant unit group (most columns); break ties by preferring raw > currency > pct
    best_unit = max(unit_groups, key=lambda u: (len(unit_groups[u]), {"raw": 2, "currency": 1, "pct": 0}.get(u, 0)))
    selected_cols = unit_groups[best_unit]

    # First column = categories; collect values only from selected columns
    raw_categories: list[str] = []
    raw_values: dict[int, list[float]] = {ci: [] for ci in selected_cols}

    for row in table.rows[:20]:  # cap rows
        if not row:
            continue
        raw_categories.append(str(row[0])[:30])
        for ci in selected_cols:
            val = _parse_number(row[ci] if ci < len(row) else "")
            raw_values[ci].append(val)

    # Filter out rows where ALL numeric values are NaN (pure missing data)
    keep_mask = []
    for row_idx in range(len(raw_categories)):
        has_value = any(
            not math.isnan(raw_values[ci][row_idx]) for ci in selected_cols
        )
        keep_mask.append(has_value)

    categories = [c for c, keep in zip(raw_categories, keep_mask) if keep]
    if not categories:
        return [], []

    series_list = []
    for ci in selected_cols:
        name = table.headers[ci][:30]
        values = [
            (0.0 if math.isnan(v) else v)
            for v, keep in zip(raw_values[ci], keep_mask)
            if keep
        ]
        if any(v != 0 for v in values):  # skip all-zero series
            series_list.append(ChartSeries(name=name, values=values))

    return categories, series_list


def _build_table_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build a table slide using AI-written table summary."""
    table = None
    for sec in sections:
        for t in sec.tables:
            if t.headers and t.rows:
                table = t
                break
        if table:
            break

    if not table:
        return _build_bullets_slide(plan, sections, sc)

    headers = table.headers[:6]
    rows = [row[:6] for row in table.rows[:8]]
    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    elements = [
        SlideElement(
            element_type="table",
            position=_grid.table(),
            content=TableContent(headers=headers, rows=rows),
        )
    ]

    # Add table summary as subtitle if AI provided one
    subtitle = sc.table_summary if sc and sc.table_summary else plan.subtitle

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="table",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=subtitle,
        elements=elements,
    )


def _build_kpi_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
    sc: SlideContent | None = None,
) -> SlideSpec:
    """Build a KPI cards slide using AI-written infographic items or raw metrics."""
    # AI-written infographic items preferred
    if sc and sc.infographic_items:
        items = sc.infographic_items[:6]
    else:
        metrics: list[KeyMetric] = []
        for sec in sections:
            metrics.extend(sec.metrics)
            for sub in sec.subsections:
                metrics.extend(sub.metrics)

        if not metrics:
            metrics = tree.all_metrics[:5]

        if not metrics:
            return _build_bullets_slide(plan, sections, sc)

        items = []
        for m in metrics[:5]:
            items.append(InfographicItem(
                title=m.label,
                description="",
                value=m.value,
            ))

    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    elements = [
        SlideElement(
            element_type="infographic",
            position=_grid.full(),
            content=InfographicContent(
                infographic_type="kpi_cards",
                items=items,
            ),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="infographic",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )


_YEAR_RE = re.compile(r'\b((?:19|20)\d{2})\b')


def _extract_year_from_text(text: str) -> str | None:
    """Extract a 4-digit year from text for timeline labels."""
    m = _YEAR_RE.search(text)
    return m.group(1) if m else None


def _build_infographic_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build an infographic slide using AI-written items."""
    infographic_type = plan.infographic_type_hint or "process_flow"

    # AI-written infographic items preferred
    if sc and sc.infographic_items:
        items = sc.infographic_items[:6]
    else:
        is_timeline = infographic_type == "timeline"
        items = []
        for sec in sections:
            if sec.subsections:
                for sub in sec.subsections[:6]:
                    desc = sub.text[:150] if sub.text else (sub.bullets[0][:150] if sub.bullets else "")
                    title = sub.heading[:80]
                    value = sub.metrics[0].value if sub.metrics else None
                    if is_timeline and not value:
                        value = _extract_year_from_text(sub.heading) or _extract_year_from_text(sub.text or "")
                    if title.strip():
                        items.append(InfographicItem(title=title, description=desc, value=value))
            elif sec.bullets:
                for b in sec.bullets[:6]:
                    parts = b.split(":", 1) if ":" in b else (b.split("–", 1) if "–" in b else [b])
                    title = parts[0].strip()[:80]
                    desc = parts[1].strip()[:150] if len(parts) > 1 else ""
                    value = _extract_year_from_text(b) if is_timeline else None
                    if title.strip():
                        items.append(InfographicItem(title=title, description=desc, value=value))

    if not items:
        return _build_bullets_slide(plan, sections, sc)

    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    elements = [
        SlideElement(
            element_type="infographic",
            position=_grid.full(),
            content=InfographicContent(
                infographic_type=infographic_type,
                items=items[:6],
            ),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="infographic",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )


def _build_mixed_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
    sc: SlideContent | None = None,
) -> SlideSpec:
    """Build a mixed slide: visual element (left) + AI-written bullets (right)."""
    elements = []

    # AI-written bullets preferred
    if sc and sc.bullets:
        bullets = sc.bullets[:config.MAX_BULLETS_PER_SLIDE]
    else:
        bullets = []
        for sec in sections:
            bullets.extend(sec.bullets[:4])
            for sub in sec.subsections:
                bullets.extend(sub.bullets[:3])
        bullets = bullets[:config.MAX_BULLETS_PER_SLIDE]

    # Try to build a chart for left side
    chart_elem = None
    for sec in sections:
        for t in sec.tables:
            if t.headers and t.rows and len(t.headers) >= 2:
                chart_type = plan.chart_type_hint or _auto_detect_chart_type(t)
                series_list = _extract_chart_series(t, chart_type)
                if series_list:
                    categories = [str(row[0]) for row in t.rows[:8]] if t.rows else []
                    pos_left, pos_right = _grid.two_column()
                    chart_title = sc.chart_insight if sc and sc.chart_insight else ""
                    chart_elem = SlideElement(
                        element_type="chart",
                        position=pos_left,
                        content=ChartContent(
                            chart_type=chart_type,
                            title=chart_title,
                            categories=categories,
                            series=series_list,
                        ),
                    )
                    break
        if chart_elem:
            break

    # If no chart, try infographic on left using AI items
    if not chart_elem:
        inf_items = []
        if sc and sc.infographic_items:
            inf_items = sc.infographic_items[:4]
        else:
            for sec in sections:
                for sub in sec.subsections[:4]:
                    inf_items.append(InfographicItem(
                        title=sub.heading[:40],
                        description=(sub.text or "")[:80],
                    ))
                if not inf_items and sec.bullets:
                    for b in sec.bullets[:4]:
                        parts = b.split(":", 1) if ":" in b else [b]
                        inf_items.append(InfographicItem(
                            title=parts[0].strip()[:40],
                            description=parts[1].strip()[:80] if len(parts) > 1 else "",
                        ))
        if inf_items:
            pos_left, pos_right = _grid.two_column()
            chart_elem = SlideElement(
                element_type="infographic",
                position=pos_left,
                content=InfographicContent(
                    infographic_type="comparison",
                    items=inf_items[:4],
                ),
            )

    if not chart_elem and not bullets:
        return _build_bullets_slide(plan, sections, sc)

    if chart_elem:
        elements.append(chart_elem)

    if bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.two_column()[1] if chart_elem else _grid.full(),
            content=BulletContent(items=bullets, font_size=13),
        ))

    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="mixed",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )


def _build_conclusion_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
    slide_plan: SlidePlan | None = None,
    sc: SlideContent | None = None,
) -> SlideSpec:
    """Build conclusion / key takeaways slide using AI-written synthesis."""
    # AI-written bullets preferred — conclusion should be AI-synthesized from entire deck
    if sc and sc.bullets:
        bullets = sc.bullets[:config.MAX_BULLETS_PER_SLIDE]
    else:
        bullets = []
        for sec in sections:
            if sec.bullets:
                bullets.extend(sec.bullets)
            elif sec.text:
                sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
                bullets.extend(sentences)
        bullets = bullets[:config.MAX_BULLETS_PER_SLIDE]

    title = sc.title if sc and sc.title else (plan.action_title or plan.title or "Key Takeaways")

    elements = []
    if bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.full(),
            content=BulletContent(items=bullets, font_size=14),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="conclusion",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=sc.subtitle if sc and sc.subtitle else plan.subtitle,
        elements=elements,
    )
