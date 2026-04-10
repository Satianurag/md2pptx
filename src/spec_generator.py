from __future__ import annotations
import logging
import json
import re
from typing import Optional
from .schemas import (
    ContentTree, ContentSection, SlidePlan, SlidePlanItem,
    SlideMasterInfo, PresentationSpec, SlideSpec, SlideElement,
    Position, TextContent, BulletContent, ChartContent, ChartSeries,
    TableContent, ShapeContent, InfographicContent, InfographicItem,
    DataTable, KeyMetric,
)
from .llm import invoke_llm
from .grid_system import grid
from . import config

logger = logging.getLogger(__name__)


def _smart_truncate(text: str, max_chars: int = 200) -> str:
    """Truncate text at a word boundary, never mid-word. Appends '…' only if truncated."""
    text = text.strip()
    if len(text) <= max_chars:
        return text
    # Find the last space before the limit
    truncated = text[:max_chars]
    last_space = truncated.rfind(' ')
    if last_space > max_chars * 0.6:  # only break at word if we keep >60% of content
        truncated = truncated[:last_space]
    return truncated.rstrip('.,;:- ') + '…'


# ── Grid-based position presets ──
POS_FULL = grid.full()
POS_LEFT_HALF, POS_RIGHT_HALF = grid.two_column()
POS_TOP_HALF, POS_BOTTOM_HALF = grid.top_bottom()
POS_CHART = grid.chart()
POS_TABLE = grid.table()


SPEC_SYSTEM_PROMPT = """\
You are a presentation content specialist. Given a slide plan item and the relevant source \
content from a research report, generate the EXACT text content for that slide.

RULES:
1. Be CONCISE. Max 6 bullet points per slide. Max 100 chars per bullet.
2. Extract the KEY insight — don't copy paragraphs verbatim.
3. For charts: extract actual numeric data (categories + series with float values).
4. For tables: extract headers and row data accurately.
5. For infographics: create 3-6 items with short title + description.
6. For KPI cards: extract 3-5 key metrics with label + value.
7. Never invent data — only use what's in the source content.
8. Write in professional business language.

Respond with a JSON object with these fields:
- "title": slide title (string)
- "subtitle": optional subtitle (string or null)
- "content_type": one of "bullets", "chart", "table", "infographic", "kpi", "text", "mixed"
- "bullets": list of strings (if bullets)
- "chart_type": one of "bar","column","line","pie","area","doughnut" (if chart)
- "chart_title": string (if chart)
- "categories": list of strings (if chart)
- "series": list of {"name": str, "values": list of floats} (if chart)
- "table_headers": list of strings (if table)
- "table_rows": list of list of strings (if table)
- "infographic_type": one of "process_flow","timeline","comparison","kpi_cards","hierarchy" (if infographic)
- "infographic_items": list of {"title": str, "description": str, "value": str or null} (if infographic)
- "text": plain text content (if text type)
"""


def generate_presentation_spec(
    content_tree: ContentTree,
    slide_plan: SlidePlan,
    master_info: SlideMasterInfo | None = None,
    template_path: str = "",
) -> PresentationSpec:
    """Convert a SlidePlan + ContentTree into a full PresentationSpec."""

    slides: list[SlideSpec] = []

    for plan_item in slide_plan.slides:
        slide_spec = _generate_slide_spec(content_tree, plan_item, slide_plan)
        slides.append(slide_spec)

    _ensure_visual_coverage(slides, content_tree)

    return PresentationSpec(
        title=content_tree.title,
        subtitle=content_tree.subtitle,
        slides=slides,
        template_path=template_path,
        target_slide_count=len(slides),
    )


def _ensure_visual_coverage(slides: list[SlideSpec], tree: ContentTree) -> None:
    """Ensure the deck contains dedicated visual slides without cluttering existing ones."""
    if not tree.all_tables:
        return

    if not _deck_has_element(slides, "chart"):
        for table in tree.all_tables:
            if not (table.rows and table.headers and len(table.headers) >= 2):
                continue
            categories, series = _extract_chart_data(table)
            if categories and series:
                if _insert_support_slide(
                    slides,
                    _build_support_chart_slide(table, categories, series),
                ):
                    logger.info("Added dedicated chart slide for deck-level visual coverage")
                break

    if not _deck_has_element(slides, "table"):
        best_table = None
        for table in tree.all_tables:
            if table.rows and table.headers and len(table.headers) >= 2:
                if _should_render_as_table(table) or best_table is None:
                    best_table = table
                    if _should_render_as_table(table):
                        break
        if best_table and _insert_support_slide(slides, _build_support_table_slide(best_table)):
            logger.info("Added dedicated table slide for deck-level visual coverage")

    _renumber_slides(slides)


def _deck_has_element(slides: list[SlideSpec], element_type: str) -> bool:
    return any(el.element_type == element_type for slide in slides for el in slide.elements)


def _insert_support_slide(slides: list[SlideSpec], support_slide: SlideSpec) -> bool:
    """Insert a new support slide before closing slides, or replace a low-value content slide."""
    insert_at = next(
        (idx for idx, slide in enumerate(slides) if slide.slide_type in ("conclusion", "thank_you")),
        len(slides),
    )
    if len(slides) < config.MAX_SLIDES:
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
                position=POS_CHART,
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
                position=POS_TABLE,
                content=TableContent(headers=headers, rows=rows),
            )
        ],
    )


def _generate_slide_spec(
    content_tree: ContentTree,
    plan_item: SlidePlanItem,
    slide_plan: SlidePlan | None = None,
) -> SlideSpec:
    """Generate a single SlideSpec from a SlidePlanItem."""

    # Special slides that don't need LLM
    if plan_item.slide_type == "cover":
        return _build_cover_slide(content_tree, plan_item)

    if plan_item.slide_type == "thank_you":
        return _build_thank_you_slide(plan_item)

    if plan_item.slide_type == "section_divider":
        return _build_divider_slide(plan_item)

    # For content slides, gather relevant source content
    source_sections = _find_source_sections(content_tree, plan_item.content_source)
    source_text = _sections_to_text(source_sections, content_tree)

    # Try rule-based generation first for simple cases
    rule_based = _try_rule_based_generation(plan_item, source_sections, content_tree, slide_plan)
    if rule_based:
        return rule_based

    # Fall back to LLM for complex content decisions
    return _llm_generate_slide(plan_item, source_text)


def _build_cover_slide(tree: ContentTree, plan: SlidePlanItem) -> SlideSpec:
    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="cover",
        layout_name="cover",
        title=tree.title or plan.title,
        subtitle=tree.subtitle or plan.subtitle or "",
        elements=[],
    )


def _build_thank_you_slide(plan: SlidePlanItem) -> SlideSpec:
    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="thank_you",
        layout_name="thank_you",
        title=plan.title or "Thank You",
        subtitle=plan.subtitle or "Questions & Discussion",
        elements=[],
    )


def _build_divider_slide(plan: SlidePlanItem) -> SlideSpec:
    subtitle = plan.subtitle or plan.key_message or ""
    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="section_divider",
        layout_name="divider",
        title=plan.title,
        subtitle=subtitle,
        elements=[],
    )


def _find_source_sections(
    tree: ContentTree,
    content_source: list[str],
) -> list[ContentSection]:
    """Find sections matching the content_source headings."""
    if not content_source:
        return []

    found = []
    source_lower = [s.lower().strip() for s in content_source]

    def search(sections: list[ContentSection]) -> None:
        for sec in sections:
            if sec.heading.lower().strip() in source_lower:
                found.append(sec)
            search(sec.subsections)

    search(tree.sections)
    return found


def _sections_to_text(sections: list[ContentSection], tree: ContentTree) -> str:
    """Convert sections to a text block for the LLM."""
    parts = []
    for sec in sections:
        parts.append(f"## {sec.heading}")
        if sec.text:
            parts.append(sec.text[:500])
        for b in sec.bullets[:8]:
            parts.append(f"- {b[:150]}")
        for t in sec.tables[:2]:
            parts.append(f"Table: {', '.join(t.headers[:6])}")
            for row in t.rows[:5]:
                parts.append(f"  | {' | '.join(str(c)[:30] for c in row[:6])} |")
        for m in sec.metrics[:5]:
            parts.append(f"Metric: {m.label} = {m.value}")
        for sub in sec.subsections:
            parts.append(f"### {sub.heading}")
            if sub.text:
                parts.append(sub.text[:300])
            for b in sub.bullets[:5]:
                parts.append(f"- {b[:120]}")

    result = "\n".join(parts)
    if len(result) > 4000:
        result = result[:4000] + "\n...[truncated]"
    return result


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
        nums = 0
        for row in table.rows:
            if col_idx < len(row):
                val = _parse_number(str(row[col_idx]))
                if val != 0 or str(row[col_idx]).strip() in ("0", "0.0", "0%"):
                    nums += 1
        if nums > len(table.rows) * 0.5:
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


_PROCESS_KEYWORDS = re.compile(
    r'\b(step|stage|phase|process|pipeline|workflow|procedure|method|approach|strategy)\b', re.I
)
_TIMELINE_KEYWORDS = re.compile(
    r'\b(20\d{2}|19\d{2}|year|quarter|Q[1-4]|month|before|after|first|then|finally|timeline)\b', re.I
)
_COMPARISON_KEYWORDS = re.compile(
    r'\b(vs\.?|versus|compared|comparison|advantage|disadvantage|pro|con|differ|alternative)\b', re.I
)


def _detect_infographic_pattern(sections: list[ContentSection]) -> tuple[str | None, str | None]:
    """Detect if content can be visualized as an infographic. Returns (infographic_type, None) or (None, None)."""
    all_bullets: list[str] = []
    all_text = ""
    has_metrics = False
    for sec in sections:
        all_bullets.extend(sec.bullets)
        all_text += sec.text + " "
        if sec.metrics:
            has_metrics = True
        for sub in sec.subsections:
            all_bullets.extend(sub.bullets)
            all_text += sub.text + " "
            if sub.metrics:
                has_metrics = True

    n = len(all_bullets)
    combined = " ".join(all_bullets) + " " + all_text

    # KPI cards: section has 3-6 metrics
    if has_metrics:
        metric_count = sum(len(s.metrics) for s in sections) + sum(
            len(sub.metrics) for s in sections for sub in s.subsections
        )
        if 2 <= metric_count <= 6:
            return "kpi_cards", None

    # Process flow: bullets with step-like patterns or numbered items
    numbered_count = sum(1 for b in all_bullets if re.match(r'^\d+[\.\)]\s', b))
    if numbered_count >= 3 or (n >= 3 and _PROCESS_KEYWORDS.search(combined)):
        return "process_flow", None

    # Timeline: chronological patterns
    if n >= 3 and _TIMELINE_KEYWORDS.search(combined):
        return "timeline", None

    # Comparison: comparative language
    if 2 <= n <= 6 and _COMPARISON_KEYWORDS.search(combined):
        return "comparison", None

    # Bullets with colons (key: value) → comparison or KPI
    colon_count = sum(1 for b in all_bullets if ":" in b and len(b.split(":")[0]) < 30)
    if colon_count >= 3:
        return "comparison", None

    return None, None


def _try_rule_based_generation(
    plan_item: SlidePlanItem,
    source_sections: list[ContentSection],
    tree: ContentTree,
    slide_plan: SlidePlan | None = None,
) -> SlideSpec | None:
    """Try to generate a slide without LLM for straightforward cases."""

    # Agenda slide
    if plan_item.slide_type == "agenda":
        return _build_agenda_slide(plan_item, tree)

    # Executive summary
    if plan_item.slide_type == "executive_summary":
        return _build_exec_summary_slide(plan_item, tree)

    # Conclusion slides should preserve their dedicated takeaways pattern and
    # should not be auto-converted into KPI/infographic/chart layouts.
    if plan_item.slide_type == "conclusion":
        return _build_conclusion_slide(plan_item, source_sections, tree, slide_plan)

    # Chart slides with table data
    if plan_item.visualization_hint == "chart" and source_sections:
        chart_slide = _build_chart_slide(plan_item, source_sections)
        if chart_slide:
            return chart_slide

    # Table slides
    if plan_item.visualization_hint == "table" and source_sections:
        table_slide = _build_table_slide(plan_item, source_sections)
        if table_slide:
            return table_slide

    # KPI / metric cards
    if plan_item.visualization_hint == "kpi" and source_sections:
        kpi_slide = _build_kpi_slide(plan_item, source_sections, tree)
        if kpi_slide:
            return kpi_slide

    # Infographic slides
    if plan_item.visualization_hint == "infographic" and source_sections:
        infographic_slide = _build_infographic_slide(plan_item, source_sections)
        if infographic_slide:
            return infographic_slide

    # Mixed slides: chart/infographic on left, bullets on right
    if plan_item.visualization_hint == "mixed" and source_sections:
        mixed_slide = _build_mixed_slide(plan_item, source_sections, tree)
        if mixed_slide:
            return mixed_slide

    # INFOGRAPHIC-FIRST: before falling back to bullets, detect visual patterns
    if plan_item.visualization_hint == "bullets" and source_sections:
        inf_type, _ = _detect_infographic_pattern(source_sections)
        if inf_type == "kpi_cards":
            kpi = _build_kpi_slide(plan_item, source_sections, tree)
            if kpi:
                return kpi
        elif inf_type:
            # Override the plan item's hint and build infographic
            logger.info(f"Infographic-first: converting slide {plan_item.slide_number} "
                        f"from bullets to {inf_type}")
            inf_slide = _build_infographic_slide_with_type(
                plan_item, source_sections, inf_type
            )
            if inf_slide:
                return inf_slide

    # CHART-FIRST / TABLE-FIRST: if bullets slide has table data, decide
    if plan_item.visualization_hint == "bullets" and source_sections:
        _tables_found = [
            t for sec in source_sections for t in sec.tables if t.rows and t.headers
        ] + [
            t for sec in source_sections for sub in sec.subsections
            for t in sub.tables if t.rows and t.headers
        ]
        if _tables_found:
            # Check if the first table is better as a table
            if _should_render_as_table(_tables_found[0]):
                tbl_slide = _build_table_slide(plan_item, source_sections)
                if tbl_slide:
                    logger.info(f"Table-first: slide {plan_item.slide_number} auto-upgraded to table")
                    return tbl_slide
            # Otherwise try chart/mixed
            mixed = _build_mixed_slide(plan_item, source_sections, tree)
            if mixed:
                logger.info(f"Chart-first: slide {plan_item.slide_number} auto-upgraded to mixed")
                return mixed
            chart = _build_chart_slide(plan_item, source_sections)
            if chart:
                logger.info(f"Chart-first: slide {plan_item.slide_number} auto-upgraded to chart")
                return chart

    # Plain bullet slides (last resort)
    if plan_item.visualization_hint == "bullets" and source_sections:
        return _build_bullets_slide(plan_item, source_sections)

    # Content slides with table data → try chart
    if plan_item.slide_type == "content" and source_sections:
        chart = _build_chart_slide(plan_item, source_sections)
        if chart:
            return chart
        return _build_bullets_slide(plan_item, source_sections)

    return None


def _build_agenda_slide(plan: SlidePlanItem, tree: ContentTree) -> SlideSpec:
    """Build agenda from section headings."""
    items = []
    for sec in tree.sections[:10]:
        heading = sec.heading.strip()
        if heading.lower() not in ("executive summary", "table of contents", "references", "references and source documentation"):
            items.append(heading)

    elements = []
    if items:
        elements.append(SlideElement(
            element_type="bullets",
            position=POS_FULL,
            content=BulletContent(items=items[:8], font_size=14),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="agenda",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title or "Agenda",
        elements=elements,
    )


def _build_exec_summary_slide(plan: SlidePlanItem, tree: ContentTree) -> SlideSpec:
    """Build executive summary slide."""
    summary = tree.executive_summary or ""
    elements = []

    if summary:
        # Break into bullets at sentence boundaries
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', summary) if s.strip()]
        # Take top 5 key sentences
        key_points = sentences[:5]
        if key_points:
            # Truncate each point
            key_points = [_smart_truncate(p, config.MAX_CHARS_PER_BULLET) for p in key_points]
            elements.append(SlideElement(
                element_type="bullets",
                position=POS_FULL,
                content=BulletContent(items=key_points, font_size=13),
            ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="executive_summary",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title or "Executive Summary",
        elements=elements,
    )


def _build_bullets_slide(plan: SlidePlanItem, sections: list[ContentSection]) -> SlideSpec:
    """Build a bullet list slide from section content."""
    bullets = []
    for sec in sections:
        # Prefer existing bullets
        if sec.bullets:
            bullets.extend(sec.bullets)
        elif sec.text:
            # Split text into sentences
            sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
            bullets.extend(sentences)
        # Also get subsection bullets
        for sub in sec.subsections:
            if sub.bullets:
                bullets.extend(sub.bullets)

    # Truncate
    bullets = [_smart_truncate(b, config.MAX_CHARS_PER_BULLET) for b in bullets[:config.MAX_BULLETS_PER_SLIDE]]

    elements = []
    if bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=POS_FULL,
            content=BulletContent(items=bullets, font_size=14),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type=plan.slide_type,
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title,
        subtitle=plan.subtitle,
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


def _build_chart_slide(plan: SlidePlanItem, sections: list[ContentSection]) -> SlideSpec | None:
    """Build a chart slide from table data in sections."""
    # Find the first table with numeric data
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
        return None

    chart_type = plan.chart_type_hint or _auto_detect_chart_type(table)
    categories, series_list = _extract_chart_data(table)

    if not categories or not series_list:
        return None

    elements = [
        SlideElement(
            element_type="chart",
            position=POS_CHART,
            content=ChartContent(
                chart_type=chart_type,
                title=table.title or plan.title,
                categories=categories,
                series=series_list,
            ),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="chart",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _extract_chart_data(table: DataTable) -> tuple[list[str], list[ChartSeries]]:
    """Extract chart-friendly data from a DataTable."""
    if len(table.headers) < 2 or not table.rows:
        return [], []

    # Strategy: first column = categories, remaining columns = series
    categories = []
    value_columns: dict[str, list[float]] = {h: [] for h in table.headers[1:]}

    for row in table.rows[:20]:  # cap rows
        if not row:
            continue
        categories.append(str(row[0])[:30])
        for i, header in enumerate(table.headers[1:], 1):
            val = _parse_number(row[i] if i < len(row) else "0")
            value_columns[header].append(val)

    series_list = []
    for name, values in value_columns.items():
        if any(v != 0 for v in values):  # skip all-zero series
            series_list.append(ChartSeries(name=name[:30], values=values))

    return categories, series_list


def _parse_number(s: str) -> float:
    """Parse a string into a float, handling currency/percent/comma formats."""
    s = str(s).strip()
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
        return 0.0


def _build_table_slide(plan: SlidePlanItem, sections: list[ContentSection]) -> SlideSpec | None:
    """Build a table slide."""
    table = None
    for sec in sections:
        for t in sec.tables:
            if t.headers and t.rows:
                table = t
                break
        if table:
            break

    if not table:
        return None

    # Limit table size
    headers = table.headers[:6]
    rows = [row[:6] for row in table.rows[:8]]

    elements = [
        SlideElement(
            element_type="table",
            position=POS_TABLE,
            content=TableContent(headers=headers, rows=rows),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="table",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _build_kpi_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
) -> SlideSpec | None:
    """Build a KPI cards slide from metrics."""
    metrics: list[KeyMetric] = []
    for sec in sections:
        metrics.extend(sec.metrics)
        for sub in sec.subsections:
            metrics.extend(sub.metrics)

    if not metrics:
        metrics = tree.all_metrics[:5]

    if not metrics:
        return None

    items = []
    for m in metrics[:5]:
        items.append(InfographicItem(
            title=_smart_truncate(m.label, 50),
            description="",
            value=m.value,
        ))

    elements = [
        SlideElement(
            element_type="infographic",
            position=POS_FULL,
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
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _build_infographic_slide(plan: SlidePlanItem, sections: list[ContentSection]) -> SlideSpec | None:
    """Build an infographic slide."""
    infographic_type = plan.infographic_type_hint or "process_flow"

    items = []
    for sec in sections:
        # Use subsection headings as items
        if sec.subsections:
            for sub in sec.subsections[:6]:
                desc = _smart_truncate(sub.text, 150) if sub.text else (_smart_truncate(sub.bullets[0], 150) if sub.bullets else "")
                items.append(InfographicItem(
                    title=_smart_truncate(sub.heading, 80),
                    description=desc,
                    value=sub.metrics[0].value if sub.metrics else None,
                ))
        elif sec.bullets:
            for b in sec.bullets[:6]:
                # Split bullet into title and description
                parts = b.split(":", 1) if ":" in b else (b.split("–", 1) if "–" in b else [b])
                title = _smart_truncate(parts[0].strip(), 80)
                desc = _smart_truncate(parts[1].strip(), 150) if len(parts) > 1 else ""
                items.append(InfographicItem(title=title, description=desc))

    if not items:
        return None

    elements = [
        SlideElement(
            element_type="infographic",
            position=POS_FULL,
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
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _build_infographic_slide_with_type(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    infographic_type: str,
) -> SlideSpec | None:
    """Build an infographic slide with a specific type (used by infographic-first detection)."""
    items = []
    for sec in sections:
        if sec.subsections:
            for sub in sec.subsections[:6]:
                desc = _smart_truncate(sub.text, 100) if sub.text else (_smart_truncate(sub.bullets[0], 100) if sub.bullets else "")
                items.append(InfographicItem(
                    title=_smart_truncate(sub.heading, 50),
                    description=desc,
                    value=sub.metrics[0].value if sub.metrics else None,
                ))
        elif sec.bullets:
            for b in sec.bullets[:6]:
                parts = b.split(":", 1) if ":" in b else (b.split("–", 1) if "–" in b else [b])
                title = _smart_truncate(parts[0].strip(), 50)
                desc = _smart_truncate(parts[1].strip(), 100) if len(parts) > 1 else ""
                items.append(InfographicItem(title=title, description=desc))
        elif sec.text:
            sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
            for s in sentences[:6]:
                items.append(InfographicItem(title=_smart_truncate(s, 50), description=_smart_truncate(s[50:], 100) if len(s) > 50 else ""))

    if len(items) < 2:
        return None

    elements = [
        SlideElement(
            element_type="infographic",
            position=POS_FULL,
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
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _build_mixed_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
) -> SlideSpec | None:
    """Build a mixed slide: visual element (left) + bullets (right) in two-column layout."""
    elements = []

    # Collect bullets
    bullets: list[str] = []
    for sec in sections:
        bullets.extend(sec.bullets[:4])
        for sub in sec.subsections:
            bullets.extend(sub.bullets[:3])
    bullets = [_smart_truncate(b, config.MAX_CHARS_PER_BULLET) for b in bullets[:config.MAX_BULLETS_PER_SLIDE]]

    # Try to build a chart for left side
    chart_elem = None
    for sec in sections:
        for t in sec.tables:
            if t.headers and t.rows and len(t.headers) >= 2:
                chart_type = plan.chart_type_hint or _auto_detect_chart_type(t)
                series_list = _extract_chart_series(t, chart_type)
                if series_list:
                    categories = [str(row[0]) for row in t.rows[:8]] if t.rows else []
                    chart_elem = SlideElement(
                        element_type="chart",
                        position=POS_LEFT_HALF,
                        content=ChartContent(
                            chart_type=chart_type,
                            title="",
                            categories=categories,
                            series=series_list,
                        ),
                    )
                    break
        if chart_elem:
            break

    # If no chart, try infographic on left
    if not chart_elem:
        inf_items = []
        for sec in sections:
            for sub in sec.subsections[:4]:
                inf_items.append(InfographicItem(
                    title=_smart_truncate(sub.heading, 40),
                    description=_smart_truncate(sub.text or "", 80),
                ))
            if not inf_items and sec.bullets:
                for b in sec.bullets[:4]:
                    parts = b.split(":", 1) if ":" in b else (b.split("–", 1) if "–" in b else [b])
                    inf_items.append(InfographicItem(
                        title=_smart_truncate(parts[0].strip(), 40),
                        description=_smart_truncate(parts[1].strip(), 80) if len(parts) > 1 else "",
                    ))
        if inf_items:
            chart_elem = SlideElement(
                element_type="infographic",
                position=POS_LEFT_HALF,
                content=InfographicContent(
                    infographic_type="comparison",
                    items=inf_items[:4],
                ),
            )

    if not chart_elem and not bullets:
        return None

    if chart_elem:
        elements.append(chart_elem)

    if bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=POS_RIGHT_HALF if chart_elem else POS_FULL,
            content=BulletContent(items=bullets, font_size=13),
        ))

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="mixed",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title,
        subtitle=plan.subtitle,
        elements=elements,
    )


def _build_conclusion_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
    slide_plan: SlidePlan | None = None,
) -> SlideSpec:
    """Build conclusion / key takeaways slide with robust multi-source fallback."""
    bullets = []

    # Source 1: Try to get content from conclusion sections
    for sec in sections:
        if sec.bullets:
            bullets.extend(sec.bullets)
        elif sec.text:
            sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
            bullets.extend(sentences)

    # Source 2: Aggregate key_messages from all plan items
    if not bullets and slide_plan:
        for item in slide_plan.slides:
            if item.key_message and item.key_message.strip():
                bullets.append(item.key_message.strip())

    # Source 3: Executive summary key sentences
    if not bullets and tree.executive_summary:
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', tree.executive_summary) if s.strip()]
        bullets.extend(sentences[:6])

    # Source 4: First bullets from each top section
    if not bullets:
        for sec in tree.sections[:8]:
            if sec.bullets:
                bullets.append(sec.bullets[0])
            elif sec.text:
                first_sentence = re.split(r'(?<=[.!?])\s+', sec.text)
                if first_sentence:
                    bullets.append(first_sentence[0].strip())

    # Absolute fallback
    if not bullets:
        bullets = ["Key findings and strategic recommendations from this analysis"]

    cleaned: list[str] = []
    seen: set[str] = set()
    for bullet in bullets:
        normalized = bullet.strip().lstrip("•-").strip()
        if not normalized:
            continue
        key = normalized.lower()
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(normalized)

    if len(cleaned) < 3 and tree.executive_summary:
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', tree.executive_summary) if s.strip()]
        for sentence in sentences:
            key = sentence.lower()
            if key not in seen:
                seen.add(key)
                cleaned.append(sentence)
            if len(cleaned) >= config.MAX_BULLETS_PER_SLIDE:
                break

    if len(cleaned) < 3:
        for sec in tree.sections[:8]:
            candidate = ""
            if sec.bullets:
                candidate = sec.bullets[0].strip()
            elif sec.text:
                sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if s.strip()]
                candidate = sentences[0] if sentences else ""
            key = candidate.lower()
            if candidate and key not in seen:
                seen.add(key)
                cleaned.append(candidate)
            if len(cleaned) >= config.MAX_BULLETS_PER_SLIDE:
                break

    if not cleaned:
        cleaned = ["Key findings and strategic recommendations from this analysis"]

    bullets = [_smart_truncate(b, config.MAX_CHARS_PER_BULLET) for b in cleaned[:config.MAX_BULLETS_PER_SLIDE]]

    elements = [
        SlideElement(
            element_type="bullets",
            position=POS_FULL,
            content=BulletContent(items=bullets, font_size=14),
        )
    ]

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type="conclusion",
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan.title or "Key Takeaways",
        subtitle=plan.subtitle,
        elements=elements,
    )


def _llm_generate_slide(plan_item: SlidePlanItem, source_text: str) -> SlideSpec:
    """Use LLM to generate slide content for complex cases."""
    user_prompt = f"""\
Generate slide content for:
- Slide type: {plan_item.slide_type}
- Title: {plan_item.title}
- Visualization: {plan_item.visualization_hint}
- Key message: {plan_item.key_message}

Source content:
{source_text}

Return JSON with the slide content.\
"""

    try:
        raw = invoke_llm(
            system_prompt=SPEC_SYSTEM_PROMPT,
            user_prompt=user_prompt,
            estimated_tokens=4000,
        )
        return _parse_llm_slide_response(raw, plan_item)
    except Exception as e:
        logger.warning(f"LLM slide generation failed for slide {plan_item.slide_number}: {e}")
        # Fallback: create a simple text slide
        return _fallback_text_slide(plan_item, source_text)


def _parse_llm_slide_response(raw: str, plan_item: SlidePlanItem) -> SlideSpec:
    """Parse the LLM JSON response into a SlideSpec."""
    # Extract JSON from response (may be wrapped in markdown code blocks)
    json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', raw, re.DOTALL)
    if json_match:
        json_str = json_match.group(1)
    else:
        # Try to find raw JSON
        json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', raw, re.DOTALL)
        json_str = json_match.group(0) if json_match else raw

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError:
        logger.warning(f"Failed to parse LLM JSON for slide {plan_item.slide_number}")
        return _fallback_text_slide(plan_item, "")

    elements = []
    content_type = data.get("content_type", "bullets")

    if content_type == "bullets" and data.get("bullets"):
        items = [_smart_truncate(str(b), config.MAX_CHARS_PER_BULLET) for b in data["bullets"][:config.MAX_BULLETS_PER_SLIDE]]
        elements.append(SlideElement(
            element_type="bullets",
            position=POS_FULL,
            content=BulletContent(items=items, font_size=14),
        ))

    elif content_type == "chart" and data.get("categories") and data.get("series"):
        series_list = []
        for s in data["series"][:5]:
            vals = [float(v) for v in s.get("values", [])]
            if vals:
                series_list.append(ChartSeries(name=str(s.get("name", ""))[:30], values=vals))
        if series_list:
            elements.append(SlideElement(
                element_type="chart",
                position=POS_CHART,
                content=ChartContent(
                    chart_type=data.get("chart_type", "column"),
                    title=data.get("chart_title", ""),
                    categories=[str(c)[:30] for c in data["categories"]],
                    series=series_list,
                ),
            ))

    elif content_type == "table" and data.get("table_headers") and data.get("table_rows"):
        elements.append(SlideElement(
            element_type="table",
            position=POS_TABLE,
            content=TableContent(
                headers=[str(h)[:30] for h in data["table_headers"][:6]],
                rows=[[str(c)[:30] for c in row[:6]] for row in data["table_rows"][:8]],
            ),
        ))

    elif content_type in ("infographic", "kpi") and data.get("infographic_items"):
        items = []
        for item in data["infographic_items"][:6]:
            items.append(InfographicItem(
                title=_smart_truncate(str(item.get("title", "")), 50),
                description=_smart_truncate(str(item.get("description", "")), 100),
                value=str(item.get("value", "")) if item.get("value") else None,
            ))
        inf_type = data.get("infographic_type", "kpi_cards")
        if inf_type not in ("process_flow", "timeline", "comparison", "kpi_cards", "hierarchy"):
            inf_type = "kpi_cards"
        elements.append(SlideElement(
            element_type="infographic",
            position=POS_FULL,
            content=InfographicContent(infographic_type=inf_type, items=items),
        ))

    elif content_type == "text" and data.get("text"):
        elements.append(SlideElement(
            element_type="text",
            position=POS_FULL,
            content=TextContent(text=data["text"][:500], font_size=14),
        ))

    # If nothing was parsed, create fallback
    if not elements:
        bullets = data.get("bullets", [])
        if bullets:
            items = [_smart_truncate(str(b), config.MAX_CHARS_PER_BULLET) for b in bullets[:config.MAX_BULLETS_PER_SLIDE]]
            elements.append(SlideElement(
                element_type="bullets",
                position=POS_FULL,
                content=BulletContent(items=items, font_size=14),
            ))

    return SlideSpec(
        slide_number=plan_item.slide_number,
        slide_type=plan_item.slide_type,
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=data.get("title", plan_item.title),
        subtitle=data.get("subtitle", plan_item.subtitle),
        elements=elements,
    )


def _fallback_text_slide(plan_item: SlidePlanItem, source_text: str) -> SlideSpec:
    """Create a minimal text slide as fallback."""
    text = plan_item.key_message or source_text[:300] or "Content pending"
    elements = [
        SlideElement(
            element_type="text",
            position=POS_FULL,
            content=TextContent(text=text, font_size=14),
        )
    ]
    return SlideSpec(
        slide_number=plan_item.slide_number,
        slide_type=plan_item.slide_type,
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=plan_item.title,
        subtitle=plan_item.subtitle,
        elements=elements,
    )
