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

# ── Per-deck archetype rotation ─────────────────────────────────────
# Tracks the archetype we chose for the previous content slide so we can
# rotate across consecutive slides (avoiding the "every slide is a sidebar
# icon_list" monotony seen in the Accenture deck). Reset at the start of
# every ``generate_presentation_spec`` call — strictly in-process state.
_archetype_history: list[str] = []


def generate_presentation_spec(
    content_tree: ContentTree,
    slide_plan: SlidePlan,
    master_info: SlideMasterInfo | None = None,
    template_path: str = "",
    content_profile: ContentProfile | None = None,
    deck_content: DeckContent | None = None,
) -> PresentationSpec:
    """Convert a SlidePlan + ContentTree + AI-written DeckContent into a full PresentationSpec."""
    global _grid, _archetype_history
    _grid = Grid.from_template(master_info) if master_info else Grid.default()
    _archetype_history = []  # Fresh rotation state per deck
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
    chart_type = _auto_detect_chart_type(table)
    # Horizontal bar charts place tiny values in the category-label gutter
    # on a log axis; keep log-scale for column/line/area only.
    use_log = _needs_log_scale(series) and chart_type in ("column", "line", "area")
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
                    chart_type=chart_type,
                    title=table.title or title,
                    categories=categories,
                    series=series,
                    log_scale=use_log,
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


_DATE_LIKE_RE = re.compile(
    r"^\s*(?:"
    r"\d{4}[-/.\s]\d{1,2}[-/.\s]\d{1,2}"                  # 2026-04-17 | 2026 04 17
    r"|\d{1,2}[-/.\s]\d{1,2}[-/.\s]\d{2,4}"                # 17-04-2026 | 17 04 2026
    r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2}"
    r"|Q[1-4][\s-]?\d{2,4}"                                # Q1 2024
    r")\s*$",
    re.IGNORECASE,
)

_ID_LIKE_RE = re.compile(r"^\d{7,}$")   # 7+ consecutive digits with no separators → likely ID/timestamp


def _parse_number(s: str) -> float:
    """Parse a string into a float, handling currency/percent/comma formats.

    Returns ``math.nan`` for non-numeric values like 'N/A', empty strings,
    purely textual content, date-like strings, or long digit IDs (7+ digits
    with no separators — typically dates or codes, never metric values).
    """
    s = str(s).strip()
    if not s or s.upper() in ('N/A', 'NA', '-', '—', 'N.A.', 'N.A', 'NULL', 'NONE', 'TBD', 'TBC'):
        return math.nan

    # Reject date-like strings (e.g., 2026-04-17, Apr 17, Q1 2024)
    if _DATE_LIKE_RE.match(s):
        return math.nan

    # Reject raw digit IDs (e.g., 20260123, 865000000 would need suffix) — these are never metrics
    cleaned = re.sub(r'[,$€£%\s]', '', s)
    if _ID_LIKE_RE.match(cleaned):
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


def _is_date_column(table: DataTable, col_idx: int) -> bool:
    """Return True if the column contains mostly date-like strings."""
    total = 0
    date_like = 0
    for row in table.rows:
        if col_idx < len(row):
            cell = str(row[col_idx]).strip()
            if not cell:
                continue
            total += 1
            if _DATE_LIKE_RE.match(cell):
                date_like += 1
    if total == 0:
        return False
    return date_like / total > 0.50


def _is_year_column(table: DataTable, col_idx: int) -> bool:
    """Return True if the column is predominantly 4-digit years (1900–2100).

    A year column is *technically* numeric but plotting years as bar heights
    produces nonsense charts (the y-axis labelled ``2022…2031``). Treat year
    columns as categorical x-axis labels instead of numeric series.
    """
    total = 0
    year_like = 0
    for row in table.rows:
        if col_idx < len(row):
            cell = str(row[col_idx]).strip()
            if not cell:
                continue
            total += 1
            try:
                v = int(float(re.sub(r"[,\s]", "", cell)))
                if 1900 <= v <= 2100:
                    year_like += 1
            except (ValueError, TypeError):
                pass
    if total == 0:
        return False
    return year_like / total > 0.80


def _is_numeric_column(table: DataTable, col_idx: int) -> bool:
    """Return True if > 40% of non-empty cells are genuinely numeric (not dates/IDs/years)."""
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
    # Require a higher bar: ≥60% numeric AND not a date or year column
    if _is_date_column(table, col_idx):
        return False
    if _is_year_column(table, col_idx):
        return False
    return numeric / total > 0.60


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


def _bullets_to_infographic_items(bullets: list[str]) -> list[InfographicItem]:
    """Convert raw bullet strings into InfographicItem instances.

    Splits each bullet at ``:`` or ``—`` to separate title from description
    where possible. Titles that overflow the icon-row zone are cut at a
    word boundary (never mid-word) so users never see dangling stubs like
    ``...focusing on uni``.
    """
    def _cut_at_word(s: str, max_chars: int) -> str:
        s = (s or "").strip()
        if len(s) <= max_chars:
            return s
        cut = s[:max_chars]
        idx = cut.rfind(" ")
        if idx >= max_chars * 0.55:
            cut = cut[:idx]
        return cut.rstrip(" ,.;:-") + "…"

    items: list[InfographicItem] = []
    for raw in bullets:
        text = (raw or "").strip()
        if not text:
            continue
        title = text
        desc = ""
        for sep in (": ", " — ", " – ", " - "):
            if sep in text:
                head, tail = text.split(sep, 1)
                if 5 <= len(head) <= 100 and tail.strip():
                    title = head.strip()
                    desc = tail.strip()
                    break
        # If title is still long and contains a sentence-ending period, use
        # the first sentence. Match `.` followed by whitespace or end-of-string
        # so we don't split "$1.5B" at the decimal point.
        if len(title) > 110 and re.search(r"\.(?:\s|$)", title):
            m = re.search(r"\.(?:\s|$)", title)
            cut_at = m.start() if m else -1
            head = title[:cut_at].strip() if cut_at > 0 else title
            if 10 <= len(head) <= 110:
                desc = title[cut_at + 1:].strip() + (" " + desc if desc else "")
                title = head
        items.append(InfographicItem(
            title=_cut_at_word(title, 110),
            description=_cut_at_word(desc, 180),
        ))
    return items


def _extract_enumeration(bullet: str) -> list[str]:
    """Return 2+ items extracted from an enumeration inside *bullet*, else ``[]``.

    Handles shapes like:
      - ``such as A, B, and C``
      - ``including A, B, C``
      - ``in A, B, and C`` / ``across A, B, C`` / ``for A, B, C`` (when items
        are capitalised proper nouns — keeps us from splitting prose lists)
      - ``between A and B``
      - ``A versus B`` / ``A vs B`` / ``A vs. B``
      - ``A: X, Y, Z``  (after a colon)

    Used to enrich single-bullet slides that would otherwise fall through to
    a sparse ``pull_quote`` layout.
    """
    if not bullet:
        return []
    # Trigger phrases that precede an explicit list
    for trigger in (r"\bsuch as\b", r"\bincluding\b", r"\bnamely\b", r"\be\.g\.", r"\bi\.e\.",):
        m = re.search(trigger + r"[:,\s]+(.+?)(?:\.\s|$)", bullet, re.IGNORECASE)
        if m:
            tail = m.group(1)
            parts = _split_enum_tail(tail)
            if len(parts) >= 2:
                return parts
    # Preposition + comma-separated proper-noun list: "in Europe, APAC, and North America"
    for prep in (r"\bin\b", r"\bacross\b", r"\bfor\b", r"\bfrom\b"):
        m = re.search(
            prep + r"\s+((?:[A-Z][\w\-\u2019'&/]*(?:\s+[A-Z][\w\-\u2019'&/]*)*)"
                   r"(?:\s*,\s*(?:and\s+)?[A-Z][\w\-\u2019'&/]*(?:\s+[A-Z][\w\-\u2019'&/]*)*){1,5})",
            bullet,
        )
        if m:
            parts = _split_enum_tail(m.group(1))
            if len(parts) >= 2:
                return parts
    # "A versus B" / "A vs B"
    m = re.search(r"\b([A-Z][\w\s\-\u2019']{2,40})\s+(?:vs\.?|versus)\s+([A-Z][\w\s\-\u2019']{2,40})\b", bullet)
    if m:
        return [m.group(1).strip(), m.group(2).strip()]
    # "between A and B" — capture two noun-phrases
    m = re.search(r"\bbetween\s+([\w\s\-\u2019']{3,40})\s+and\s+([\w\s\-\u2019']{3,40})\b", bullet, re.IGNORECASE)
    if m:
        return [m.group(1).strip().rstrip(",."), m.group(2).strip().rstrip(",.")]
    # Colon-prefixed list: "Key domains: AI, cybersecurity, data"
    if ":" in bullet:
        head, tail = bullet.split(":", 1)
        parts = _split_enum_tail(tail)
        if len(parts) >= 2 and len(head) < 60:
            return parts
    return []


def _split_enum_tail(tail: str) -> list[str]:
    """Split an enumeration tail like "A, B, and C" → ["A", "B", "C"]."""
    tail = tail.strip().rstrip(".;:")
    # Replace "Oxford" and/or final "and" with a comma
    tail = re.sub(r",?\s+and\s+", ",", tail, count=1)
    parts = [p.strip().strip('"\'') for p in tail.split(",")]
    parts = [p for p in parts if 2 <= len(p) <= 60]
    # Require each item to contain at least one letter
    parts = [p for p in parts if re.search(r"[A-Za-z]", p)]
    return parts[:4]


def _build_bullets_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build a content slide — routes to an infographic archetype when possible.

    Routing rules (content-shape driven, not LLM-driven — deterministic):
    - 3–6 short bullets → ``icon_list`` (auto-picked icons, replaces numbered cards)
    - 2 items with strong contrast → ``comparison``
    - 1 bullet with enumeration → enriched ``comparison``
    - 1 bullet (pure quote) → ``pull_quote`` with sub-attribution
    - Empty / no bullets → text fallback
    """
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

    # Drop empty / punctuation-only bullets (LLM sometimes emits "-", "—", etc.)
    bullets = [
        b.strip() for b in (b or "" for b in bullets)
        if b.strip() and re.sub(r"[\W_]+", "", b.strip())
    ]
    bullets = bullets[:config.MAX_BULLETS_PER_SLIDE]

    title = sc.title if sc and sc.title else (plan.action_title or plan.title)
    subtitle = sc.subtitle if sc and sc.subtitle else plan.subtitle
    key_msg = sc.key_takeaway if sc and sc.key_takeaway else (plan.key_message.strip() if plan.key_message else "")

    elements: list[SlideElement] = []

    # ── Enrichment: single-bullet slides are visually sparse. Try to extract
    # an enumeration from the bullet; if that fails but we have a key_msg,
    # promote it to a second bullet so the archetype picker can do better. ──
    if len(bullets) == 1:
        enum_parts = _extract_enumeration(bullets[0])
        if len(enum_parts) >= 2:
            # Replace the lone prose bullet with the extracted facets so we
            # can render a proper comparison grid.
            bullets = enum_parts
        elif key_msg and len(key_msg) > 30:
            # Pair the bullet with the takeaway as 2 complementary statements.
            bullets = [bullets[0], key_msg]

    # Archetype decision tree
    archetype = _pick_content_archetype(bullets, key_msg, plan)

    if archetype == "icon_list":
        items = _bullets_to_infographic_items(bullets)
        if items:
            elements.append(SlideElement(
                element_type="infographic",
                position=_grid.full(),
                content=InfographicContent(
                    infographic_type="icon_list",
                    items=items,
                ),
            ))
    elif archetype == "pull_quote" and bullets:
        # Attach the key_takeaway as attribution so the canvas isn't empty.
        attribution = key_msg[:80] if key_msg else ""
        quote_item = InfographicItem(
            title=bullets[0][:220],
            description=(bullets[1][:120] if len(bullets) > 1 else ""),
            value=attribution or None,
        )
        elements.append(SlideElement(
            element_type="infographic",
            position=_grid.full(),
            content=InfographicContent(
                infographic_type="pull_quote",
                items=[quote_item],
            ),
        ))
    elif archetype == "comparison" and len(bullets) >= 2:
        items = _bullets_to_infographic_items(bullets)[:3]
        elements.append(SlideElement(
            element_type="infographic",
            position=_grid.full(),
            content=InfographicContent(
                infographic_type="comparison",
                items=items,
            ),
        ))
    elif archetype == "sidebar" and key_msg and bullets and len(bullets) >= 3:
        pos_side, pos_main = _grid.sidebar_main()
        elements.append(SlideElement(
            element_type="text",
            position=pos_side,
            content=_create_text_content(text=key_msg, font_size=13, italic=True),
        ))
        elements.append(SlideElement(
            element_type="infographic",
            position=pos_main,
            content=InfographicContent(
                infographic_type="icon_list",
                items=_bullets_to_infographic_items(bullets),
            ),
        ))
    elif bullets:
        # Last-resort plain text bullets (truly no structure worth visualising)
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.full(),
            content=BulletContent(items=bullets, font_size=14),
        ))

    # Emergency fallback: if we have a title + no content at all, surface the
    # key_message / subtitle as a pull-quote so the slide isn't a bare header.
    if not elements:
        fallback_text = (key_msg or subtitle or plan.key_message or "").strip()
        if fallback_text and len(fallback_text) > 15:
            quote_item = InfographicItem(
                title=fallback_text[:220],
                description="",
            )
            elements.append(SlideElement(
                element_type="infographic",
                position=_grid.full(),
                content=InfographicContent(
                    infographic_type="pull_quote",
                    items=[quote_item],
                ),
            ))
            archetype = "pull_quote"

    # Record the final archetype so the next slide can rotate away from it.
    _archetype_history.append(archetype)

    return SlideSpec(
        slide_number=plan.slide_number,
        slide_type=plan.slide_type,
        layout_name=config.LAYOUT_TITLE_ONLY,
        title=title,
        subtitle=subtitle,
        elements=elements,
    )


def _pick_content_archetype(bullets: list[str], key_msg: str, plan: SlidePlanItem) -> str:
    """Deterministically pick the best archetype for a content-bullets slide.

    Reads ``_archetype_history`` (reset per deck run) to avoid repeating the
    same layout on consecutive slides — the reference Accenture deck looked
    monotonous because 3 of its 15 slides all routed to ``sidebar``.
    """
    n = len(bullets)
    avg_len = sum(len(b) for b in bullets) / max(n, 1) if bullets else 0
    prev = _archetype_history[-1] if _archetype_history else ""

    # 1 short punchy bullet — routed by the enrichment step in the caller.
    # This function only sees the enriched bullet list, so a single bullet
    # here means enrichment failed and we genuinely have a quote.
    if n == 1 and avg_len < 160:
        return "pull_quote"

    # 2 items with similar length → comparison (side-by-side)
    if n == 2 and avg_len < 220:
        return "comparison"

    # ≥7 bullets or very long prose — let the fallback render plain bullets
    if n >= 7 or avg_len >= 240:
        return "plain"

    # Long sidebar key-message + many bullets → sidebar layout by default,
    # BUT rotate off sidebar when the previous slide already used it.
    sidebar_eligible = bool(key_msg) and n >= 4 and len(key_msg) > 40

    if sidebar_eligible and prev != "sidebar":
        return "sidebar"

    # 3-6 bullets → icon_list (the default replacement for numbered cards),
    # except rotate to sidebar (if available) or comparison when the
    # previous slide was also icon_list.
    if 3 <= n <= 6:
        if prev == "icon_list":
            if sidebar_eligible:
                return "sidebar"
            if n <= 4:
                return "comparison"
            return "plain"
        return "icon_list"

    return "plain"


def _auto_detect_chart_type(table: DataTable) -> str:
    """Automatically select the best chart type based on data structure.

    Priority order:
    1. Time-series first column (date / year / quarter) → line or area
    2. Single series with few positive categories → pie/doughnut
    3. Many categories (>6) with long labels → bar (horizontal)
    4. Default → column (vertical)
    """
    n_rows = len(table.rows)
    n_cols = len(table.headers) - 1  # exclude category column

    if n_cols <= 0:
        return "column"

    first_col_values = [str(row[0]).strip() if row else "" for row in table.rows]

    # Time-series: date-like strings or years → line (preferred) / area (when ≥4 series)
    year_pattern = sum(1 for v in first_col_values if re.match(r'^(19|20)\d{2}', v))
    quarter_pattern = sum(1 for v in first_col_values if re.match(r'^Q[1-4]', v, re.I))
    date_pattern = sum(1 for v in first_col_values if _DATE_LIKE_RE.match(v))
    if year_pattern >= 3 or quarter_pattern >= 3 or date_pattern >= 3:
        return "line" if n_cols <= 3 else "area"

    # Single series with few positive categories → pie/doughnut, but only
    # when the values actually represent shares of a single whole. If the
    # values are a grab-bag of unrelated percentages (92.5 + 34 + 51 = 177.5)
    # a pie chart is misleading — fall back to a horizontal bar instead.
    if n_cols == 1 and 2 <= n_rows <= 8:
        values = [_parse_number(row[1] if len(row) > 1 else "0") for row in table.rows]
        values = [v for v in values if not math.isnan(v)]
        all_positive = all(v >= 0 for v in values)
        if all_positive and len(values) >= 2 and sum(values) > 0 and len(set(values)) > 1:
            # Detect percentage unit from the value column
            is_pct = _detect_column_unit(table, 1) == "pct"
            total = sum(values)
            if is_pct:
                # For percentages, only use pie if values roughly sum to 100 (± 10%)
                if 85 <= total <= 115:
                    return "pie" if n_rows <= 5 else "doughnut"
                # Otherwise: independent percentages → horizontal bar chart
                return "bar"
            # For raw counts, pie is fine as long as values aren't wildly skewed
            if max(values) <= 10 * min(v for v in values if v > 0):
                return "pie" if n_rows <= 5 else "doughnut"
            return "bar"

    # Many categories → horizontal bar (long labels fit better)
    if n_rows > 6:
        return "bar"

    # Default → vertical column
    return "column"


def _extract_chart_series(table: DataTable, chart_type: str) -> list[ChartSeries]:
    """Extract chart series from table data, adapted for the chart type."""
    categories, series_list = _extract_chart_data(table)
    return series_list


def _needs_log_scale(series_list: list[ChartSeries]) -> bool:
    """Return True if combined series values span >2 orders of magnitude.

    Additional guards (learned the hard way from the UAE solar deck):
    - Require the minimum positive value to be ≥ 1 so tiny values don't
      fall off the axis into the category-label gutter (log10(0.3) < 0).
    - Still require ≥2 positive values and a max/min ratio > 100.
    """
    all_vals: list[float] = []
    for s in series_list:
        all_vals.extend(v for v in s.values if v and not math.isnan(v))
    if not all_vals:
        return False
    pos = [abs(v) for v in all_vals if v > 0]
    if len(pos) < 2:
        return False
    if min(pos) < 1.0:
        return False
    return max(pos) / min(pos) > 100


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
        return _build_bullets_slide(plan, sections, sc)

    chart_type = plan.chart_type_hint or _auto_detect_chart_type(table)
    # Override pie/doughnut hints that can't actually be rendered as a pie —
    # fall back to the auto-detected choice (which applies the "percentages
    # must sum to ~100" rule).
    if chart_type in ("pie", "doughnut"):
        auto = _auto_detect_chart_type(table)
        if auto not in ("pie", "doughnut"):
            chart_type = auto
    categories, series_list = _extract_chart_data(table)

    if not categories or not series_list:
        logger.info(f"Chart slide '{plan.title}' has no valid numeric data; falling back to bullets")
        return _build_bullets_slide(plan, sections, sc)

    # If there's only one category or one value in total, chart is pointless
    if len(categories) < 2 or sum(len(s.values) for s in series_list) < 2:
        return _build_bullets_slide(plan, sections, sc)

    # Log scale when values span more than 2 orders of magnitude — excluded
    # for horizontal bar charts where tiny values overflow into the labels.
    use_log = _needs_log_scale(series_list) and chart_type in ("column", "line", "area")

    # Use AI-written chart insight as the chart title; fall back to table title then plan title
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
                log_scale=use_log,
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


def _smart_label(s: str, max_chars: int = 28) -> str:
    """Shorten a category label at word boundary when possible.

    Preserves full text up to *max_chars*, then cuts at the last space
    before that limit. Never mid-word. Adds an ellipsis only when cut.
    """
    s = str(s).strip()
    if len(s) <= max_chars:
        return s
    cut = s[:max_chars]
    last_space = cut.rfind(' ')
    if last_space >= max_chars * 0.5:
        cut = cut[:last_space]
    return cut.rstrip(" ,.;:-") + "…"


def _dedupe_categories(cats: list[str]) -> list[str]:
    """Ensure every category label is unique — append a disambiguator suffix to collisions.

    If two entries shorten to the same label (truncation collision), we append
    ``(2)``, ``(3)`` etc. This matches the Excel-style behaviour while making
    the duplication visible so the root cause is obvious.
    """
    seen: dict[str, int] = {}
    out: list[str] = []
    for c in cats:
        if c in seen:
            seen[c] += 1
            out.append(f"{c} ({seen[c]})")
        else:
            seen[c] = 1
            out.append(c)
    return out


def _extract_chart_data(table: DataTable) -> tuple[list[str], list[ChartSeries]]:
    """Extract chart-friendly data from a DataTable.

    Improvements over the naive 'first-col = categories, rest = series':
    1. Skips text-heavy columns that aren't chartable (e.g. 'Bank', 'Year').
    2. Handles N/A / missing values — rows where ALL numeric cells are NaN
       are dropped; remaining NaN values become 0.0.
    3. Separates columns by unit type (%, $, raw) and only charts the
       dominant unit group to prevent misleading mixed-axis visuals.
    4. Rejects flat series (all equal values) — caller should fall back to
       bullets/infographic when this happens.
    5. De-duplicates categories so truncation collisions don't produce
       invisible chart bars.
    """
    if len(table.headers) < 2 or not table.rows:
        return [], []

    # Identify which columns (index ≥ 1) are genuinely numeric (not dates/IDs)
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

    best_unit = max(unit_groups, key=lambda u: (len(unit_groups[u]), {"raw": 2, "currency": 1, "pct": 0}.get(u, 0)))
    selected_cols = unit_groups[best_unit]

    # First column = categories; collect values only from selected columns
    raw_categories: list[str] = []
    raw_values: dict[int, list[float]] = {ci: [] for ci in selected_cols}

    for row in table.rows[:20]:
        if not row:
            continue
        raw_categories.append(_smart_label(str(row[0]), max_chars=28))
        for ci in selected_cols:
            val = _parse_number(row[ci] if ci < len(row) else "")
            raw_values[ci].append(val)

    # Drop rows where either (a) the category is blank or (b) ALL numeric
    # values are NaN. Blank categories yield unlabelled bars which look broken
    # (the exact defect on slide 12 of the AI Bubble deck).
    keep_mask = []
    for row_idx, cat in enumerate(raw_categories):
        has_value = any(
            not math.isnan(raw_values[ci][row_idx]) for ci in selected_cols
        )
        has_label = bool(cat and cat.strip())
        keep_mask.append(has_value and has_label)

    categories = [c for c, keep in zip(raw_categories, keep_mask) if keep]
    if not categories:
        return [], []

    # De-dup so collisions don't produce invisible bars
    categories = _dedupe_categories(categories)

    series_list = []
    for ci in selected_cols:
        name = _smart_label(table.headers[ci], max_chars=28)
        values = [
            (0.0 if math.isnan(v) else v)
            for v, keep in zip(raw_values[ci], keep_mask)
            if keep
        ]
        if not values:
            continue
        # Skip all-zero series (nothing to show)
        non_zero = [v for v in values if v != 0]
        if not non_zero:
            continue
        # Skip flat series (every non-zero value identical — produces meaningless chart)
        if len(set(non_zero)) == 1 and len(non_zero) == len(values):
            logger.debug(f"Skipping flat series '{name}': all values = {non_zero[0]}")
            continue
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


_METRIC_VALUE_RE = re.compile(
    r"^\s*(?:"
    r"[\$€£¥]?\s*-?\d[\d,.]*\s*(?:[KkMmBbTt])?"          # 326, $6.6B, 51%, 12.5%
    r"(?:\s*[KkMmBbTt])?"
    r"(?:\s*%|\s*[KkMmBbTt][Bb]?)?"
    r"|[~≈>±<≤≥]?\s*\d[\d,.]*\s*(?:[KkMmBbTt])?\s*%?"
    r")\s*$",
)


def _sanitize_kpi_value(value: str | None) -> tuple[str, bool]:
    """Return (cleaned_value, is_valid).

    A valid KPI value looks like a metric (number + optional unit/suffix) and
    is ≤15 characters. Dates, long sentences, and plain text are rejected and
    should be demoted into the item's description instead.
    """
    if not value:
        return "", False
    v = str(value).strip()
    if not v:
        return "", False
    if len(v) > 18:
        return v, False
    if _DATE_LIKE_RE.match(v):
        return v, False
    if _METRIC_VALUE_RE.match(v):
        return v, True
    # Short strings with ≥1 digit count as metrics (e.g., "$1B", "3x")
    if len(v) <= 10 and any(ch.isdigit() for ch in v):
        return v, True
    return v, False


def _build_kpi_slide(
    plan: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
    sc: SlideContent | None = None,
) -> SlideSpec:
    """Build a KPI cards slide using AI-written infographic items or raw metrics.

    KPI values are sanitized — items whose *value* is a date or long sentence
    are rewritten so the bad value moves into the description and the card
    shows the item *title* as a fallback value (much better than printing a
    date as the hero metric).
    """
    # AI-written infographic items preferred
    if sc and sc.infographic_items:
        raw_items = sc.infographic_items[:6]
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

        raw_items = []
        for m in metrics[:5]:
            raw_items.append(InfographicItem(
                title=m.label,
                description="",
                value=m.value,
            ))

    # Sanitize — demote bad values into description
    items: list[InfographicItem] = []
    for it in raw_items:
        v_clean, ok = _sanitize_kpi_value(it.value)
        if ok:
            items.append(it)
            continue
        if v_clean:
            new_desc = (it.description or "").strip()
            if v_clean not in new_desc:
                new_desc = f"{v_clean}. {new_desc}".strip(". ") if new_desc else v_clean
            items.append(InfographicItem(
                title=it.title,
                description=new_desc[:config.MAX_INFOGRAPHIC_DESC],
                value=None,
            ))
        else:
            items.append(InfographicItem(
                title=it.title,
                description=it.description,
                value=None,
            ))

    title = sc.title if sc and sc.title else (plan.action_title or plan.title)

    # Count items that still have a valid (non-None) KPI value after sanitization
    valid_value_count = sum(1 for it in items if it.value and it.value.strip())
    step_like_count = sum(1 for it in items if _is_step_like(it.title or ""))

    # If the item titles are "Step 1", "Phase 2", etc. we have a process, not
    # a KPI slide — stat_grid would render meaningless hero numbers. Prefer
    # process_flow so each step is labelled and sequenced correctly.
    if step_like_count >= max(2, len(items) // 2):
        infographic_type = "process_flow"
        final_items = items[:6]
    # If we have 3-4 clean numeric values → use stat_grid (much more visually
    # dominant per the Common-Mistakes note "numbers not dominant enough").
    # If 1-2 clean values → hero_number.
    # Else fall back to icon_list (no empty-value KPI cards).
    elif valid_value_count >= 3:
        infographic_type = "stat_grid"
        final_items = [it for it in items if it.value and it.value.strip()][:4]
    elif valid_value_count >= 1 and len(items) >= 1:
        infographic_type = "hero_number"
        # Sort so the item with a value comes first
        final_items = sorted(items, key=lambda it: 0 if (it.value and it.value.strip()) else 1)[:4]
    else:
        infographic_type = "icon_list"
        final_items = items[:5]

    elements = [
        SlideElement(
            element_type="infographic",
            position=_grid.full(),
            content=InfographicContent(
                infographic_type=infographic_type,
                items=final_items,
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


def _is_meaningful_item(it: InfographicItem) -> bool:
    """Return True iff the item has at least a non-empty, non-punctuation title or description."""
    def _norm(s: str | None) -> str:
        return re.sub(r"[\W_]+", "", (s or "").strip()).strip()
    return bool(_norm(it.title)) or bool(_norm(it.description)) or bool(_norm(it.value))


_STEP_LIKE_RE = re.compile(r"^\s*(step|phase|stage|part|tier|level)\s*[#:]?\s*\d+\s*$", re.IGNORECASE)


def _is_step_like(title: str) -> bool:
    """Titles like 'Step 1', 'Phase 2', 'Stage 3' shouldn't be promoted to hero numbers."""
    return bool(title and _STEP_LIKE_RE.match(title.strip()))


_YEAR_ONLY_RE = re.compile(r"^\s*((?:19|20)\d{2})\s*$")


def _sanitize_timeline_value(item: InfographicItem) -> InfographicItem:
    """Ensure a timeline item's ``value`` is a 4-digit year.

    The LLM occasionally emits dollar amounts or percentages as the timeline
    label (e.g. ``$400K``), which turns a milestone chip into a meaningless
    badge. If ``value`` isn't already a year, extract one from the title /
    description; otherwise clear it.
    """
    v = (item.value or "").strip()
    if _YEAR_ONLY_RE.match(v):
        return item
    year = _extract_year_from_text(v) or _extract_year_from_text(item.title or "") or _extract_year_from_text(item.description or "")
    return InfographicItem(
        title=item.title,
        description=item.description,
        value=year,
        icon=item.icon,
    )


def _build_infographic_slide(plan: SlidePlanItem, sections: list[ContentSection], sc: SlideContent | None = None) -> SlideSpec:
    """Build an infographic slide using AI-written items.

    Filters out LLM-emitted blank / dash-only items — they render as empty
    ``- \\x0b -`` boxes which is the exact defect flagged in QA.
    """
    infographic_type = plan.infographic_type_hint or "process_flow"

    # AI-written infographic items preferred
    if sc and sc.infographic_items:
        items = [it for it in sc.infographic_items if _is_meaningful_item(it)][:6]
    else:
        items = []

    # Fall back to section-derived items if LLM items were all empty
    if not items:
        is_timeline = infographic_type == "timeline"
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

    # Fall back to bullets if we still have no usable items — that path has
    # its own emergency pull_quote fallback to prevent bare title slides.
    if not items:
        return _build_bullets_slide(plan, sections, sc)

    # Hierarchy is rarely visually useful — redirect to icon_list for ≤6 items
    if infographic_type == "hierarchy":
        infographic_type = "icon_list"

    # Timeline values must be years — sanitize each item's ``value``
    if infographic_type == "timeline":
        items = [_sanitize_timeline_value(it) for it in items]

    # Auto-promote to stat_grid when most items have clean numeric values
    # (matches the "numbers aren't visually dominant" defect note).
    # BUT: reject promotion if titles are "Step 1", "Phase 2", etc. — those
    # are sequence labels, not statistics, and read badly as giant numbers.
    if infographic_type not in ("timeline", "stat_grid", "hero_number", "pull_quote"):
        step_like = sum(1 for it in items if _is_step_like(it.title or ""))
        if step_like >= max(2, len(items) // 2):
            # Looks like a process: prefer process_flow archetype
            infographic_type = "process_flow"
        else:
            with_values = sum(1 for it in items if it.value and it.value.strip())
            if with_values >= 3 and with_values >= len(items) // 2:
                infographic_type = "stat_grid"

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
                # Use full pipeline (dedup + flat filter + smart_label) instead of raw row[0]
                categories, series_list = _extract_chart_data(t)
                if series_list and len(categories) >= 2:
                    pos_left, pos_right = _grid.two_column()
                    chart_title = sc.chart_insight if sc and sc.chart_insight else ""
                    use_log = _needs_log_scale(series_list) and chart_type in ("column", "line", "area")
                    chart_elem = SlideElement(
                        element_type="chart",
                        position=pos_left,
                        content=ChartContent(
                            chart_type=chart_type,
                            title=chart_title,
                            categories=categories,
                            series=series_list,
                            log_scale=use_log,
                        ),
                    )
                    break
        if chart_elem:
            break

    # If no chart, try infographic on left using AI items — but only accept
    # items that actually carry meaningful content; the LLM sometimes emits
    # placeholder items whose title/description collapse to the same text as
    # the adjacent bullet, yielding two identical columns.
    if not chart_elem:
        inf_items: list[InfographicItem] = []
        if sc and sc.infographic_items:
            inf_items = [it for it in sc.infographic_items if _is_meaningful_item(it)][:4]
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
            inf_items = [it for it in inf_items if _is_meaningful_item(it)]

        # Require substantively different content from the bullets so we don't
        # paint the same message twice on one slide.
        bullet_titles = {(b or "").strip().lower()[:40] for b in bullets}
        inf_titles = {(it.title or "").strip().lower()[:40] for it in inf_items}
        overlap = bullet_titles & inf_titles
        if inf_items and len(overlap) >= len(inf_items) - 1:
            inf_items = []  # too much duplication — just show bullets

        if len(inf_items) >= 2:
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

    # When no chart / left-side visual is available, redirect to the richer
    # bullets builder — it routes to icon_list / pull_quote / comparison
    # archetypes which fill the full slide width instead of parking the
    # bullets in a single narrow column with nothing on the other side.
    if not chart_elem and bullets:
        return _build_bullets_slide(plan, sections, sc)

    if chart_elem:
        elements.append(chart_elem)

    if bullets:
        elements.append(SlideElement(
            element_type="bullets",
            position=_grid.two_column()[1],
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
    """Build conclusion / key takeaways slide using AI-written synthesis.

    Filters blank / punctuation-only bullets to prevent empty rendering, and
    surfaces the takeaways as an ``icon_list`` for visual consistency with
    the rest of the deck.
    """
    def _clean(lst: list[str]) -> list[str]:
        out = []
        for b in lst or []:
            s = (b or "").strip()
            if s and re.sub(r"[\W_]+", "", s):
                out.append(s)
        return out

    # AI-written bullets preferred — conclusion should be AI-synthesized from entire deck
    bullets = _clean(sc.bullets if sc else []) if sc else []
    bullets = bullets[:config.MAX_BULLETS_PER_SLIDE]

    # Fallback: derive from source sections or from earlier slides' key_messages
    if not bullets:
        fallback: list[str] = []
        for sec in sections or []:
            fallback.extend(sec.bullets or [])
            if sec.text:
                fallback.extend(
                    s.strip() for s in re.split(r"(?<=[.!?])\s+", sec.text) if s.strip()
                )
        # Also pull from earlier slide key_messages as last resort
        if slide_plan and slide_plan.slides:
            for sp in slide_plan.slides:
                if sp.key_message:
                    fallback.append(sp.key_message)
        bullets = _clean(fallback)[:config.MAX_BULLETS_PER_SLIDE]

    title = sc.title if sc and sc.title else (plan.action_title or plan.title or "Key Takeaways")
    subtitle = sc.subtitle if sc and sc.subtitle else plan.subtitle

    elements: list[SlideElement] = []

    if bullets:
        # Prefer icon_list for 3-6 bullets (matches deck style), fall back to plain bullets
        if 3 <= len(bullets) <= 6:
            items = _bullets_to_infographic_items(bullets)
            if items:
                elements.append(SlideElement(
                    element_type="infographic",
                    position=_grid.full(),
                    content=InfographicContent(
                        infographic_type="icon_list",
                        items=items,
                    ),
                ))
        if not elements:
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
        subtitle=subtitle,
        elements=elements,
    )
