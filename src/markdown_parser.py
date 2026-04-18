"""Markdown parser with two-pass support for large files.

Key design decisions (from research §4.1, §5.2, §19):
- Two-pass parsing for files >5 MB: first pass extracts text only (~250 KB),
  images are lazy-loaded on demand.
- Skip-section detection at ALL heading levels (not just H2): TOC, References,
  Appendix, Citations, Bibliography, Acknowledgements, etc.
- Hyperlinks are stripped, keeping display text only.
- Nested lists are flattened to max 2 levels for slide readability.
- Paragraph-heavy content is converted to bullet points.
"""
from __future__ import annotations
import re
import logging
import mistune
from typing import Optional
from .schemas import ContentTree, ContentSection, DataTable, KeyMetric

logger = logging.getLogger(__name__)

# ── Skip-section patterns (research §5.2) ────────────────────────────
# These sections are excluded from slide content at ANY heading level.
_SKIP_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"^table\s+of\s+contents?$", re.IGNORECASE),
    re.compile(r"^(toc|contents?)$", re.IGNORECASE),
    re.compile(r"^references?$", re.IGNORECASE),
    re.compile(r"^bibliography$", re.IGNORECASE),
    re.compile(r"^citations?$", re.IGNORECASE),
    re.compile(r"^appendix", re.IGNORECASE),
    re.compile(r"^acknowledg[e]?ments?$", re.IGNORECASE),
    re.compile(r"^about\s+the\s+author", re.IGNORECASE),
    re.compile(r"^disclaimer$", re.IGNORECASE),
    re.compile(r"^glossary$", re.IGNORECASE),
    re.compile(r"^footnotes?$", re.IGNORECASE),
    re.compile(r"^endnotes?$", re.IGNORECASE),
    re.compile(r"^sources?$", re.IGNORECASE),
    re.compile(r"^works?\s+cited$", re.IGNORECASE),
    re.compile(r"^further\s+reading$", re.IGNORECASE),
    re.compile(r"^related\s+(articles?|resources?)$", re.IGNORECASE),
]

# Base64 image pattern — used for two-pass stripping on large files.
_BASE64_IMG_RE = re.compile(
    r"!\[[^\]]*\]\(data:image/[^;]+;base64,[A-Za-z0-9+/=\s]+\)",
    re.DOTALL,
)

# Markdown link pattern — [text](url) → text
_MD_LINK_RE = re.compile(r"\[([^\]]+)\]\([^)]+\)")

# Threshold for two-pass parsing (bytes).
_LARGE_FILE_THRESHOLD = 5 * 1024 * 1024  # 5 MB


def _is_skip_section(heading: str) -> bool:
    """Return True if *heading* matches a skip-section pattern."""
    heading_clean = heading.strip()
    return any(pat.match(heading_clean) for pat in _SKIP_PATTERNS)


def _strip_hyperlinks(md_text: str) -> str:
    """Replace markdown links [text](url) with just the display text."""
    return _MD_LINK_RE.sub(r"\1", md_text)


def _strip_base64_images(md_text: str) -> str:
    """Remove base64-encoded images from markdown, leaving a placeholder."""
    return _BASE64_IMG_RE.sub("[image]", md_text)


def preprocess_markdown(md_text: str) -> str:
    """Pre-process raw markdown before AST parsing.

    For files > 5 MB (research §19: Clinical Trial 25.2 MB, Used Taxi 19.2 MB),
    strip base64 images first to reduce to ~250 KB of text.  Images can be
    lazy-loaded later from the original text if needed.

    Always strips hyperlinks (keeping display text).
    """
    text_size = len(md_text.encode("utf-8", errors="replace"))
    if text_size > _LARGE_FILE_THRESHOLD:
        logger.info(
            f"Large file detected ({text_size / 1024 / 1024:.1f} MB), "
            f"stripping base64 images for first-pass parsing"
        )
        md_text = _strip_base64_images(md_text)
        stripped_size = len(md_text.encode("utf-8", errors="replace"))
        logger.info(f"After base64 strip: {stripped_size / 1024:.0f} KB")

    md_text = _strip_hyperlinks(md_text)
    return md_text


def parse_markdown(md_text: str) -> ContentTree:
    """Parse markdown text into a structured ContentTree.

    Applies pre-processing (hyperlink stripping, base64 stripping for large
    files) before AST parsing.  Skip sections (TOC, References, Appendix,
    etc.) are detected at all heading levels and excluded.
    """
    md_text = preprocess_markdown(md_text)
    md = mistune.create_markdown(renderer='ast', plugins=['table', 'strikethrough'])
    tokens = md(md_text)

    title = ""
    subtitle = ""
    executive_summary = ""
    sections: list[ContentSection] = []
    all_tables: list[DataTable] = []
    all_metrics: list[KeyMetric] = []

    current_h1: Optional[ContentSection] = None
    current_h2: Optional[ContentSection] = None
    current_h3: Optional[ContentSection] = None
    skip_until_level: int | None = None  # when inside a skip section, ignore until next heading at this level or higher

    for token in tokens:
        tok_type = token.get('type', '')

        # ── Headings ──
        if tok_type == 'heading':
            level = token.get('attrs', {}).get('level', 1) if 'attrs' in token else token.get('level', 1)
            heading_text = _extract_text(token.get('children', []))

            # End skip zone if we've reached a heading at the same or higher level
            if skip_until_level is not None and level <= skip_until_level:
                skip_until_level = None

            # Skip-section detection at ALL heading levels (GAP-1)
            if _is_skip_section(heading_text):
                logger.debug(f"Skipping section: '{heading_text}' (H{level})")
                skip_until_level = level
                continue

            # If inside a skip zone, ignore everything
            if skip_until_level is not None:
                continue

            if level == 1 and not title:
                title = heading_text
                continue

            # A level-3 heading immediately after the H1 title (before any H2)
            # is conventionally the deck's tagline / subtitle.
            if (
                level == 3 and title and not subtitle
                and current_h1 is None and current_h2 is None
                and len(heading_text) < 220
            ):
                subtitle = heading_text
                continue

            if level == 2 and heading_text.lower().startswith('executive summary'):
                current_h1 = ContentSection(heading="Executive Summary", level=2)
                sections.append(current_h1)
                current_h2 = None
                current_h3 = None
                continue

            section = ContentSection(heading=heading_text, level=level)

            if level <= 2:
                sections.append(section)
                current_h1 = section
                current_h2 = None
                current_h3 = None
            elif level == 3 and current_h1:
                current_h1.subsections.append(section)
                current_h2 = section
                current_h3 = None
            elif level >= 4 and current_h2:
                current_h2.subsections.append(section)
                current_h3 = section
            elif level >= 4 and current_h1:
                current_h1.subsections.append(section)
                current_h3 = section
            else:
                sections.append(section)
                current_h1 = section

        # If inside a skip zone, ignore all content tokens
        if skip_until_level is not None:
            continue

        # ── Paragraphs ──
        if tok_type == 'paragraph':
            text = _extract_text(token.get('children', []))
            target = current_h3 or current_h2 or current_h1
            if target:
                if target.heading.lower().startswith('executive summary') and not executive_summary:
                    executive_summary = text
                else:
                    target.text += ("\n" if target.text else "") + text
                # Extract metrics from paragraph
                metrics = _extract_metrics(text)
                if metrics:
                    target.metrics.extend(metrics)
                    all_metrics.extend(metrics)

        # ── Lists ──
        elif tok_type == 'list':
            items = _extract_list_items(token)
            target = current_h3 or current_h2 or current_h1
            if target:
                target.bullets.extend(items)

        # ── Tables ──
        elif tok_type == 'table':
            table = _parse_table_token(token)
            if table:
                all_tables.append(table)
                target = current_h3 or current_h2 or current_h1
                if target:
                    target.tables.append(table)
                    # Extract metrics from table
                    metrics = _extract_metrics_from_table(table)
                    if metrics:
                        target.metrics.extend(metrics)
                        all_metrics.extend(metrics)

        # ── Block quote ──
        elif tok_type == 'block_quote':
            text = _extract_block_text(token)
            target = current_h3 or current_h2 or current_h1
            if target:
                target.text += ("\n" if target.text else "") + text

        # ── Code blocks ──
        elif tok_type == 'block_code':
            code_text = token.get('text', '')
            target = current_h3 or current_h2 or current_h1
            if target:
                target.code_blocks.append(code_text)

    # If subtitle wasn't found, try to get it from first paragraph-like content after title
    if not subtitle and sections:
        first = sections[0]
        if first.text and len(first.text) < 200:
            subtitle = first.text

    return ContentTree(
        title=title,
        subtitle=subtitle,
        sections=sections,
        executive_summary=executive_summary,
        all_tables=all_tables,
        all_metrics=all_metrics,
    )


def _extract_text(children: list) -> str:
    """Recursively extract plain text from AST children."""
    parts = []
    for child in children:
        if isinstance(child, str):
            parts.append(child)
        elif isinstance(child, dict):
            raw = child.get('raw', '')
            if raw:
                parts.append(raw)
            text = child.get('text', '')
            if text:
                parts.append(text)
            nested = child.get('children', [])
            if nested:
                parts.append(_extract_text(nested))
    return ''.join(parts).strip()


def _extract_list_items(token: dict, depth: int = 0, max_depth: int = 2) -> list[str]:
    """Extract list items, flattening nested lists to *max_depth* levels.

    Research §15.3 / §19: some test files have 500+ nested items.
    Slides can only show ~2 levels, so deeper nesting is flattened.
    """
    items: list[str] = []
    for child in token.get('children', []):
        if child.get('type') == 'list_item':
            text_parts: list[str] = []
            for sub in child.get('children', []):
                if sub.get('type') == 'paragraph':
                    text_parts.append(_extract_text(sub.get('children', [])))
                elif sub.get('type') == 'list':
                    if depth < max_depth:
                        nested = _extract_list_items(sub, depth + 1, max_depth)
                        text_parts.extend(f"  - {n}" for n in nested)
                    else:
                        # Flatten: just extract text without nesting prefix
                        flat = _extract_list_items(sub, depth + 1, max_depth)
                        text_parts.extend(flat)
            combined = '\n'.join(text_parts)
            if combined.strip():
                items.append(combined)
    return items


def _parse_table_token(token: dict) -> Optional[DataTable]:
    """Parse a table AST token into DataTable."""
    children = token.get('children', [])
    if not children:
        return None

    headers = []
    rows = []
    alignments = []

    for child in children:
        child_type = child.get('type', '')
        if child_type == 'table_head':
            for cell in child.get('children', []):
                if cell.get('type') == 'table_cell':
                    headers.append(_extract_text(cell.get('children', [])))
                    attrs = cell.get('attrs', {})
                    alignments.append(attrs.get('align') or attrs.get('style'))
        elif child_type == 'table_body':
            for row in child.get('children', []):
                if row.get('type') == 'table_row':
                    row_data = []
                    for cell in row.get('children', []):
                        if cell.get('type') == 'table_cell':
                            row_data.append(_extract_text(cell.get('children', [])))
                    rows.append(row_data)

    if not headers and not rows:
        return None

    # Try to find a title from the first header
    title = None
    if headers and any(h.strip() for h in headers):
        pass  # title will be set by the agent based on context

    return DataTable(title=title, headers=headers, rows=rows, alignments=alignments)


def _extract_block_text(token: dict) -> str:
    """Extract text from a block_quote token."""
    parts = []
    for child in token.get('children', []):
        if child.get('type') == 'paragraph':
            parts.append(_extract_text(child.get('children', [])))
    return '\n'.join(parts)


def paragraphs_to_bullets(text: str, max_bullets: int = 6) -> list[str]:
    """Convert prose-heavy paragraph text into concise bullet points.

    Research §19 / Guidelines §14.3: "Slides feel like documents, not slides"
    is a common judge complaint.  This converts paragraphs into bullets by
    splitting on sentence boundaries and keeping the most informative ones.

    Returns at most *max_bullets* items.
    """
    if not text or not text.strip():
        return []

    # Split on sentence boundaries (period + space/newline, or newline-newline)
    sentences = re.split(r'(?<=[.!?])\s+|\n{2,}', text.strip())
    sentences = [s.strip() for s in sentences if s.strip() and len(s.strip()) > 10]

    if not sentences:
        return []

    # Prioritise sentences with numbers/percentages (research §5.4 Priority 1)
    _NUM_RE = re.compile(r'\d+[%$€£BMKTbmkt]|\$\d|\d+\.\d')

    def _score(s: str) -> float:
        score = 0.0
        if _NUM_RE.search(s):
            score += 2.0
        if any(kw in s.lower() for kw in ('key', 'important', 'significant', 'critical', 'major')):
            score += 1.0
        # Shorter sentences are preferred for bullets
        word_count = len(s.split())
        if word_count <= 15:
            score += 0.5
        return score

    ranked = sorted(sentences, key=_score, reverse=True)
    return ranked[:max_bullets]


def _extract_metrics(text: str) -> list[KeyMetric]:
    """Extract key metrics (numbers with context) from text."""
    metrics = []
    # Pattern: $X.X billion/million/thousand, X%, Xk, etc.
    patterns = [
        r'[\$€£]?\s*([\d,]+\.?\d*)\s*(billion|million|trillion|thousand|B|M|T|K)\b',
        r'(\d+\.?\d*)\s*%',
        r'(\d+\.?\d*)[xX]\s',
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            start = max(0, match.start() - 50)
            end = min(len(text), match.end() + 30)
            context = text[start:end].strip()
            # Get a short label from context
            label_match = re.search(r'([A-Z][a-z]+(?:\s+[a-z]+){0,4})', context)
            label = label_match.group(1) if label_match else context[:40]
            value = match.group(0).strip()
            unit = match.group(1) if len(match.groups()) > 0 else None
            metrics.append(KeyMetric(label=label, value=value, unit=unit))

    return metrics


def _extract_metrics_from_table(table: DataTable) -> list[KeyMetric]:
    """Extract key metrics from a DataTable."""
    metrics = []
    num_pattern = re.compile(r'^[\$€£]?\s*[\d,]+\.?\d*\s*[%BMKTbmkt]?')

    for row in table.rows:
        for i, cell in enumerate(row):
            if num_pattern.match(cell.strip()):
                label = table.headers[i] if i < len(table.headers) else ""
                if label and row[0] != cell:
                    label = f"{row[0]} - {label}"
                metrics.append(KeyMetric(label=label, value=cell.strip()))

    return metrics[:10]  # cap to avoid flooding
