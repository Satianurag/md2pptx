from __future__ import annotations
import re
import mistune
from typing import Optional
from .schemas import ContentTree, ContentSection, DataTable, KeyMetric


def parse_markdown(md_text: str) -> ContentTree:
    """Parse markdown text into a structured ContentTree."""
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

    for token in tokens:
        tok_type = token.get('type', '')

        # ── Headings ──
        if tok_type == 'heading':
            level = token.get('attrs', {}).get('level', 1) if 'attrs' in token else token.get('level', 1)
            heading_text = _extract_text(token.get('children', []))

            if level == 1 and not title:
                title = heading_text
                continue

            if level == 2 and heading_text.lower().startswith('executive summary'):
                # Next paragraph(s) will be the summary - mark flag
                current_h1 = ContentSection(heading="Executive Summary", level=2)
                sections.append(current_h1)
                current_h2 = None
                current_h3 = None
                continue

            if level == 2 and heading_text.lower().startswith('table of contents'):
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

        # ── Paragraphs ──
        elif tok_type == 'paragraph':
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


def _extract_list_items(token: dict) -> list[str]:
    """Extract list items from a list token."""
    items = []
    for child in token.get('children', []):
        if child.get('type') == 'list_item':
            text_parts = []
            for sub in child.get('children', []):
                if sub.get('type') == 'paragraph':
                    text_parts.append(_extract_text(sub.get('children', [])))
                elif sub.get('type') == 'list':
                    nested = _extract_list_items(sub)
                    text_parts.extend(f"  - {n}" for n in nested)
            items.append('\n'.join(text_parts))
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
