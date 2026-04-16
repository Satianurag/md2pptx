from __future__ import annotations
import re
import logging
from typing import Optional
from .schemas import ContentTree, ContentSection

logger = logging.getLogger(__name__)

# Late import to avoid circular dependency
def _get_profile_cls():
    from .content_profiler import ContentProfile
    return ContentProfile

# Max chars to keep — respect model limits only (~1M tokens for Gemini 2.5 Flash)
# Only chunk truly massive inputs that would exceed the model's context window.
MAX_SECTION_TEXT_CHARS = 5000
MAX_BULLET_CHARS = 500
MAX_BULLETS_PER_SECTION = 30
MAX_TABLE_ROWS = 50
MAX_TOTAL_CHARS = 2_000_000  # ~500k tokens — well within Gemini's 1M limit


def chunk_content_tree(tree: ContentTree, content_profile=None) -> ContentTree:
    """Reduce a large ContentTree to fit within token limits while preserving structure.

    Uses tiered chunking: lighter for <5MB, more aggressive for >5MB.
    When *content_profile* is available, prioritise high-value sections.
    """
    total_chars = _estimate_chars(tree)

    if total_chars <= MAX_TOTAL_CHARS:
        logger.info(f"Content size OK: {total_chars} chars — no chunking needed")
        return tree

    # Only chunk if content exceeds model limits
    logger.info(f"Content {total_chars} chars exceeds {MAX_TOTAL_CHARS} limit, chunking")

    for section in tree.sections:
        _truncate_section(section, MAX_SECTION_TEXT_CHARS, MAX_BULLETS_PER_SECTION, MAX_TABLE_ROWS)

    new_total = _estimate_chars(tree)
    logger.info(f"After chunking: {new_total} chars")

    return tree


def _truncate_section(section: ContentSection,
                      max_text: int = MAX_SECTION_TEXT_CHARS,
                      max_bullets: int = MAX_BULLETS_PER_SECTION,
                      max_rows: int = MAX_TABLE_ROWS) -> None:
    """Truncate a single section's content."""
    if len(section.text) > max_text:
        section.text = section.text[:max_text]

    if len(section.bullets) > max_bullets:
        section.bullets = section.bullets[:max_bullets]
    section.bullets = [b[:MAX_BULLET_CHARS] for b in section.bullets]

    section.tables = section.tables[:3]
    for t in section.tables:
        t.rows = t.rows[:max_rows]
        t.headers = t.headers[:8]

    section.metrics = section.metrics[:5]

    section.subsections = section.subsections[:6]
    for sub in section.subsections:
        _truncate_section(sub, max_text, max_bullets, max_rows)


def _estimate_chars(tree: ContentTree) -> int:
    """Rough estimate of total character count."""
    total = len(tree.title) + len(tree.subtitle) + len(tree.executive_summary)
    for section in tree.sections:
        total += _section_chars(section)
    for t in tree.all_tables:
        total += sum(len(h) for h in t.headers)
        total += sum(sum(len(c) for c in row) for row in t.rows)
    for m in tree.all_metrics:
        total += len(m.label) + len(m.value)
    return total


def _section_chars(section: ContentSection) -> int:
    total = len(section.heading) + len(section.text)
    total += sum(len(b) for b in section.bullets)
    for t in section.tables:
        total += sum(len(h) for h in t.headers)
        total += sum(sum(len(c) for c in row) for row in t.rows)
    for m in section.metrics:
        total += len(m.label) + len(m.value)
    for sub in section.subsections:
        total += _section_chars(sub)
    return total
