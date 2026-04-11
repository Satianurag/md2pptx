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

# Max chars to keep per section for LLM processing
MAX_SECTION_TEXT_CHARS = 500
MAX_BULLET_CHARS = 250
MAX_BULLETS_PER_SECTION = 8
MAX_TABLE_ROWS = 10
MAX_TOTAL_CHARS = 100_000  # ~25k tokens


def chunk_content_tree(tree: ContentTree, content_profile=None) -> ContentTree:
    """Reduce a large ContentTree to fit within token limits while preserving structure.

    Uses tiered chunking: lighter for <5MB, more aggressive for >5MB.
    When *content_profile* is available, prioritise high-value sections.
    """
    total_chars = _estimate_chars(tree)

    if total_chars <= MAX_TOTAL_CHARS:
        logger.info(f"Content size OK: {total_chars} chars")
        return tree

    # Determine aggressiveness tier
    tier = "standard"
    if total_chars > 500_000:  # ~125k tokens → very large file
        tier = "aggressive"
        limit = 60_000
    elif total_chars > 200_000:  # ~50k tokens → large file
        tier = "moderate"
        limit = 80_000
    else:
        limit = MAX_TOTAL_CHARS

    logger.info(f"Content {total_chars} chars, tier={tier}, target={limit}")

    # Truncate executive summary
    summary_max = 500 if tier == "aggressive" else 1000
    if len(tree.executive_summary) > summary_max:
        tree.executive_summary = tree.executive_summary[:summary_max]

    # For aggressive tier, limit sections count first
    max_sections = 15 if tier == "standard" else (10 if tier == "moderate" else 8)
    if len(tree.sections) > max_sections:
        # Use profile's section ranking to keep high-value sections
        if content_profile and hasattr(content_profile, 'sections_by_value') and content_profile.sections_by_value:
            value_order = {h.lower().strip(): i for i, h in enumerate(content_profile.sections_by_value)}
            ranked = sorted(tree.sections, key=lambda s: value_order.get(s.heading.lower().strip(), 999))
            tree.sections = ranked[:max_sections]
            logger.info(f"Profile-aware chunking: kept {max_sections} highest-value sections")
        else:
            tree.sections = tree.sections[:max_sections]

    # Truncate each section (with tier-aware limits)
    sec_text = 300 if tier == "aggressive" else (400 if tier == "moderate" else MAX_SECTION_TEXT_CHARS)
    bul_limit = 5 if tier == "aggressive" else (6 if tier == "moderate" else MAX_BULLETS_PER_SECTION)
    tbl_rows = 6 if tier == "aggressive" else (8 if tier == "moderate" else MAX_TABLE_ROWS)

    for section in tree.sections:
        _truncate_section(section, sec_text, bul_limit, tbl_rows)

    # Limit total tables and metrics
    max_tables = 8 if tier == "aggressive" else (12 if tier == "moderate" else 15)
    tree.all_tables = tree.all_tables[:max_tables]
    for t in tree.all_tables:
        t.rows = t.rows[:tbl_rows]
    max_metrics = 10 if tier == "aggressive" else (15 if tier == "moderate" else 20)
    tree.all_metrics = tree.all_metrics[:max_metrics]

    new_total = _estimate_chars(tree)
    logger.info(f"After {tier} chunking: {new_total} chars")

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
