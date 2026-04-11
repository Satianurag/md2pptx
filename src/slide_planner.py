from __future__ import annotations
import logging
from .schemas import ContentTree, ContentSection, SlideMasterInfo, SlidePlan
from .content_profiler import ContentProfile
from .llm import invoke_llm_structured
from . import config

logger = logging.getLogger(__name__)


PLANNER_SYSTEM_PROMPT = """\
You are an expert presentation architect. Your job is to take structured content from a parsed \
markdown research report and produce a slide plan — a JSON outline that decides the storyline, \
slide types, and how content maps to slides.

STRICT RULES:
1. Total slides MUST be between {min_slides} and {max_slides}.
2. Slide flow MUST follow this order:
   - Slide 1: cover (title slide)
   - Slide 2: agenda (table of contents / overview)
   - Slide 3: executive_summary
   - Slides 4–N-2: section content (content, chart, table, infographic, mixed, section_divider)
   - Slide N-1: conclusion (key takeaways)
   - Slide N: thank_you
3. Every major section from the content MUST appear in at least one slide. Do NOT drop sections.
4. Each slide should have ONE key message — avoid cramming multiple topics.

INFOGRAPHIC-FIRST APPROACH (CRITICAL — this is the most important rule):
5. NEVER default to plain "bullets". Always try to visualize content first:
   - Tables with numeric data → "chart" (ALWAYS prefer chart over showing raw table)
   - 3-6 metrics/KPIs → "kpi" (rendered as bold metric cards)
   - Steps, stages, methodology, process, pipeline, workflow → "infographic" with process_flow
   - Chronological events, milestones, history → "infographic" with timeline
   - Comparing 2-4 items, pros/cons, options, alternatives → "infographic" with comparison
   - Hierarchies, org structures, categories → "infographic" with hierarchy
   - Mixed numeric + text → "mixed" (chart + bullets side-by-side)
   - ONLY use "bullets" as LAST RESORT when content is purely textual with no structure
   - At LEAST 50% of content slides (slides 4 through N-2) MUST be chart, infographic, or mixed.
6. When visualization_hint is "chart", you MUST set chart_type_hint to one of: bar, column, line, pie, area, doughnut.
7. When visualization_hint is "infographic", you MUST set infographic_type_hint to one of: process_flow, timeline, comparison, kpi_cards, hierarchy.
8. Use section_divider slides to separate major topic groups (use sparingly, 1-2 max).
9. content_source must contain the EXACT heading text from the provided sections.
10. key_message should be a single sentence summarizing the slide's takeaway.
11. TABLE RULE: At least 1 slide MUST use visualization_hint "table" when the content contains tables \
with mostly text columns, feature comparisons, or categorical data that is better read in tabular form. \
Not every table should become a chart — keep text-heavy tables as tables.
"""


def _condense_content_tree(tree: ContentTree) -> str:
    """Create a condensed text representation of the ContentTree for the LLM prompt."""
    parts = []
    parts.append(f"TITLE: {tree.title}")
    if tree.subtitle:
        parts.append(f"SUBTITLE: {tree.subtitle}")
    if tree.executive_summary:
        parts.append(f"EXECUTIVE SUMMARY: {tree.executive_summary[:500]}")

    parts.append(f"\nTOTAL TABLES: {len(tree.all_tables)}")
    parts.append(f"TOTAL METRICS: {len(tree.all_metrics)}")

    parts.append("\nSECTIONS:")
    for section in tree.sections:
        _append_section(parts, section, depth=0)

    result = "\n".join(parts)
    # Cap to ~3000 chars to keep token usage low
    if len(result) > 6000:
        result = result[:6000] + "\n... [truncated]"
    return result


def _append_section(parts: list[str], section: ContentSection, depth: int) -> None:
    indent = "  " * depth
    has_tables = "YES" if section.tables else "no"
    has_metrics = "YES" if section.metrics else "no"
    bullet_count = len(section.bullets)
    text_len = len(section.text)

    parts.append(
        f"{indent}- [{section.level}] \"{section.heading}\" "
        f"(text:{text_len}ch, bullets:{bullet_count}, tables:{has_tables}, metrics:{has_metrics})"
    )

    # Show bullet previews (first 3)
    for b in section.bullets[:3]:
        preview = b[:80] + "..." if len(b) > 80 else b
        parts.append(f"{indent}    • {preview}")
    if bullet_count > 3:
        parts.append(f"{indent}    ... +{bullet_count - 3} more bullets")

    # Show table info
    for t in section.tables[:2]:
        cols = ", ".join(t.headers[:5])
        parts.append(f"{indent}    📊 Table: [{cols}] ({len(t.rows)} rows)")

    # Show metrics
    for m in section.metrics[:3]:
        parts.append(f"{indent}    📈 Metric: {m.label} = {m.value}")

    for sub in section.subsections:
        _append_section(parts, sub, depth + 1)


def _build_profile_context(profile: ContentProfile | None) -> str:
    """Build a prompt section from the content profile."""
    if profile is None:
        return ""
    lines = [
        "\nCONTENT PROFILE (use this to guide visualization choices):",
        f"  Archetype: {profile.archetype}",
        f"  Data richness: {profile.data_richness}",
        f"  Recommended visual ratio: {profile.recommended_visual_ratio:.0%} of content slides should be visual",
        f"  Best chart types for this content: {', '.join(profile.recommended_chart_types)}",
        f"  Best infographic types: {', '.join(profile.recommended_infographic_types)}",
        f"  Total tables: {profile.total_tables}, Total metrics: {profile.total_metrics}",
    ]
    if profile.best_tables:
        lines.append(f"  Top chart-worthy tables: {len(profile.best_tables)} ranked")
    if profile.best_metrics:
        top_m = profile.best_metrics[:5]
        lines.append(f"  Top KPI metrics: {', '.join(m.metric.label + '=' + m.metric.value for m in top_m)}")
    if profile.sections_by_value:
        lines.append(f"  Sections by data value (highest first): {', '.join(profile.sections_by_value[:6])}")
    return "\n".join(lines)


def plan_slides(
    content_tree: ContentTree,
    master_info: SlideMasterInfo | None = None,
    target_slide_count: int = 12,
    content_profile: ContentProfile | None = None,
) -> SlidePlan:
    """Use the LLM to generate a SlidePlan from the ContentTree."""
    min_slides = config.MIN_SLIDES
    max_slides = config.MAX_SLIDES

    # Clamp target
    target_slide_count = max(min_slides, min(max_slides, target_slide_count))

    system_prompt = PLANNER_SYSTEM_PROMPT.format(
        min_slides=min_slides,
        max_slides=max_slides,
    )

    condensed = _condense_content_tree(content_tree)
    profile_context = _build_profile_context(content_profile)

    # Build available layouts info
    layouts_info = ""
    if master_info:
        layout_names = [l.name for l in master_info.layouts]
        layouts_info = f"\nAVAILABLE TEMPLATE LAYOUTS: {', '.join(layout_names)}"

    user_prompt = f"""\
Create a slide plan for the following content. Target {target_slide_count} slides (must be {min_slides}-{max_slides}).

{condensed}
{layouts_info}
{profile_context}

Generate the SlidePlan JSON with storyline_summary, target_slide_count, and the slides array.\
"""

    logger.info(f"Planning slides: target={target_slide_count}, sections={len(content_tree.sections)}")

    plan = invoke_llm_structured(
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        output_schema=SlidePlan,
        estimated_tokens=5000,
    )

    # Post-process: enforce slide count bounds
    if len(plan.slides) < min_slides:
        logger.warning(f"Plan has {len(plan.slides)} slides, below minimum {min_slides}")
    if len(plan.slides) > max_slides:
        logger.warning(f"Plan has {len(plan.slides)} slides, above maximum {max_slides}. Trimming.")
        plan.slides = plan.slides[:max_slides]

    # Renumber slides
    for i, slide in enumerate(plan.slides):
        slide.slide_number = i + 1

    plan.target_slide_count = len(plan.slides)

    logger.info(f"Slide plan generated: {len(plan.slides)} slides")
    return plan
