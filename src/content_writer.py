from __future__ import annotations
import logging
from typing import Optional

from .schemas import (
    ContentTree, ContentSection, SlidePlan, SlidePlanItem,
    DeckContent, SlideContent, InfographicItem, DataTable, KeyMetric,
)
from .llm import invoke_llm_structured

logger = logging.getLogger(__name__)

# Optional import — used only if available
try:
    from .content_profiler import ContentProfile
except ImportError:
    ContentProfile = None  # type: ignore


# ── System prompt ────────────────────────────────────────────────────

CONTENT_WRITER_SYSTEM_PROMPT = """\
You are an expert presentation content writer. Your job is to transform raw research \
report content into polished, presentation-ready slide content.

You will receive:
1. A slide plan with slide types, narrative roles, and visualization hints.
2. The full source content (sections, bullets, tables, metrics) mapped to each slide.

For EACH slide in the plan, produce:
- **title**: An action-oriented, data-driven title that conveys the key takeaway. \
  NOT a generic heading. Example: "ROE Gap Widened to 8.2%% — FAB Leads" instead of "ROE Analysis".
- **subtitle**: Brief contextual subtitle (1 line, optional).
- **key_takeaway**: A single sentence summarizing the most important point of this slide.
- **bullets**: 4-6 concise, insight-driven bullet points. Each bullet MUST convey a \
  specific insight — never copy raw paragraphs. Max ~25 words per bullet.
- **chart_insight**: If the slide has a chart, write an insight title for the chart \
  (e.g., "Market share grew 3x in 2024-2025"). Leave empty if no chart.
- **infographic_items**: If the slide is an infographic (process_flow, timeline, comparison, \
  kpi_cards, hierarchy), provide structured items with title, description, and value.
- **table_summary**: If the slide has a table, write a 1-line context sentence \
  (e.g., "Comparison of key financial metrics across top 5 banks"). Leave empty if no table.
- **speaker_notes**: 2-3 sentences of what the presenter should say for this slide.

CRITICAL RULES:
1. ONLY use facts from the provided source sections. NEVER invent numbers, statistics, \
   percentages, or claims not present in the source.
2. Every bullet must convey a specific, verifiable insight from the source material.
3. Transform verbose report language into punchy, executive-friendly slide language.
4. For cover/thank_you slides: use the report title and a compelling subtitle.
5. For agenda slides: list the key topics in narrative order.
6. For executive_summary: synthesize the ENTIRE report into 4-5 executive-ready bullets.
7. For conclusion: synthesize key takeaways from the ENTIRE deck, not just one section.
8. For chart slides: the chart_insight should describe what the data shows.
9. For infographic slides: each item needs a concise title (≤8 words) and brief description.
10. For KPI cards: each item needs a label (title) and a numeric value.
11. Bullet points should NOT start with "The" — use active, direct language.
12. Avoid filler phrases like "It is important to note that" or "Research shows that".
"""


# ── Build source content for each slide ──────────────────────────────

def _section_to_text(sec: ContentSection, depth: int = 0) -> str:
    """Convert a ContentSection to readable text for the LLM prompt."""
    parts = []
    indent = "  " * depth

    parts.append(f"{indent}## {sec.heading}")

    if sec.text:
        parts.append(f"{indent}{sec.text}")

    if sec.bullets:
        for b in sec.bullets:
            parts.append(f"{indent}- {b}")

    if sec.metrics:
        metrics_str = ", ".join(f"{m.label}: {m.value}" for m in sec.metrics)
        parts.append(f"{indent}Metrics: {metrics_str}")

    if sec.tables:
        for t in sec.tables:
            title = f" ({t.title})" if t.title else ""
            parts.append(f"{indent}Table{title}: {' | '.join(t.headers)}")
            for row in t.rows:
                parts.append(f"{indent}  {' | '.join(str(c) for c in row)}")

    for sub in sec.subsections:
        parts.append(_section_to_text(sub, depth + 1))

    return "\n".join(parts)


def _find_sections_for_slide(
    plan_item: SlidePlanItem,
    tree: ContentTree,
) -> list[ContentSection]:
    """Find the ContentSections mapped to a slide via content_source headings."""
    if not plan_item.content_source:
        return []

    source_headings = {h.lower().strip() for h in plan_item.content_source}
    matched = []

    for sec in tree.sections:
        if sec.heading.lower().strip() in source_headings:
            matched.append(sec)
        for sub in sec.subsections:
            if sub.heading.lower().strip() in source_headings:
                matched.append(sub)

    return matched


def _build_slide_source_context(
    plan_item: SlidePlanItem,
    sections: list[ContentSection],
    tree: ContentTree,
) -> str:
    """Build the source content text for a single slide."""
    parts = []

    parts.append(f"Slide {plan_item.slide_number}: type={plan_item.slide_type}, "
                 f"role={plan_item.narrative_role}, viz={plan_item.visualization_hint}")
    if plan_item.chart_type_hint:
        parts.append(f"  Chart type: {plan_item.chart_type_hint}")
    if plan_item.infographic_type_hint:
        parts.append(f"  Infographic type: {plan_item.infographic_type_hint}")
    if plan_item.key_message:
        parts.append(f"  Planner key message: {plan_item.key_message}")

    # Special slides get special content
    if plan_item.slide_type == "cover":
        parts.append(f"\nReport title: {tree.title}")
        if tree.subtitle:
            parts.append(f"Report subtitle: {tree.subtitle}")

    elif plan_item.slide_type == "agenda":
        parts.append("\nSections in the report:")
        for sec in tree.sections:
            parts.append(f"  - {sec.heading}")

    elif plan_item.slide_type in ("executive_summary", "conclusion"):
        if tree.executive_summary:
            parts.append(f"\nExecutive Summary:\n{tree.executive_summary}")
        # Also include section highlights for conclusion
        if plan_item.slide_type == "conclusion":
            parts.append("\nAll section headings and key data:")
            for sec in tree.sections:
                first_insight = ""
                if sec.metrics:
                    first_insight = f" — Key metric: {sec.metrics[0].label}: {sec.metrics[0].value}"
                elif sec.bullets:
                    first_insight = f" — {sec.bullets[0][:100]}"
                parts.append(f"  - {sec.heading}{first_insight}")

    elif plan_item.slide_type == "thank_you":
        parts.append(f"\nReport title: {tree.title}")

    # Add mapped section content
    if sections:
        parts.append("\nSource content:")
        for sec in sections:
            parts.append(_section_to_text(sec))

    return "\n".join(parts)


# ── Main content writer ──────────────────────────────────────────────

def write_deck_content(
    slide_plan: SlidePlan,
    content_tree: ContentTree,
    content_profile=None,
) -> DeckContent:
    """Write presentation-ready content for ALL slides via a single LLM call.

    This is the core AI content generation step. It transforms raw markdown
    content into polished, insight-driven slide content.
    """
    logger.info(f"Writing content for {len(slide_plan.slides)} slides")

    # Build the user prompt with all slide contexts
    user_parts = []
    user_parts.append(f"PRESENTATION: {content_tree.title}")
    user_parts.append(f"Total slides: {len(slide_plan.slides)}")
    user_parts.append(f"Storyline: {slide_plan.storyline_summary}")
    user_parts.append("")

    # Add profile context if available
    if content_profile:
        user_parts.append(f"Content archetype: {content_profile.archetype}")
        user_parts.append(f"Data richness: {content_profile.data_richness}")
        user_parts.append("")

    # Build per-slide source content (skip thank_you — template provides it)
    for plan_item in slide_plan.slides:
        if plan_item.slide_type == "thank_you":
            continue
        sections = _find_sections_for_slide(plan_item, content_tree)
        slide_context = _build_slide_source_context(plan_item, sections, content_tree)
        user_parts.append(f"--- SLIDE {plan_item.slide_number} ---")
        user_parts.append(slide_context)
        user_parts.append("")

    user_prompt = "\n".join(user_parts)
    user_prompt += "\n\nGenerate the DeckContent JSON with storyline_summary and content for ALL slides."

    logger.info(f"Content writer prompt: {len(user_prompt)} chars")

    deck_content = invoke_llm_structured(
        system_prompt=CONTENT_WRITER_SYSTEM_PROMPT,
        user_prompt=user_prompt,
        output_schema=DeckContent,
        estimated_tokens=15000,
    )

    logger.info(f"Content written: {len(deck_content.slides)} slides generated")

    # Validate slide count matches
    if len(deck_content.slides) != len(slide_plan.slides):
        logger.warning(
            f"Content writer produced {len(deck_content.slides)} slides "
            f"but plan has {len(slide_plan.slides)} — adjusting"
        )

    return deck_content


# ── Fix content (for validation recovery) ────────────────────────────

FIX_CONTENT_SYSTEM_PROMPT = """\
You are a presentation content editor. You will receive slide content that failed \
validation checks, along with specific issues to fix.

Fix ONLY the reported issues while preserving the content quality:
- If bullets are too long, condense them (keep the insight, reduce words).
- If there are too many bullets, merge similar ones or drop the weakest.
- If text overflows, shorten it while keeping the key message.
- If content is duplicated across slides, rewrite to be unique.

Return the corrected DeckContent with all slides.
"""


def fix_deck_content(
    deck_content: DeckContent,
    validation_issues: list[str],
) -> DeckContent:
    """Ask the AI to fix specific content issues flagged by validation."""
    logger.info(f"Fixing {len(validation_issues)} content issues")

    user_prompt = f"CURRENT DECK CONTENT:\n"
    for sc in deck_content.slides:
        user_prompt += f"\nSlide {sc.slide_number}: {sc.title}\n"
        if sc.bullets:
            for b in sc.bullets:
                user_prompt += f"  - {b}\n"
        if sc.chart_insight:
            user_prompt += f"  Chart insight: {sc.chart_insight}\n"

    user_prompt += "\n\nISSUES TO FIX:\n"
    for issue in validation_issues:
        user_prompt += f"  - {issue}\n"

    user_prompt += "\n\nReturn the corrected DeckContent JSON."

    fixed = invoke_llm_structured(
        system_prompt=FIX_CONTENT_SYSTEM_PROMPT,
        user_prompt=user_prompt,
        output_schema=DeckContent,
        estimated_tokens=5000,
    )

    logger.info(f"Content fixed: {len(fixed.slides)} slides")
    return fixed
