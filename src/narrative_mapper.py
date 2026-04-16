"""Narrative Mapper — rule-based narrative intelligence for slide planning.

Builds a narrative-arc-aware slide plan by:
1. Classifying sections into semantic roles (market_landscape, case_study, etc.)
2. Scoring sections by importance (data richness, substance)
3. Selecting which roles to keep based on target slide count
4. Merging cut-role content into surviving neighbours
5. Generating action-oriented titles from section data

Pure Python, no LLM call.  Runs in <50 ms.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

from .schemas import ContentTree, ContentSection
from .content_profiler import (
    ContentProfile,
    classify_sections,
    score_section_importance,
    generate_action_title,
)
from . import config

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# NarrativeSlot — one slot in the narrative arc
# ---------------------------------------------------------------------------

@dataclass
class NarrativeSlot:
    """A single planned slide slot in the narrative arc."""
    role: str                                    # narrative_role from config.NARRATIVE_ROLES
    sections: list[ContentSection] = field(default_factory=list)  # mapped source sections
    importance: float = 0.0                      # max importance of mapped sections
    action_title: str = ""                       # data-driven title
    visualization_hint: str = "bullets"          # suggested visualization
    merged_from: list[str] = field(default_factory=list)  # roles merged into this slot
    key_message: str = ""                        # extracted key message


# ---------------------------------------------------------------------------
# Visualization hint selection (rule-based)
# ---------------------------------------------------------------------------

_ROLE_VIZ_DEFAULTS: dict[str, str] = {
    "cover": "text",
    "agenda": "bullets",
    "executive_summary": "bullets",
    "market_landscape": "chart",
    "methodology": "infographic",
    "key_findings": "mixed",
    "data_evidence": "chart",
    "timeline_roadmap": "infographic",
    "case_study": "mixed",
    "regional_analysis": "chart",
    "challenges_risks": "infographic",
    "recommendations": "bullets",
    "impact_analysis": "chart",
    "conclusion": "bullets",
    "thank_you": "text",
}


def _pick_viz_hint(sec: ContentSection, role: str) -> str:
    """Choose visualization hint based on section data signals + role default."""
    n_tables = len(sec.tables) + sum(len(s.tables) for s in sec.subsections)
    n_metrics = len(sec.metrics) + sum(len(s.metrics) for s in sec.subsections)

    # Strong data signals override role default
    if n_tables >= 2:
        return "chart"
    if n_metrics >= 3:
        return "kpi"
    if n_tables == 1 and n_metrics >= 1:
        return "mixed"

    return _ROLE_VIZ_DEFAULTS.get(role, "bullets")


def _extract_key_message(sec: ContentSection) -> str:
    """Extract the single most important sentence from a section."""
    # Prefer first bullet
    if sec.bullets:
        return sec.bullets[0][:200]
    # Then first sentence of text
    if sec.text:
        import re
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', sec.text) if len(s.strip()) > 15]
        if sentences:
            return sentences[0][:200]
    # Then subsection content
    for sub in sec.subsections:
        msg = _extract_key_message(sub)
        if msg:
            return msg
    return ""


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_narrative_plan(
    tree: ContentTree,
    profile: ContentProfile,
    target_slides: int = 15,
) -> list[NarrativeSlot]:
    """Build a narrative-arc-aware slide plan, purely rule-based.

    Returns ordered ``NarrativeSlot`` list ready to feed the LLM planner
    as a suggestion or to be used directly by the rule-based fallback.
    """
    target_slides = max(config.MIN_SLIDES, min(config.MAX_SLIDES, target_slides))

    # 1. Classify all sections → roles
    role_map = classify_sections(tree)
    log.info("Section role classification: %s",
             {h: r for h, r in list(role_map.items())[:10]})

    # 2. Score all sections for importance
    importance_map: dict[str, float] = {}
    for sec in tree.sections:
        importance_map[sec.heading] = score_section_importance(sec)

    # 3. Group sections by role
    sections_by_role: dict[str, list[ContentSection]] = {}
    for sec in tree.sections:
        role = role_map.get(sec.heading, "key_findings")
        sections_by_role.setdefault(role, []).append(sec)

    # 4. Look up which roles to keep at this slide count
    keep_roles = config.SLIDE_ROLES_BY_COUNT.get(
        target_slides, config.SLIDE_ROLES_BY_COUNT[15]
    )

    # 5. Build merge map: cut roles → surviving neighbour
    merge_targets: dict[str, str] = {}
    all_roles_with_sections = set(sections_by_role.keys())
    for role in all_roles_with_sections:
        if role not in keep_roles and role not in ("cover", "agenda", "conclusion", "thank_you"):
            target_role = config.MERGE_TARGET.get(role, "key_findings")
            # Ensure target itself survives
            if target_role not in keep_roles:
                target_role = "key_findings"
            merge_targets[role] = target_role

    # 6. Apply merges
    for cut_role, target_role in merge_targets.items():
        sections_to_merge = sections_by_role.pop(cut_role, [])
        if sections_to_merge:
            sections_by_role.setdefault(target_role, []).extend(sections_to_merge)
            log.info("Merged role '%s' (%d sections) into '%s'",
                     cut_role, len(sections_to_merge), target_role)

    # 7. Build ordered slots following the narrative arc
    slots: list[NarrativeSlot] = []
    used_content_roles: set[str] = set()

    for role in keep_roles:
        # Structural bookends — always present even without content
        if role in ("cover", "agenda", "conclusion", "thank_you"):
            slot = NarrativeSlot(role=role, importance=1.0)
            if role == "cover":
                slot.action_title = tree.title or "Presentation"
                slot.visualization_hint = "text"
            elif role == "agenda":
                slot.action_title = "Agenda"
                slot.visualization_hint = "bullets"
            elif role == "conclusion":
                slot.action_title = "Key Takeaways & Recommendations"
                slot.visualization_hint = "bullets"
            elif role == "thank_you":
                slot.action_title = "Thank You"
                slot.visualization_hint = "text"
            slots.append(slot)
            continue

        # Content roles — map sections
        matched = sections_by_role.get(role, [])
        if not matched:
            # No sections match this role — try to find unmapped sections
            unmapped = [
                sec for sec in tree.sections
                if sec.heading not in used_content_roles
                and role_map.get(sec.heading) == "key_findings"  # generic sections
            ]
            if unmapped:
                matched = [unmapped[0]]

        if not matched:
            continue  # skip empty roles (deck will be shorter)

        # Pick the best section(s) for this slot
        matched.sort(key=lambda s: importance_map.get(s.heading, 0), reverse=True)
        primary = matched[0]
        used_content_roles.add(primary.heading)

        slot = NarrativeSlot(
            role=role,
            sections=[primary] + matched[1:2],  # primary + up to 1 merged
            importance=importance_map.get(primary.heading, 0.5),
            action_title=generate_action_title(primary, role),
            visualization_hint=_pick_viz_hint(primary, role),
            key_message=_extract_key_message(primary),
        )

        # Track merged roles
        for merged_role, target in merge_targets.items():
            if target == role:
                slot.merged_from.append(merged_role)

        slots.append(slot)

    # 8. Fill remaining slots if we're short and have unused sections
    if len(slots) < target_slides:
        used_headings = set()
        for sl in slots:
            for sec in sl.sections:
                used_headings.add(sec.heading)

        unused = [
            sec for sec in tree.sections
            if sec.heading not in used_headings
        ]
        unused.sort(key=lambda s: importance_map.get(s.heading, 0), reverse=True)

        # Insert before conclusion
        insert_idx = next(
            (i for i, sl in enumerate(slots) if sl.role == "conclusion"),
            len(slots),
        )
        for sec in unused:
            if len(slots) >= target_slides:
                break
            role = role_map.get(sec.heading, "key_findings")
            slot = NarrativeSlot(
                role=role,
                sections=[sec],
                importance=importance_map.get(sec.heading, 0.3),
                action_title=generate_action_title(sec, role),
                visualization_hint=_pick_viz_hint(sec, role),
                key_message=_extract_key_message(sec),
            )
            slots.insert(insert_idx, slot)
            insert_idx += 1

    log.info("Narrative plan: %d slots for target %d slides — roles: %s",
             len(slots), target_slides,
             [s.role for s in slots])

    return slots


# ---------------------------------------------------------------------------
# Helper: convert NarrativeSlot list → text for LLM planner prompt
# ---------------------------------------------------------------------------

def narrative_plan_to_prompt(slots: list[NarrativeSlot]) -> str:
    """Format the narrative plan as structured context for the LLM planner."""
    lines = [
        "NARRATIVE PLAN (use as primary guide for slide structure):",
        f"Total planned slots: {len(slots)}",
        "",
    ]
    for i, slot in enumerate(slots, 1):
        sec_names = [s.heading for s in slot.sections]
        merged = f" [merged: {', '.join(slot.merged_from)}]" if slot.merged_from else ""
        lines.append(
            f"  Slot {i}: role={slot.role}, viz={slot.visualization_hint}, "
            f"importance={slot.importance:.2f}{merged}"
        )
        lines.append(f"    Title: {slot.action_title}")
        if sec_names:
            lines.append(f"    Source sections: {', '.join(sec_names)}")
        if slot.key_message:
            lines.append(f"    Key message: {slot.key_message[:120]}")
        lines.append("")

    return "\n".join(lines)
