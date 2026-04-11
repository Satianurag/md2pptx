"""Content Profiler — signal-based analysis of ContentTree.

Produces a ``ContentProfile`` that drives adaptive downstream decisions:
slide planning, spec generation, visualization choices, and validation.

Key design principles:
- **Signal-based scoring, not hard categories.** Every archetype gets a
  score 0-100 from weighted signals. The highest wins but ALL signals
  feed into recommendations.
- **"mixed" is fully functional.** If no archetype exceeds a threshold,
  ``mixed`` produces the same visual quality — balanced charts,
  infographics, bullets, and all visual effects.
- **Pure Python, no LLM call.** Runs in <100 ms even on 25 MB files.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

from .schemas import ContentTree, ContentSection, DataTable, KeyMetric

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Keyword patterns for signal detection
# ---------------------------------------------------------------------------
_PROCESS_RE = re.compile(
    r"\b(step|stage|phase|process|pipeline|workflow|methodology|procedure|framework"
    r"|approach|implementation|deployment|rollout)\b",
    re.I,
)
_TIMELINE_RE = re.compile(
    r"\b(year|quarter|month|20[12]\d|Q[1-4]|timeline|milestone|roadmap"
    r"|2020|2021|2022|2023|2024|2025|2026|FY\d{2}|fiscal\s+year)\b",
    re.I,
)
_COMPARISON_RE = re.compile(
    r"\b(vs\.?|versus|compare|comparison|benchmark|ranking|competitor"
    r"|alternative|pros|cons|advantage|disadvantage|differ)\b",
    re.I,
)
_MARKET_RE = re.compile(
    r"\b(market|revenue|growth|CAGR|valuation|investment|billion|million"
    r"|market\s+share|forecast|projection|TAM|SAM|SOM)\b",
    re.I,
)
_FINANCIAL_RE = re.compile(
    r"\b(ROE|ROA|ROI|EBITDA|EPS|P/E|margin|ratio|debt|equity|asset"
    r"|profit|loss|balance\s+sheet|cash\s+flow|dividend|capital)\b",
    re.I,
)
_NUMBER_RE = re.compile(r"[\$€£]?\d[\d,.]*[%BbMmKk]?")


# ---------------------------------------------------------------------------
# Scored items
# ---------------------------------------------------------------------------

@dataclass
class ScoredTable:
    """A DataTable ranked by chart-worthiness."""
    table: DataTable
    score: float = 0.0
    preferred_chart_type: str = "column"


@dataclass
class ScoredMetric:
    """A KeyMetric ranked by visual impact."""
    metric: KeyMetric
    score: float = 0.0


# ---------------------------------------------------------------------------
# ContentProfile  (the output)
# ---------------------------------------------------------------------------

@dataclass
class ContentProfile:
    archetype: str = "mixed"
    archetype_scores: dict[str, float] = field(default_factory=dict)
    data_richness: str = "medium"  # "high", "medium", "low"
    recommended_visual_ratio: float = 0.6
    recommended_chart_types: list[str] = field(default_factory=list)
    recommended_infographic_types: list[str] = field(default_factory=list)
    best_tables: list[ScoredTable] = field(default_factory=list)
    best_metrics: list[ScoredMetric] = field(default_factory=list)
    sections_by_value: list[str] = field(default_factory=list)
    # Raw signal counts for debugging / downstream use
    total_sections: int = 0
    total_tables: int = 0
    total_metrics: int = 0
    total_bullets: int = 0
    total_text_chars: int = 0


# ---------------------------------------------------------------------------
# Profiler
# ---------------------------------------------------------------------------

def profile_content(tree: ContentTree) -> ContentProfile:
    """Analyse *tree* and return a :class:`ContentProfile`.

    Pure-Python, no LLM call.
    """
    prof = ContentProfile()

    # ── Gather raw counts ──
    all_text = tree.executive_summary + " "
    all_bullets: list[str] = []

    def _walk(sec: ContentSection) -> None:
        prof.total_sections += 1
        all_text_parts.append(sec.text)
        all_bullets.extend(sec.bullets)
        prof.total_bullets += len(sec.bullets)
        prof.total_text_chars += len(sec.text)
        prof.total_tables += len(sec.tables)
        prof.total_metrics += len(sec.metrics)
        for sub in sec.subsections:
            _walk(sub)

    all_text_parts: list[str] = [all_text]
    for sec in tree.sections:
        _walk(sec)

    combined_text = " ".join(all_text_parts) + " " + " ".join(all_bullets)
    prof.total_tables += len(tree.all_tables)
    prof.total_metrics += len(tree.all_metrics)

    n_sec = max(prof.total_sections, 1)

    # ── Signal densities ──
    table_density = prof.total_tables / n_sec
    metric_density = prof.total_metrics / n_sec
    text_density = prof.total_text_chars / n_sec
    bullet_density = prof.total_bullets / n_sec

    process_hits = len(_PROCESS_RE.findall(combined_text))
    timeline_hits = len(_TIMELINE_RE.findall(combined_text))
    comparison_hits = len(_COMPARISON_RE.findall(combined_text))
    market_hits = len(_MARKET_RE.findall(combined_text))
    financial_hits = len(_FINANCIAL_RE.findall(combined_text))

    # ── Archetype scoring ──
    scores: dict[str, float] = {}

    # Financial
    scores["financial"] = (
        min(financial_hits * 3, 40)
        + min(metric_density * 20, 30)
        + min(table_density * 10, 20)
        + (10 if "ROE" in combined_text or "ROA" in combined_text else 0)
    )

    # Market analysis
    scores["market_analysis"] = (
        min(market_hits * 2.5, 35)
        + min(table_density * 12, 25)
        + min(metric_density * 15, 20)
        + min(timeline_hits * 2, 20)
    )

    # Narrative tech (text-heavy, fewer numbers)
    scores["narrative_tech"] = (
        min(text_density / 50, 30)
        + min(bullet_density * 5, 25)
        + max(0, 20 - metric_density * 10)
        + max(0, 15 - table_density * 15)
        + min(process_hits * 2, 10)
    )

    # Process / methodology
    scores["process_methodology"] = (
        min(process_hits * 4, 40)
        + min(bullet_density * 5, 25)
        + min(timeline_hits * 2, 15)
        + (20 if prof.total_tables <= 2 and process_hits >= 5 else 0)
    )

    # Comparative
    scores["comparative"] = (
        min(comparison_hits * 4, 40)
        + min(table_density * 15, 25)
        + min(metric_density * 10, 15)
        + (20 if prof.total_tables >= 3 and comparison_hits >= 3 else 0)
    )

    prof.archetype_scores = scores

    # Determine winner (threshold: 20)
    best = max(scores, key=scores.get)  # type: ignore[arg-type]
    if scores[best] >= 20:
        prof.archetype = best
    else:
        prof.archetype = "mixed"

    log.info("Content profile: archetype=%s (scores=%s)", prof.archetype,
             {k: round(v, 1) for k, v in scores.items()})

    # ── Data richness ──
    if prof.total_tables >= 4 or prof.total_metrics >= 10:
        prof.data_richness = "high"
    elif prof.total_tables >= 2 or prof.total_metrics >= 5:
        prof.data_richness = "medium"
    else:
        prof.data_richness = "low"

    # ── Recommended visual ratio ──
    prof.recommended_visual_ratio = {
        "financial": 0.7,
        "market_analysis": 0.7,
        "comparative": 0.65,
        "process_methodology": 0.6,
        "narrative_tech": 0.5,
        "mixed": 0.6,
    }.get(prof.archetype, 0.6)

    # ── Recommended chart types ──
    chart_types: list[str] = []
    if timeline_hits >= 3:
        chart_types.append("line")
    if prof.total_tables >= 2:
        chart_types.append("bar")
    if metric_density > 1:
        chart_types.append("column")
    if comparison_hits >= 2 and prof.total_tables >= 1:
        chart_types.append("bar")
    if not chart_types:
        chart_types = ["column", "bar"]
    prof.recommended_chart_types = list(dict.fromkeys(chart_types))  # dedup

    # ── Recommended infographic types ──
    infographic_types: list[str] = []
    if process_hits >= 3:
        infographic_types.append("process_flow")
    if timeline_hits >= 3:
        infographic_types.append("timeline")
    if comparison_hits >= 2:
        infographic_types.append("comparison")
    if metric_density >= 1:
        infographic_types.append("kpi_cards")
    if not infographic_types:
        infographic_types = ["comparison", "kpi_cards"]
    prof.recommended_infographic_types = list(dict.fromkeys(infographic_types))

    # ── Score tables ──
    all_tables = list(tree.all_tables)
    for sec in tree.sections:
        all_tables.extend(sec.tables)
        for sub in sec.subsections:
            all_tables.extend(sub.tables)

    scored_tables: list[ScoredTable] = []
    for tbl in all_tables:
        st = _score_table(tbl)
        scored_tables.append(st)
    scored_tables.sort(key=lambda s: s.score, reverse=True)
    prof.best_tables = scored_tables[:10]

    # ── Score metrics ──
    all_metrics = list(tree.all_metrics)
    for sec in tree.sections:
        all_metrics.extend(sec.metrics)
        for sub in sec.subsections:
            all_metrics.extend(sub.metrics)

    scored_metrics: list[ScoredMetric] = []
    seen_labels: set[str] = set()
    for m in all_metrics:
        key = (m.label.strip().lower(), m.value.strip().lower())
        if key in seen_labels:
            continue
        seen_labels.add(key)
        scored_metrics.append(_score_metric(m))
    scored_metrics.sort(key=lambda s: s.score, reverse=True)
    prof.best_metrics = scored_metrics[:10]

    # ── Sections by data value ──
    sec_scores: list[tuple[str, float]] = []
    for sec in tree.sections:
        s = (
            len(sec.tables) * 3
            + len(sec.metrics) * 2
            + len(sec.bullets) * 0.5
            + len(sec.text) / 200
        )
        for sub in sec.subsections:
            s += len(sub.tables) * 3 + len(sub.metrics) * 2
        sec_scores.append((sec.heading, s))
    sec_scores.sort(key=lambda x: -x[1])
    prof.sections_by_value = [h for h, _ in sec_scores]

    return prof


# ---------------------------------------------------------------------------
# Table scoring
# ---------------------------------------------------------------------------

def _score_table(tbl: DataTable) -> ScoredTable:
    """Score a table for chart-worthiness."""
    score = 0.0
    n_rows = len(tbl.rows)
    n_cols = len(tbl.headers)

    if n_cols <= 1 or n_rows == 0:
        return ScoredTable(table=tbl, score=0.0)

    # Count numeric columns
    numeric_cols = 0
    for col_idx in range(1, n_cols):
        nums = sum(
            1 for row in tbl.rows
            if col_idx < len(row) and _NUMBER_RE.search(str(row[col_idx]))
        )
        if nums > n_rows * 0.4:
            numeric_cols += 1

    score += numeric_cols * 3
    if n_rows >= 5:
        score += 2
    score += min(n_rows, 20)
    text_cols = max(n_cols - 1 - numeric_cols, 0)
    if text_cols > numeric_cols:
        score -= 2
    if n_rows <= 2:
        score -= 1
    # Time-series bonus
    first_vals = [str(row[0]) if row else "" for row in tbl.rows]
    year_hits = sum(1 for v in first_vals if re.match(r"^(19|20)\d{2}", v))
    quarter_hits = sum(1 for v in first_vals if re.match(r"^Q[1-4]", v, re.I))
    if year_hits >= 3 or quarter_hits >= 3:
        score += 2

    # Determine preferred chart type
    chart_type = "column"
    if year_hits >= 3 or quarter_hits >= 3:
        chart_type = "line" if numeric_cols <= 3 else "area"
    elif numeric_cols == 1 and 2 <= n_rows <= 6:
        chart_type = "pie"
    elif n_rows > 6:
        chart_type = "bar"

    return ScoredTable(table=tbl, score=score, preferred_chart_type=chart_type)


def _score_metric(m: KeyMetric) -> ScoredMetric:
    """Score a metric for KPI card impact."""
    score = 0.0
    val = m.value
    label = m.label

    if "$" in val or "billion" in val.lower() or "million" in val.lower():
        score += 3
    if "%" in val:
        score += 2
    if len(label) < 30:
        score += 1
    if any(w in label.lower() for w in ("the", "this", "that", "also")):
        score -= 1
    if _NUMBER_RE.search(val):
        score += 1

    return ScoredMetric(metric=m, score=score)
