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

from .schemas import (
    ContentTree, ContentSection, DataTable, KeyMetric,
    ContentInventory, SectionInventory, TableInventory,
)

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


# ---------------------------------------------------------------------------
# Narrative role classification (pure Python, no LLM)
# ---------------------------------------------------------------------------

_ROLE_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("executive_summary", re.compile(
        r"\b(executive\s+summary|overview|abstract|key\s+highlights|at\s+a\s+glance)\b", re.I)),
    ("market_landscape", re.compile(
        r"\b(market\s+(overview|landscape|opportunity|size|analysis|dynamics|context)"
        r"|industry\s+(overview|landscape|analysis)|current\s+(landscape|state|market)"
        r"|competitive\s+landscape|sector\s+overview)\b", re.I)),
    ("methodology", re.compile(
        r"\b(methodo|research\s+design|approach|framework|sampling|data\s+collection"
        r"|analytical\s+framework|study\s+design|research\s+scope)\b", re.I)),
    ("timeline_roadmap", re.compile(
        r"\b(timeline|roadmap|future\s+outlook|forecast|projection|outlook"
        r"|phased?\s+plan|implementation\s+timeline|20\d{2}\s*[-–]\s*20\d{2})\b", re.I)),
    ("case_study", re.compile(
        r"\b(case\s+stud|real[\s-]world\s+example|success\s+stor|proof\s+of\s+concept"
        r"|toyota|nissan|pfizer|schneider|company\s+profile)\b", re.I)),
    ("regional_analysis", re.compile(
        r"\b(regional\s+(analysis|deep\s+dive|overview|breakdown|focus)"
        r"|middle\s+east|NCR|europe|asia|india|UAE|GCC|country[\s-]specific"
        r"|geographic|local\s+market)\b", re.I)),
    ("challenges_risks", re.compile(
        r"\b(challeng|risk\s*(assessment|analysis|factor)?|barrier|limitation|SWOT"
        r"|threat|obstacle|constraint|gap\s+analysis|vulnerabilit)\b", re.I)),
    ("recommendations", re.compile(
        r"\b(recommend|strategic\s+(guidance|imperative|direction)|action\s+item"
        r"|next\s+step|implementation\s+plan|policy\s+implication|call\s+to\s+action"
        r"|strategic\s+action)\b", re.I)),
    ("impact_analysis", re.compile(
        r"\b(economic\s+impact|GDP\s+contribution|ROI\s+analysis|cost[\s-]benefit"
        r"|job\s+creation|financial\s+impact|impact\s+assessment|multiplier\s+effect"
        r"|value\s+creation)\b", re.I)),
    ("data_evidence", re.compile(
        r"\b(data\s+analysis|statistical|evidence|performance\s+data"
        r"|quantitative|metric|KPI|benchmark|scorecard)\b", re.I)),
    ("key_findings", re.compile(
        r"\b(finding|result|analysis|trend|insight|assessment|evaluation"
        r"|comparative|impact|overview\s+of\s+result|key\s+observation)\b", re.I)),
]


def classify_sections(tree: ContentTree) -> dict[str, str]:
    """Map section headings → narrative_role using keyword patterns.

    Returns ``{heading: role}``.  Pure Python, no LLM call.
    """
    mapping: dict[str, str] = {}

    for sec in tree.sections:
        role = _classify_one_section(sec)
        mapping[sec.heading] = role
        for sub in sec.subsections:
            mapping[sub.heading] = _classify_one_section(sub)

    return mapping


def _classify_one_section(sec: ContentSection) -> str:
    """Classify a single section by heading + content signals."""
    heading_lower = sec.heading.lower()

    # Skip structural headings
    if heading_lower in ("table of contents", "contents", "references",
                         "references and source documentation", "bibliography"):
        return "key_findings"  # safe fallback

    # Pattern matching on heading text (highest priority)
    for role, pattern in _ROLE_PATTERNS:
        if pattern.search(sec.heading):
            return role

    # Data-signal fallback: sections rich in tables/metrics → data_evidence
    n_tables = len(sec.tables) + sum(len(s.tables) for s in sec.subsections)
    n_metrics = len(sec.metrics) + sum(len(s.metrics) for s in sec.subsections)
    if n_tables >= 2 or n_metrics >= 4:
        return "data_evidence"

    # Conclusion-like headings
    if re.search(r"\b(conclusion|summary|key\s+takeaway|wrap[\s-]up)\b", heading_lower):
        return "conclusion"

    return "key_findings"  # safest default


def score_section_importance(sec: ContentSection) -> float:
    """Score a section 0.0–1.0 based on data richness and substance.

    Pure Python, no LLM call.
    """
    score = 0.0

    # Tables: +0.15 each, capped at 0.4
    n_tables = len(sec.tables) + sum(len(s.tables) for s in sec.subsections)
    score += min(n_tables * 0.15, 0.4)

    # Metrics: +0.08 each, capped at 0.3
    n_metrics = len(sec.metrics) + sum(len(s.metrics) for s in sec.subsections)
    score += min(n_metrics * 0.08, 0.3)

    # Numeric density in text: count numbers in text + bullets
    combined = sec.text + " " + " ".join(sec.bullets)
    num_count = len(_NUMBER_RE.findall(combined))
    score += min(num_count * 0.02, 0.15)

    # Substance: text length
    total_chars = len(sec.text) + sum(len(b) for b in sec.bullets)
    for sub in sec.subsections:
        total_chars += len(sub.text) + sum(len(b) for b in sub.bullets)
    if total_chars > 200:
        score += 0.1
    if total_chars > 500:
        score += 0.05

    return min(score, 1.0)


def generate_action_title(sec: ContentSection, role: str) -> str:
    """Generate a data-driven action title from section content.

    Extracts the top metric or numeric finding and combines it with the
    section theme.  Falls back to the original heading if no data found.

    Pure Python, no LLM call.

    Examples:
        "Market Overview"  → "Global Semiconductor Market Reached $580B in 2025"
        "ROE Analysis"     → "FAB Leads with 20% ROTE — Bank ABC Trails at 7.1%"
    """
    heading = sec.heading.strip()

    # Collect all metrics from section + subsections
    all_metrics: list[KeyMetric] = list(sec.metrics)
    for sub in sec.subsections:
        all_metrics.extend(sub.metrics)

    # Strategy 1: Use the best metric
    if all_metrics:
        scored = [_score_metric(m) for m in all_metrics]
        scored.sort(key=lambda s: s.score, reverse=True)
        best = scored[0].metric
        title = f"{heading} — {best.label}: {best.value}"
        if len(title) <= 80:
            return title
        # Truncate label if too long
        return f"{heading} — {best.value}"[:80]

    # Strategy 2: Extract first strong number from text/bullets
    combined = sec.text + " " + " ".join(sec.bullets[:4])
    # Look for sentences with numbers
    sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', combined) if s.strip()]
    for sentence in sentences[:5]:
        nums = _NUMBER_RE.findall(sentence)
        if nums and len(sentence) <= 90:
            # Use this sentence as the title (it has data)
            return sentence[:80]

    # Strategy 3: Extract key number from first table
    all_tables = list(sec.tables)
    for sub in sec.subsections:
        all_tables.extend(sub.tables)
    if all_tables:
        tbl = all_tables[0]
        if tbl.rows and tbl.headers and len(tbl.headers) >= 2:
            # Use first row's key value
            first_row = tbl.rows[0]
            if len(first_row) >= 2:
                category = str(first_row[0])[:30]
                value = str(first_row[1])[:20]
                title = f"{heading}: {category} at {value}"
                if len(title) <= 80:
                    return title

    # Fallback: original heading
    return heading


# ---------------------------------------------------------------------------
# Content Inventory builder (research §5.2 — Phase 1 → Phase 2 bridge)
# ---------------------------------------------------------------------------

def build_content_inventory(tree: ContentTree) -> ContentInventory:
    """Build a compact ~500-token ContentInventory from a parsed ContentTree.

    This is the PRIMARY input to the LLM content triage call (Phase 2).
    It summarises structure and signals without including full text.
    """
    _num_re = re.compile(r"[\$€£]?\d[\d,.]*[%BMKTbmkt]?")
    _temporal_re = re.compile(
        r"\b(20[12]\d|Q[1-4]|year|month|quarter|FY\d{2})\b", re.I
    )

    section_items: list[SectionInventory] = []
    table_items: list[TableInventory] = []
    total_words = 0
    total_bullets = 0
    total_images = 0

    def _count_words(text: str) -> int:
        return len(text.split()) if text else 0

    def _walk_section(sec: ContentSection) -> SectionInventory:
        nonlocal total_words, total_bullets, total_images

        wc = _count_words(sec.text) + sum(_count_words(b) for b in sec.bullets)
        total_words += wc
        total_bullets += len(sec.bullets)

        combined = sec.text + " " + " ".join(sec.bullets[:10])
        has_numeric = bool(_num_re.search(combined))
        has_temporal = bool(_temporal_re.search(combined))

        for tbl in sec.tables:
            tbl_text = " ".join(tbl.headers) + " ".join(
                " ".join(row) for row in tbl.rows[:3]
            )
            t_inv = TableInventory(
                section_heading=sec.heading,
                column_count=len(tbl.headers),
                row_count=len(tbl.rows),
                has_numeric=bool(_num_re.search(tbl_text)),
                has_temporal=bool(_temporal_re.search(tbl_text)),
                chart_worthy=len(tbl.rows) >= 2 and bool(_num_re.search(tbl_text)),
            )
            table_items.append(t_inv)

        inv = SectionInventory(
            heading=sec.heading,
            level=sec.level,
            word_count=wc,
            bullet_count=len(sec.bullets),
            table_count=len(sec.tables),
            image_count=0,
            has_numeric_data=has_numeric,
            has_temporal_data=has_temporal,
            subsection_count=len(sec.subsections),
        )
        section_items.append(inv)

        for sub in sec.subsections:
            _walk_section(sub)

        return inv

    for sec in tree.sections:
        _walk_section(sec)

    return ContentInventory(
        title=tree.title,
        subtitle=tree.subtitle,
        total_sections=len(section_items),
        total_words=total_words,
        total_tables=len(tree.all_tables),
        total_images=total_images,
        total_bullets=total_bullets,
        has_executive_summary=bool(tree.executive_summary),
        sections=section_items,
        tables=table_items,
    )
