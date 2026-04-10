from __future__ import annotations
import logging
from pathlib import Path
from typing import TypedDict, Optional, Any

from langgraph.graph import StateGraph, START, END
from langgraph.types import RetryPolicy

from .schemas import (
    ContentTree, SlideMasterInfo, SlidePlan, PresentationSpec,
)
from .markdown_parser import parse_markdown
from .slide_master import read_slide_master, auto_detect_template
from .slide_planner import plan_slides
from .spec_generator import generate_presentation_spec
from .validator import validate_and_fix, ValidationResult
from .pptx_renderer import render_presentation
from .content_chunker import chunk_content_tree
from . import config

logger = logging.getLogger(__name__)


# ── Pipeline state ──

class PipelineState(TypedDict, total=False):
    md_text: str
    md_path: str
    content_tree: Optional[ContentTree]
    template_path: str
    master_info: Optional[SlideMasterInfo]
    slide_plan: Optional[SlidePlan]
    presentation_spec: Optional[PresentationSpec]
    output_path: str
    target_slide_count: int
    errors: list[str]
    warnings: list[str]
    retry_count: int
    validation_result: Optional[ValidationResult]


# ── Node functions ──

def parse_md_node(state: PipelineState) -> dict:
    """Parse markdown text into ContentTree."""
    logger.info("Node: parse_markdown")
    md_text = state["md_text"]

    if not md_text:
        return {"errors": state.get("errors", []) + ["Empty markdown input"]}

    # Log file size — no hard reject (FAQ Q15 requires all test cases)
    size_mb = len(md_text.encode("utf-8")) / (1024 * 1024)
    if size_mb > 5:
        logger.warning(f"Large input ({size_mb:.1f}MB). Aggressive chunking will be applied.")

    content_tree = parse_markdown(md_text)

    if not content_tree.sections:
        return {"errors": state.get("errors", []) + ["No sections found in markdown"]}

    # Chunk if too large
    content_tree = chunk_content_tree(content_tree)

    logger.info(f"Parsed: title='{content_tree.title}', sections={len(content_tree.sections)}, "
                f"tables={len(content_tree.all_tables)}, metrics={len(content_tree.all_metrics)}")

    return {"content_tree": content_tree}


def analyze_template_node(state: PipelineState) -> dict:
    """Read template and extract layout metadata."""
    logger.info("Node: analyze_template")

    template_path = state.get("template_path", "")
    md_path = state.get("md_path", "")

    # Auto-detect template if not provided
    if not template_path and md_path:
        detected = auto_detect_template(md_path)
        if detected:
            template_path = str(detected)
            logger.info(f"Auto-detected template: {detected.name}")

    master_info = None
    if template_path and Path(template_path).exists():
        try:
            master_info = read_slide_master(template_path)
            logger.info(f"Template loaded: {len(master_info.layouts)} layouts")
        except Exception as e:
            logger.warning(f"Failed to read template: {e}")
            template_path = ""
    else:
        if template_path:
            logger.warning(f"Template not found: {template_path}")
        template_path = ""

    return {"template_path": template_path, "master_info": master_info}


def plan_slides_node(state: PipelineState) -> dict:
    """Use LLM to generate slide plan."""
    logger.info("Node: plan_slides")

    content_tree = state.get("content_tree")
    if not content_tree:
        return {"errors": state.get("errors", []) + ["No content_tree available for planning"]}

    master_info = state.get("master_info")
    target = state.get("target_slide_count", 12)

    try:
        slide_plan = plan_slides(content_tree, master_info, target)
        logger.info(f"Plan: {len(slide_plan.slides)} slides, storyline: {slide_plan.storyline_summary[:80]}...")
        return {"slide_plan": slide_plan}
    except Exception as e:
        logger.warning(f"LLM planning failed ({e}), using rule-based fallback")
        slide_plan = _rule_based_plan_fallback(content_tree, target)
        return {"slide_plan": slide_plan}


def generate_spec_node(state: PipelineState) -> dict:
    """Generate full PresentationSpec from plan."""
    logger.info("Node: generate_spec")

    content_tree = state.get("content_tree")
    slide_plan = state.get("slide_plan")
    master_info = state.get("master_info")
    template_path = state.get("template_path", "")

    if not content_tree or not slide_plan:
        return {"errors": state.get("errors", []) + ["Missing content_tree or slide_plan"]}

    try:
        spec = generate_presentation_spec(
            content_tree=content_tree,
            slide_plan=slide_plan,
            master_info=master_info,
            template_path=template_path,
        )
        logger.info(f"Spec generated: {len(spec.slides)} slides")
        return {"presentation_spec": spec}
    except Exception as e:
        logger.error(f"Spec generation failed: {e}")
        return {"errors": state.get("errors", []) + [f"Spec generation failed: {e}"]}


def validate_node(state: PipelineState) -> dict:
    """Validate the presentation spec."""
    logger.info("Node: validate")

    spec = state.get("presentation_spec")
    if not spec:
        return {"errors": state.get("errors", []) + ["No presentation_spec to validate"]}

    result = validate_and_fix(spec)

    warnings = state.get("warnings", []) + result.warnings

    if result.fixes_applied:
        logger.info(f"Applied {len(result.fixes_applied)} auto-fixes")

    if not result.passed:
        logger.warning(f"Validation failed: {result.errors}")
        return {
            "validation_result": result,
            "warnings": warnings,
            "errors": state.get("errors", []) + result.errors,
        }

    return {"validation_result": result, "warnings": warnings}


def render_node(state: PipelineState) -> dict:
    """Render the final PPTX."""
    logger.info("Node: render")

    spec = state.get("presentation_spec")
    if not spec:
        return {"errors": state.get("errors", []) + ["No presentation_spec to render"]}

    output_path = state.get("output_path", "")
    if not output_path:
        md_path = state.get("md_path", "output")
        stem = Path(md_path).stem if md_path else "presentation"
        output_path = str(config.OUTPUT_DIR / f"{stem}.pptx")

    try:
        result_path = render_presentation(spec, output_path)
        logger.info(f"Rendered: {result_path}")
        return {"output_path": str(result_path)}
    except Exception as e:
        logger.error(f"Rendering failed: {e}")
        return {"errors": state.get("errors", []) + [f"Rendering failed: {e}"]}


# ── Conditional edge ──

def should_render(state: PipelineState) -> str:
    """Decide whether to render or report errors."""
    errors = state.get("errors", [])
    if errors:
        return "error"
    return "render"


def error_node(state: PipelineState) -> dict:
    """Handle pipeline errors gracefully."""
    errors = state.get("errors", [])
    logger.error(f"Pipeline failed with {len(errors)} errors: {errors}")
    return state


def _rule_based_plan_fallback(tree: ContentTree, target: int = 12) -> SlidePlan:
    """Generate a slide plan without LLM when the planner fails."""
    from .schemas import SlidePlan, SlidePlanItem
    slides = [SlidePlanItem(slide_number=1, slide_type="cover", title=tree.title or "Presentation",
                            subtitle=tree.subtitle)]
    idx = 2
    if tree.executive_summary:
        slides.append(SlidePlanItem(slide_number=idx, slide_type="executive_summary",
                                    title="Executive Summary", visualization_hint="bullets"))
        idx += 1
    # Section slides
    max_sections = min(len(tree.sections), target - 3)  # reserve agenda + conclusion + thank_you
    for sec in tree.sections[:max_sections]:
        hint = "bullets"
        if sec.tables:
            hint = "chart"
        elif sec.metrics and len(sec.metrics) >= 3:
            hint = "kpi"
        slides.append(SlidePlanItem(
            slide_number=idx, slide_type="content", title=sec.heading,
            content_source=[sec.heading], visualization_hint=hint,
            key_message=sec.text[:100] if sec.text else "",
        ))
        idx += 1
    slides.append(SlidePlanItem(slide_number=idx, slide_type="conclusion",
                                title="Conclusion & Key Takeaways", visualization_hint="bullets"))
    idx += 1
    slides.append(SlidePlanItem(slide_number=idx, slide_type="thank_you", title="Thank You"))
    logger.info(f"Rule-based plan: {len(slides)} slides")
    return SlidePlan(storyline_summary="Auto-generated from section structure",
                     target_slide_count=len(slides), slides=slides)


# ── Build the graph ──

def build_pipeline() -> StateGraph:
    """Build and compile the LangGraph pipeline."""
    workflow = StateGraph(PipelineState)

    # Retry policy for LLM nodes (exponential backoff)
    llm_retry = RetryPolicy(
        initial_interval=2.0,
        backoff_factor=2.0,
        max_interval=30.0,
        max_attempts=3,
        jitter=True,
    )

    # Add nodes
    workflow.add_node("parse_md", parse_md_node)
    workflow.add_node("analyze_template", analyze_template_node)
    workflow.add_node("plan_slides", plan_slides_node, retry=llm_retry)
    workflow.add_node("generate_spec", generate_spec_node, retry=llm_retry)
    workflow.add_node("validate", validate_node)
    workflow.add_node("render", render_node)
    workflow.add_node("error", error_node)

    # Add edges
    workflow.add_edge(START, "parse_md")
    workflow.add_edge("parse_md", "analyze_template")
    workflow.add_edge("analyze_template", "plan_slides")
    workflow.add_edge("plan_slides", "generate_spec")
    workflow.add_edge("generate_spec", "validate")
    workflow.add_conditional_edges("validate", should_render, {"render": "render", "error": "error"})
    workflow.add_edge("render", END)
    workflow.add_edge("error", END)

    return workflow.compile()


def run_pipeline(
    md_text: str,
    md_path: str = "",
    template_path: str = "",
    output_path: str = "",
    target_slide_count: int = 12,
) -> dict:
    """Run the full MD → PPTX pipeline. Returns final state dict."""
    pipeline = build_pipeline()

    initial_state: PipelineState = {
        "md_text": md_text,
        "md_path": md_path,
        "template_path": template_path,
        "output_path": output_path,
        "target_slide_count": target_slide_count,
        "errors": [],
        "warnings": [],
        "retry_count": 0,
    }

    result = pipeline.invoke(initial_state)
    return result
