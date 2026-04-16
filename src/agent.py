from __future__ import annotations
import logging
from pathlib import Path
from typing import TypedDict, Optional, Any

from langgraph.graph import StateGraph, START, END

from .schemas import (
    ContentTree, SlideMasterInfo, SlidePlan, PresentationSpec, DeckContent,
)
from .markdown_parser import parse_markdown
from .slide_master import read_slide_master, auto_detect_template
from .slide_planner import plan_slides
from .spec_generator import generate_presentation_spec
from .content_writer import write_deck_content, fix_deck_content
from .validator import validate_and_fix, ValidationResult
from .pptx_renderer import render_presentation
from .content_chunker import chunk_content_tree
from .content_profiler import profile_content, ContentProfile
from . import config

logger = logging.getLogger(__name__)


# ── Pipeline state ──

class PipelineState(TypedDict, total=False):
    md_text: str
    md_path: str
    content_tree: Optional[ContentTree]
    content_profile: Optional[ContentProfile]
    template_path: str
    master_info: Optional[SlideMasterInfo]
    slide_plan: Optional[SlidePlan]
    deck_content: Optional[DeckContent]
    presentation_spec: Optional[PresentationSpec]
    output_path: str
    target_slide_count: int
    errors: list[str]
    warnings: list[str]
    fix_attempts: int
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


def profile_content_node(state: PipelineState) -> dict:
    """Analyse content tree and produce a ContentProfile for adaptive decisions."""
    logger.info("Node: profile_content")
    content_tree = state.get("content_tree")
    if not content_tree:
        return {}
    profile = profile_content(content_tree)
    logger.info(f"Profile: archetype={profile.archetype}, data_richness={profile.data_richness}, "
                f"vis_ratio={profile.recommended_visual_ratio:.1f}, "
                f"charts={profile.recommended_chart_types}, infographics={profile.recommended_infographic_types}")
    return {"content_profile": profile}


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
    """Use LLM to generate slide plan — AI decides narrative dynamically."""
    logger.info("Node: plan_slides")

    content_tree = state.get("content_tree")
    if not content_tree:
        return {"errors": state.get("errors", []) + ["No content_tree available for planning"]}

    master_info = state.get("master_info")
    target = state.get("target_slide_count", config.DEFAULT_SLIDE_COUNT)
    content_profile = state.get("content_profile")

    slide_plan = plan_slides(content_tree, master_info, target, content_profile)
    logger.info(f"Plan: {len(slide_plan.slides)} slides, storyline: {slide_plan.storyline_summary[:80]}...")
    return {"slide_plan": slide_plan}


def write_content_node(state: PipelineState) -> dict:
    """Use LLM to write presentation-ready content for all slides."""
    logger.info("Node: write_content")

    slide_plan = state.get("slide_plan")
    content_tree = state.get("content_tree")
    content_profile = state.get("content_profile")

    if not slide_plan or not content_tree:
        return {"errors": state.get("errors", []) + ["Missing slide_plan or content_tree for content writing"]}

    deck_content = write_deck_content(slide_plan, content_tree, content_profile)
    logger.info(f"Content written: {len(deck_content.slides)} slides")
    return {"deck_content": deck_content}


def generate_spec_node(state: PipelineState) -> dict:
    """Generate full PresentationSpec from plan + AI-written content."""
    logger.info("Node: generate_spec")

    content_tree = state.get("content_tree")
    slide_plan = state.get("slide_plan")
    master_info = state.get("master_info")
    template_path = state.get("template_path", "")
    deck_content = state.get("deck_content")

    if not content_tree or not slide_plan:
        return {"errors": state.get("errors", []) + ["Missing content_tree or slide_plan"]}

    content_profile = state.get("content_profile")

    spec = generate_presentation_spec(
        content_tree=content_tree,
        slide_plan=slide_plan,
        master_info=master_info,
        template_path=template_path,
        content_profile=content_profile,
        deck_content=deck_content,
    )
    logger.info(f"Spec generated: {len(spec.slides)} slides")
    return {"presentation_spec": spec}


def validate_node(state: PipelineState) -> dict:
    """Validate the presentation spec."""
    logger.info("Node: validate")

    spec = state.get("presentation_spec")
    if not spec:
        return {"errors": state.get("errors", []) + ["No presentation_spec to validate"]}

    content_profile = state.get("content_profile")
    master_info = state.get("master_info")
    sw = master_info.slide_width if master_info else None
    sh = master_info.slide_height if master_info else None
    result = validate_and_fix(spec, content_profile, slide_width=sw, slide_height=sh, master_info=master_info)

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


# ── Conditional edges ──

def should_render_or_fix(state: PipelineState) -> str:
    """After validation: render if pass, fix_content if fixable, error if fatal."""
    errors = state.get("errors", [])
    if errors:
        return "error"

    vr = state.get("validation_result")
    if vr and not vr.passed:
        # Allow 1 fix attempt
        if state.get("fix_attempts", 0) < 1:
            return "fix_content"
        return "error"

    return "render"


def fix_content_node(state: PipelineState) -> dict:
    """Ask AI to fix content issues flagged by validation."""
    logger.info("Node: fix_content")

    deck_content = state.get("deck_content")
    vr = state.get("validation_result")

    if not deck_content or not vr:
        return {"errors": state.get("errors", []) + ["No content/validation to fix"]}

    issues = vr.errors + vr.warnings
    fixed = fix_deck_content(deck_content, issues)

    return {
        "deck_content": fixed,
        "fix_attempts": state.get("fix_attempts", 0) + 1,
    }


def error_node(state: PipelineState) -> dict:
    """Handle pipeline errors."""
    errors = state.get("errors", [])
    logger.error(f"Pipeline failed with {len(errors)} errors: {errors}")
    return state


# ── Build the graph ──

def build_pipeline() -> StateGraph:
    """Build and compile the LangGraph pipeline.

    Graph:
      START → parse_md → ┌─ profile_content ─┐
                        │                    ├→ plan_slides → write_content → generate_spec → validate
                        └─ analyze_template ─┘
      validate → render | fix_content → generate_spec | error
    """
    workflow = StateGraph(PipelineState)

    # Add nodes
    workflow.add_node("parse_md", parse_md_node)
    workflow.add_node("profile_content", profile_content_node)
    workflow.add_node("analyze_template", analyze_template_node)
    workflow.add_node("plan_slides", plan_slides_node)
    workflow.add_node("write_content", write_content_node)
    workflow.add_node("generate_spec", generate_spec_node)
    workflow.add_node("validate", validate_node)
    workflow.add_node("fix_content", fix_content_node)
    workflow.add_node("render", render_node)
    workflow.add_node("error", error_node)

    # Edges: parse_md fans out to profile + template (parallel)
    workflow.add_edge(START, "parse_md")
    workflow.add_edge("parse_md", "profile_content")
    workflow.add_edge("parse_md", "analyze_template")

    # Both must complete before planning (fan-in)
    workflow.add_edge("profile_content", "plan_slides")
    workflow.add_edge("analyze_template", "plan_slides")

    # Sequential: plan → write → spec → validate
    workflow.add_edge("plan_slides", "write_content")
    workflow.add_edge("write_content", "generate_spec")
    workflow.add_edge("generate_spec", "validate")

    # Conditional: pass → render, fail → fix_content, fatal → error
    workflow.add_conditional_edges("validate", should_render_or_fix, {
        "render": "render",
        "fix_content": "fix_content",
        "error": "error",
    })

    # Fix loop: fix_content → generate_spec (re-enters validate)
    workflow.add_edge("fix_content", "generate_spec")

    workflow.add_edge("render", END)
    workflow.add_edge("error", END)

    return workflow.compile()


def run_pipeline(
    md_text: str,
    md_path: str = "",
    template_path: str = "",
    output_path: str = "",
    target_slide_count: int = 15,
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
        "fix_attempts": 0,
    }

    result = pipeline.invoke(initial_state)
    return result
