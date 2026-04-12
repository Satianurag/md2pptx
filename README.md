# MD2PPTX — Markdown to PowerPoint AI Pipeline

Convert markdown research reports into professional, visually rich PowerPoint presentations using an AI-powered agentic pipeline built with LangGraph and Gemini.

## Architecture

```
Markdown → Parser → Chunker → [AI Planner] → [AI Spec Gen] → Validator → Renderer → .pptx
             │                    │               │            │            │
        mistune 3.2         Gemini Flash     Rule-based   Auto-fix   python-pptx
                            Lite (LLM 1)     + LLM 2     + clamp     + theme colors
```

**LangGraph StateGraph Pipeline:**

| Node | Module | Description |
|------|--------|-------------|
| `parse_md` | `markdown_parser.py` | Mistune 3.x AST → structured ContentTree |
| `analyze_template` | `slide_master.py` | Auto-detect template, extract theme colors & layouts |
| `plan_slides` | `slide_planner.py` | Gemini AI generates storyline + slide types (with rule-based fallback) |
| `generate_spec` | `spec_generator.py` | Rule-based + Gemini AI generates full PresentationSpec |
| `validate` | `validator.py` | Content density, bounds, chart/table validation & auto-fix |
| `render` | `pptx_renderer.py` | Rich multi-shape PPTX with card bullets, infographics, slide furniture |

## Setup

```bash
# 1. Clone and enter directory
cd md2pptx

# 2. Create virtual environment
python -m venv ../venv
source ../venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Set API key
cp .env.example .env
# Edit .env and add your GOOGLE_API_KEY
```

## Usage

### Single file
```bash
python main.py --input path/to/report.md
```

### With template and custom slide count
```bash
python main.py --input report.md --template template.pptx --slides 14
```

### Batch processing
```bash
python main.py --input path/to/test_cases/ --batch
```

### Options
| Flag | Description | Default |
|------|-------------|---------|
| `--input, -i` | Markdown file or directory (batch) | Required |
| `--template, -t` | .pptx template path | Auto-detect |
| `--output, -o` | Output .pptx path | `output/<name>.pptx` |
| `--slides, -s` | Target slide count (10-15) | 12 |
| `--batch, -b` | Process all .md files in directory | Off |
| `--verbose, -v` | Verbose logging | Off |

## Key Design Decisions

1. **2-3 LLM calls per presentation** — minimizes API usage under strict rate limits (15 calls/min, 250k tokens/min, 500 calls/day)
2. **Rule-based first, AI second** — cover/agenda/exec_summary/conclusion slides generated deterministically; only complex content uses LLM. Full rule-based fallback when LLM is unavailable.
3. **Infographic-first approach** — automatic pattern detection (process flows, timelines, comparisons, KPIs) converts bullets into visual infographics before falling back to text
4. **Card-based bullet rendering** — each bullet point rendered as a visual card with numbered icon circle, giving 15-25+ shapes per content slide (vs 2-3 for plain text)
5. **Slide furniture** — every content slide gets accent underline, footer bar (deck title + slide number), and left accent stripe for consistent professional look
6. **Tiered content chunking** — standard/moderate/aggressive chunking tiers handle files from <5MB to 26MB+ while preserving key structure
7. **Grid alignment system** — 9 preset layouts (full, chart, table, two_column, three_column, grid_2x2, etc.) for consistent element positioning
8. **Theme color inheritance** — all fills use MSO_THEME_COLOR (schemeClr XML) when template loaded, inheriting slide master palette. Fallback hex only without template.
9. **Chart-first detection** — auto-detects chartable table data and upgrades bullet slides to mixed/chart; guarantees at least 1 chart per presentation if numeric data exists
10. **Fuzzy template matching** — word-overlap scoring matches markdown filenames to templates
11. **Template Bookend System** — when a template has ≥2 slides, the first and last are treated as fixed bookends. The cover slide gets only title/subtitle filled into placeholders (no accent bars or extra shapes). The closing slide is added with zero modifications, preserving baked-in designs like "Thank You!". Layout **index** (not name) is used for selection because templates can have duplicate layout names pointing to different designs. Content slides are prevented from using the closing layout via `excluded_idx`.
12. **LangGraph RetryPolicy** — exponential backoff (2s→30s) with jitter on LLM nodes for transient failures
13. **Error isolation** — per-element try/except in renderer; validation warnings instead of blocking errors

## Tech Stack

| Package | Version | Purpose |
|---------|---------|---------|
| langgraph | 1.1.6 | Agent pipeline orchestration |
| langchain-google-genai | 4.2.1 | Gemini 3.1 Flash Lite integration |
| python-pptx | 1.0.2 | PowerPoint generation |
| mistune | 3.2.0 | Markdown AST parsing |
| pydantic | 2.12.5 | Data validation & structured output |
| rich | 13.0+ | CLI output formatting |

## Visual Quality Features

- **Card-based bullets**: numbered icon circles + card backgrounds (3 shapes/bullet → 18+ shapes for 6 bullets)
- **Infographic types**: process flow (chevrons + arrows), timeline (circles + alternating labels), comparison cards, KPI metric cards, hierarchy with connectors
- **Slide furniture**: accent underline under title, footer bar with deck title + slide number, left accent stripe
- **Theme-aware coloring**: all shapes use template theme colors (MSO_THEME_COLOR) for consistency
- **Smart chart type selection**: pie for single-series, line for time-series, bar for many categories
- **Content density**: auto-trim bullets/text to prevent overcrowding (max 4 elements, 800 chars)

## Project Structure

```
md2pptx/
├── main.py                # CLI entry point (single + batch mode)
├── requirements.txt       # Dependencies (all latest versions)
├── .env.example           # API key template
├── output/                # Generated .pptx files
├── test_cases/            # 24 markdown research reports for testing
├── templates/             # 3 PPTX slide master templates
├── scripts/
│   ├── capture_slides.py  # PPTX → PDF → PNG screenshot tool
│   └── verify_output.py   # Automated quality checks (contrast, overflow, structure)
└── src/
    ├── agent.py           # LangGraph StateGraph + RetryPolicy + rule-based fallback
    ├── color_utils.py     # WCAG 2.1 contrast utilities, text color picker
    ├── components.py      # Card-based bullet rendering, accent shapes
    ├── config.py          # Constants, margins, typography, slide limits
    ├── content_chunker.py # Tiered chunking (standard/moderate/aggressive)
    ├── content_profiler.py # Content archetype detection (financial, narrative, etc.)
    ├── drawingml_effects.py # Shadows, gradients, rounded corners, card styling
    ├── grid_system.py     # 9-preset grid alignment system
    ├── llm.py             # Gemini wrapper + rate limiter + structured output
    ├── markdown_parser.py # Mistune markdown → ContentTree
    ├── pptx_renderer.py   # Rich PPTX renderer (bookends, cards, infographics)
    ├── schemas.py         # 15+ Pydantic data models
    ├── slide_master.py    # Template reading, theme extraction, fuzzy matching
    ├── slide_planner.py   # AI slide storyline planning
    ├── spec_generator.py  # Rule-based + AI spec gen, infographic-first, chart-first
    └── validator.py       # Content density, bounds, chart/table auto-fix
```
