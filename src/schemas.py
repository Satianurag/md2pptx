from __future__ import annotations
from typing import Optional, Literal, Union
from pydantic import BaseModel, Field


# ── Markdown parsing output ──────────────────────────────────────────

class DataTable(BaseModel):
    title: Optional[str] = None
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str]] = Field(default_factory=list)
    alignments: list[Optional[str]] = Field(default_factory=list)

class KeyMetric(BaseModel):
    label: str
    value: str
    unit: Optional[str] = None

class ContentSection(BaseModel):
    heading: str
    level: int = 1
    text: str = ""
    bullets: list[str] = Field(default_factory=list)
    tables: list[DataTable] = Field(default_factory=list)
    metrics: list[KeyMetric] = Field(default_factory=list)
    subsections: list[ContentSection] = Field(default_factory=list)

class ContentTree(BaseModel):
    title: str = ""
    subtitle: str = ""
    sections: list[ContentSection] = Field(default_factory=list)
    executive_summary: str = ""
    all_tables: list[DataTable] = Field(default_factory=list)
    all_metrics: list[KeyMetric] = Field(default_factory=list)


# ── Position / layout primitives ─────────────────────────────────────

class Position(BaseModel):
    left: int      # EMU
    top: int       # EMU
    width: int     # EMU
    height: int    # EMU


# ── Slide element types ──────────────────────────────────────────────

class TextContent(BaseModel):
    text: str
    font_size: Optional[int] = None        # Pt value
    bold: bool = False
    italic: bool = False
    color: Optional[str] = None            # hex e.g. "1F77B4"
    alignment: Optional[str] = None        # left, center, right

class BulletContent(BaseModel):
    items: list[str]
    font_size: Optional[int] = None

class ChartContent(BaseModel):
    chart_type: Literal["bar", "column", "line", "pie", "area", "doughnut", "scatter"]
    title: Optional[str] = None
    categories: list[str] = Field(default_factory=list)
    series: list[ChartSeries] = Field(default_factory=list)

class ChartSeries(BaseModel):
    name: str
    values: list[float]

class TableContent(BaseModel):
    headers: list[str]
    rows: list[list[str]]
    col_widths: Optional[list[float]] = None   # proportional widths

class ShapeContent(BaseModel):
    shape_type: str                             # MSO_SHAPE name: "ROUNDED_RECTANGLE", "CHEVRON", etc.
    text: str = ""
    fill_color: Optional[str] = None            # hex
    line_color: Optional[str] = None
    font_size: Optional[int] = None
    bold: bool = False

class InfographicContent(BaseModel):
    infographic_type: Literal["process_flow", "timeline", "comparison", "kpi_cards", "hierarchy"]
    items: list[InfographicItem] = Field(default_factory=list)

class InfographicItem(BaseModel):
    title: str
    description: str = ""
    value: Optional[str] = None
    icon: Optional[str] = None                  # shape type for icon


# ── Slide element (union) ────────────────────────────────────────────

class SlideElement(BaseModel):
    element_type: Literal["text", "bullets", "chart", "table", "shape", "infographic", "textbox"]
    position: Position
    content: Union[TextContent, BulletContent, ChartContent, TableContent, ShapeContent, InfographicContent]


# ── Single slide specification ───────────────────────────────────────

class SlideSpec(BaseModel):
    slide_number: int
    slide_type: Literal[
        "cover", "agenda", "executive_summary", "section_divider",
        "content", "chart", "table", "infographic", "mixed", "conclusion", "thank_you"
    ]
    layout_name: str = "blank"                  # which template layout to use
    title: str = ""
    subtitle: Optional[str] = None
    elements: list[SlideElement] = Field(default_factory=list)


# ── Full presentation specification ──────────────────────────────────

class PresentationSpec(BaseModel):
    title: str
    subtitle: str = ""
    slides: list[SlideSpec] = Field(default_factory=list)
    template_path: str = ""
    target_slide_count: int = 12


# ── Slide master metadata ───────────────────────────────────────────

class ThemeColors(BaseModel):
    dk1: str = "000000"
    lt1: str = "FFFFFF"
    dk2: str = "44546A"
    lt2: str = "E7E6E6"
    accent1: str = "4472C4"
    accent2: str = "ED7D31"
    accent3: str = "A5A5A5"
    accent4: str = "FFC000"
    accent5: str = "5B9BD5"
    accent6: str = "70AD47"
    hlink: str = "0563C1"
    folHlink: str = "954F72"

    def accents(self) -> list[str]:
        return [self.accent1, self.accent2, self.accent3,
                self.accent4, self.accent5, self.accent6]

class PlaceholderInfo(BaseModel):
    idx: int
    name: str
    ph_type: Optional[str] = None
    left: int
    top: int
    width: int
    height: int

class LayoutInfo(BaseModel):
    index: int
    name: str
    category: Literal[
        "cover", "divider", "blank", "title_only",
        "title_content", "two_content", "thank_you", "other"
    ] = "other"
    placeholders: list[PlaceholderInfo] = Field(default_factory=list)

class SlideMasterInfo(BaseModel):
    template_path: str
    slide_width: int
    slide_height: int
    layouts: list[LayoutInfo] = Field(default_factory=list)
    theme_colors: ThemeColors = Field(default_factory=ThemeColors)


# ── Slide planning (AI planner output) ────────────────────────────

class SlidePlanItem(BaseModel):
    slide_number: int
    slide_type: Literal[
        "cover", "agenda", "executive_summary", "section_divider",
        "content", "chart", "table", "infographic", "mixed", "conclusion", "thank_you"
    ]
    title: str
    subtitle: Optional[str] = None
    content_source: list[str] = Field(default_factory=list)       # section heading refs from ContentTree
    visualization_hint: Literal["chart", "table", "infographic", "bullets", "mixed", "text", "kpi"] = "bullets"
    chart_type_hint: Optional[Literal["bar", "column", "line", "pie", "area", "doughnut"]] = None
    infographic_type_hint: Optional[Literal["process_flow", "timeline", "comparison", "kpi_cards", "hierarchy"]] = None
    key_message: str = ""

class SlidePlan(BaseModel):
    storyline_summary: str
    target_slide_count: int = 12
    slides: list[SlidePlanItem] = Field(default_factory=list)


# Fix forward references
ChartContent.model_rebuild()
