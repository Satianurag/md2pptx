from pathlib import Path
from pptx.util import Inches, Pt, Emu

# --- Project paths ---
PROJECT_ROOT = Path(__file__).parent.parent
OUTPUT_DIR = PROJECT_ROOT / "output"

_HACKATHON_FOLDER = "Code EZ_ Master of Agents _ Files"


def _find_hackathon_dir(subfolder: str) -> Path | None:
    """Search multiple candidate locations for the hackathon resource folder."""
    candidates = [
        # 1. Sibling of the project root (original expectation)
        PROJECT_ROOT.parent / _HACKATHON_FOLDER / subfolder,
        # 2. Inside the project root itself
        PROJECT_ROOT / _HACKATHON_FOLDER / subfolder,
        # 3. User Downloads folder
        Path.home() / "Downloads" / _HACKATHON_FOLDER / subfolder,
        # 4. User Documents folder
        Path.home() / "Documents" / _HACKATHON_FOLDER / subfolder,
        # 5. User home folder
        Path.home() / _HACKATHON_FOLDER / subfolder,
        # 6. Desktop
        Path.home() / "Desktop" / _HACKATHON_FOLDER / subfolder,
    ]
    for p in candidates:
        if p.exists() and p.is_dir():
            return p
    return None


def find_templates_dir() -> Path | None:
    return _find_hackathon_dir("Slide Master")


def find_test_cases_dir() -> Path | None:
    return _find_hackathon_dir("Test Cases")


# Eagerly resolve so downstream code can import the value directly
TEMPLATES_DIR: Path | None = find_templates_dir()
TEST_CASES_DIR: Path | None = find_test_cases_dir()

# --- Slide dimensions (standard 16:9 widescreen in EMU) ---
SLIDE_WIDTH = 12192000   # 13.333 inches
SLIDE_HEIGHT = 6858000   # 7.5 inches

# --- Margins & safe zone (EMU) ---
MARGIN_LEFT = Emu(457200)      # 0.50 inches — generous professional margin
MARGIN_RIGHT = Emu(457200)
MARGIN_TOP = Emu(600075)       # ~0.66 inches
MARGIN_BOTTOM = Emu(457200)    # 0.50 inches — room for footer

# Default content area (used by Grid when no template is loaded)
DEFAULT_CONTENT_LEFT = MARGIN_LEFT
DEFAULT_CONTENT_TOP = Emu(1300000)     # below title area
DEFAULT_CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
DEFAULT_CONTENT_HEIGHT = SLIDE_HEIGHT - DEFAULT_CONTENT_TOP - MARGIN_BOTTOM

# --- Typography (Pt) ---
FONT_TITLE = Pt(28)
FONT_SUBTITLE = Pt(16)
FONT_BODY = Pt(13)
FONT_CAPTION = Pt(10)
FONT_LABEL = Pt(11)

# --- Font Families ---
FONT_NAME_PRIMARY = "Inter"        # Primary sans-serif for all regular text
FONT_NAME_MONO = "JetBrains Mono"  # Monospace for code and technical content

# --- Text frame internal padding (EMU) ---
# Generous padding prevents text from touching shape edges
TF_MARGIN_LEFT = Emu(120000)    # ~0.13 inches
TF_MARGIN_RIGHT = Emu(120000)
TF_MARGIN_TOP = Emu(90000)      # ~0.10 inches
TF_MARGIN_BOTTOM = Emu(60000)

# --- Bullet / paragraph spacing (Pt) ---
BULLET_SPACE_BEFORE = Pt(10)    # breathing room between bullet items
PARA_SPACE_AFTER = Pt(4)        # subtle trailing space

# --- Slide count ---
MIN_SLIDES = 10
MAX_SLIDES = 15
DEFAULT_SLIDE_COUNT = 15

# --- Narrative roles (reference list — AI decides dynamically which to use) ---
NARRATIVE_ROLES: list[str] = [
    "cover",
    "agenda",
    "executive_summary",
    "market_landscape",
    "methodology",
    "key_findings",
    "data_evidence",
    "timeline_roadmap",
    "case_study",
    "regional_analysis",
    "challenges_risks",
    "recommendations",
    "impact_analysis",
    "conclusion",
    "thank_you",
]

# --- Layout name mapping (normalized) ---
LAYOUT_COVER = "cover"
LAYOUT_DIVIDER = "divider"
LAYOUT_BLANK = "blank"
LAYOUT_TITLE_ONLY = "title only"
LAYOUT_THANK_YOU = "thank"

# --- Chart defaults ---
CHART_LEFT = Emu(800000)
CHART_TOP = Emu(1500000)
CHART_WIDTH = Emu(10400000)
CHART_HEIGHT = Emu(4800000)

# --- Table defaults ---
TABLE_LEFT = Emu(500000)
TABLE_TOP = Emu(1500000)
TABLE_WIDTH = Emu(11000000)

# --- Shape spacing ---
SHAPE_GAP = Emu(200000)       # gap between shapes (~0.22in, breathing room per Guidelines §5)
SHAPE_PADDING = Emu(120000)   # internal padding

# --- Max content per slide ---
MAX_BULLETS_PER_SLIDE = 6
MAX_CHARS_PER_BULLET = 200
MAX_CHARS_PER_BULLET_CARD = 140
MAX_TEXT_LINES = 8
MAX_CARD_ITEMS = 5            # cap visual cards per slide to avoid congestion
MAX_PROCESS_ITEMS = 5         # process flow steps per slide
MAX_KPI_ITEMS = 4             # KPI metric cards per slide
MAX_INFOGRAPHIC_DESC = 100    # chars for card descriptions inside grids
MAX_CMP_DESC = 130            # chars for comparison card descriptions
MAX_CONCLUSION_CHARS = 280    # chars per bullet on conclusion slide

# --- Brand-aligned brightness levels for card backgrounds ---
# Cards use ACCENT_1 with varying brightness to stay brand-consistent
CARD_BRIGHTNESS_LEVELS = [0.85, 0.90, 0.92, 0.88, 0.86, 0.91]
