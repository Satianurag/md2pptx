from pathlib import Path
from pptx.util import Inches, Pt, Emu

# --- Project paths ---
PROJECT_ROOT = Path(__file__).parent.parent
TEMPLATES_DIR = PROJECT_ROOT.parent / "Code EZ_ Master of Agents _ Files" / "Slide Master"
TEST_CASES_DIR = PROJECT_ROOT.parent / "Code EZ_ Master of Agents _ Files" / "Test Cases"
OUTPUT_DIR = PROJECT_ROOT / "output"

# --- Slide dimensions (standard 16:9 widescreen in EMU) ---
SLIDE_WIDTH = 12192000   # 13.333 inches
SLIDE_HEIGHT = 6858000   # 7.5 inches

# --- Margins & safe zone (EMU) ---
MARGIN_LEFT = Emu(457200)      # 0.50 inches — generous professional margin
MARGIN_RIGHT = Emu(457200)
MARGIN_TOP = Emu(600075)       # ~0.66 inches
MARGIN_BOTTOM = Emu(457200)    # 0.50 inches — room for footer

CONTENT_LEFT = MARGIN_LEFT
CONTENT_TOP = Emu(1300000)     # below title area (extra breathing room)
CONTENT_WIDTH = SLIDE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
CONTENT_HEIGHT = SLIDE_HEIGHT - CONTENT_TOP - MARGIN_BOTTOM

# --- Typography (Pt) ---
FONT_TITLE = Pt(28)
FONT_SUBTITLE = Pt(16)
FONT_BODY = Pt(13)
FONT_CAPTION = Pt(10)
FONT_LABEL = Pt(11)

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
SHAPE_GAP = Emu(180000)       # gap between shapes (slightly more generous)
SHAPE_PADDING = Emu(120000)   # internal padding

# --- Max content per slide ---
MAX_BULLETS_PER_SLIDE = 5
MAX_CHARS_PER_BULLET = 160
MAX_CHARS_PER_BULLET_CARD = 120
MAX_TEXT_LINES = 8
MAX_CARD_ITEMS = 6

# --- Brand-aligned brightness levels for card backgrounds ---
# Cards use ACCENT_1 with varying brightness to stay brand-consistent
CARD_BRIGHTNESS_LEVELS = [0.85, 0.90, 0.92, 0.88, 0.86, 0.91]
