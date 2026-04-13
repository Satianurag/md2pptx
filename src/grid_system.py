"""Grid alignment system for snapping slide elements to consistent positions."""
from __future__ import annotations
import logging
from .schemas import Position, SlideMasterInfo
from . import config

logger = logging.getLogger(__name__)

# Minimum gap between adjacent elements (EMU) — per Guidelines §5 "breathing space"
_MIN_GAP = min(int(config.SHAPE_GAP), 140000)  # cap at ~0.15in to avoid excessive whitespace
# Footer zone: bottom 12 % of slide is reserved for footers / slide number
_FOOTER_ZONE_RATIO = 0.12


class Grid:
    """Provides layout presets based on slide dimensions and margins.

    Never import this at module level with hardcoded config values.
    Use the factory methods ``Grid.default()`` or ``Grid.from_template()``.
    """

    def __init__(
        self,
        slide_w: int,
        slide_h: int,
        m_left: int,
        m_right: int,
        m_top: int,
        m_bottom: int,
        content_top: int,
    ):
        self.slide_w = slide_w
        self.slide_h = slide_h
        self.m_left = m_left
        self.m_right = m_right
        self.m_top = m_top
        self.m_bottom = m_bottom
        self.content_top = content_top

        self.content_w = slide_w - m_left - m_right
        self.content_h = slide_h - content_top - m_bottom

    # ── Factory methods ──────────────────────────────────────────────

    @classmethod
    def default(cls) -> "Grid":
        """Grid using the 16:9 defaults from config (no template)."""
        return cls(
            slide_w=config.SLIDE_WIDTH,
            slide_h=config.SLIDE_HEIGHT,
            m_left=int(config.MARGIN_LEFT),
            m_right=int(config.MARGIN_RIGHT),
            m_top=int(config.MARGIN_TOP),
            m_bottom=int(config.MARGIN_BOTTOM),
            content_top=int(config.DEFAULT_CONTENT_TOP),
        )

    @classmethod
    def from_template(cls, master_info: SlideMasterInfo) -> "Grid":
        """Build a Grid that adapts to the template's actual dimensions.

        Content top is inferred from the 'title_only' layout placeholder
        positions (title bottom + gap).  Content bottom is inferred from
        footer placeholders or the default margin.
        """
        slide_w = master_info.slide_width or config.SLIDE_WIDTH
        slide_h = master_info.slide_height or config.SLIDE_HEIGHT

        m_left = int(config.MARGIN_LEFT)
        m_right = int(config.MARGIN_RIGHT)
        m_top = int(config.MARGIN_TOP)

        # --- Derive content_top from title placeholder ---
        content_top = int(config.DEFAULT_CONTENT_TOP)
        footer_top = slide_h  # assume no footer

        for layout in master_info.layouts:
            if layout.category in ("title_only", "title_content"):
                for ph in layout.placeholders:
                    ph_type = (ph.ph_type or "").upper()
                    # Title / subtitle — content starts below the lowest title placeholder
                    if ph_type in ("TITLE", "CENTER_TITLE", "SUBTITLE"):
                        bottom = ph.top + ph.height
                        candidate = bottom + 100000  # 100k EMU gap
                        if candidate > content_top:
                            content_top = candidate
                    # Footer / slide-number — content must end above the highest footer
                    if ph_type in ("SLIDE_NUMBER", "FOOTER", "DATE_TIME", "BODY"):
                        # "BODY" placeholders in the bottom 15 % of slide are footers
                        if ph.top > slide_h * 0.80:
                            if ph.top < footer_top:
                                footer_top = ph.top
                break  # use only the first matching layout

        # Determine bottom margin from footer zone
        if footer_top < slide_h:
            m_bottom = slide_h - footer_top + 50000  # 50k EMU padding above footer
        else:
            m_bottom = int(config.MARGIN_BOTTOM)

        logger.debug(
            f"Grid.from_template: slide={slide_w}x{slide_h}, "
            f"content_top={content_top}, m_bottom={m_bottom}"
        )
        return cls(
            slide_w=slide_w, slide_h=slide_h,
            m_left=m_left, m_right=m_right,
            m_top=m_top, m_bottom=m_bottom,
            content_top=content_top,
        )

    # ── Single-zone presets ──────────────────────────────────────────

    def full(self) -> Position:
        """Full content area below the title."""
        return Position(
            left=self.m_left, top=self.content_top,
            width=self.content_w, height=self.content_h,
        )

    def chart(self) -> Position:
        """Slightly inset area optimized for charts."""
        inset = min(250000, self.content_w // 20)
        return Position(
            left=self.m_left + inset,
            top=self.content_top + 120000,
            width=self.content_w - 2 * inset,
            height=self.content_h - 240000,
        )

    def table(self) -> Position:
        """Full-width area optimized for tables."""
        return Position(
            left=self.m_left,
            top=self.content_top + 100000,
            width=self.content_w,
            height=self.content_h - 100000,
        )

    # ── Two-zone presets ─────────────────────────────────────────────

    def two_column(self) -> tuple[Position, Position]:
        """Left-right split with gap."""
        half = (self.content_w - _MIN_GAP) // 2
        left = Position(
            left=self.m_left, top=self.content_top,
            width=half, height=self.content_h,
        )
        right = Position(
            left=self.m_left + half + _MIN_GAP, top=self.content_top,
            width=half, height=self.content_h,
        )
        return left, right

    def top_bottom(self) -> tuple[Position, Position]:
        """Top-bottom split with gap."""
        half_h = (self.content_h - _MIN_GAP) // 2
        top = Position(
            left=self.m_left, top=self.content_top,
            width=self.content_w, height=half_h,
        )
        bottom = Position(
            left=self.m_left, top=self.content_top + half_h + _MIN_GAP,
            width=self.content_w, height=half_h,
        )
        return top, bottom

    def sidebar_main(self, sidebar_ratio: float = 0.3) -> tuple[Position, Position]:
        """Narrow sidebar + wide main area."""
        sidebar_w = int(self.content_w * sidebar_ratio) - _MIN_GAP // 2
        main_w = self.content_w - sidebar_w - _MIN_GAP
        sidebar = Position(
            left=self.m_left, top=self.content_top,
            width=sidebar_w, height=self.content_h,
        )
        main = Position(
            left=self.m_left + sidebar_w + _MIN_GAP, top=self.content_top,
            width=main_w, height=self.content_h,
        )
        return sidebar, main

    # ── Multi-zone presets ───────────────────────────────────────────

    def three_column(self) -> tuple[Position, Position, Position]:
        """Three equal columns with gaps."""
        col_w = (self.content_w - 2 * _MIN_GAP) // 3
        cols = []
        for i in range(3):
            cols.append(Position(
                left=self.m_left + i * (col_w + _MIN_GAP),
                top=self.content_top,
                width=col_w, height=self.content_h,
            ))
        return cols[0], cols[1], cols[2]

    def grid_2x2(self) -> tuple[Position, Position, Position, Position]:
        """2×2 grid with gaps."""
        half_w = (self.content_w - _MIN_GAP) // 2
        half_h = (self.content_h - _MIN_GAP) // 2
        positions = []
        for row in range(2):
            for col in range(2):
                positions.append(Position(
                    left=self.m_left + col * (half_w + _MIN_GAP),
                    top=self.content_top + row * (half_h + _MIN_GAP),
                    width=half_w, height=half_h,
                ))
        return positions[0], positions[1], positions[2], positions[3]

    def top_wide_bottom_split(self) -> tuple[Position, Position, Position]:
        """Wide top zone + two bottom columns (for KPI + chart combos)."""
        top_h = int(self.content_h * 0.35)
        bottom_h = self.content_h - top_h - _MIN_GAP
        half_w = (self.content_w - _MIN_GAP) // 2
        top = Position(
            left=self.m_left, top=self.content_top,
            width=self.content_w, height=top_h,
        )
        bl = Position(
            left=self.m_left, top=self.content_top + top_h + _MIN_GAP,
            width=half_w, height=bottom_h,
        )
        br = Position(
            left=self.m_left + half_w + _MIN_GAP,
            top=self.content_top + top_h + _MIN_GAP,
            width=half_w, height=bottom_h,
        )
        return top, bl, br

    def n_cards(self, n: int, max_cols: int = 4, max_rows: int = 2) -> list[Position]:
        """Evenly-spaced card grid for *n* items (used by KPI, comparison, etc.).

        Caps layout to *max_rows* rows to prevent vertical congestion.
        """
        cols = min(n, max_cols)
        rows = min((n + cols - 1) // cols, max_rows)
        # Re-cap n so we don't render cards that won't fit
        n = min(n, cols * rows)

        gap_h = _MIN_GAP
        gap_v = _MIN_GAP
        card_w = max(
            (self.content_w - gap_h * max(cols - 1, 0)) // max(cols, 1),
            1500000,  # minimum card width ~1.65 in
        )
        card_h = min(
            max(
                (self.content_h - gap_v * max(rows - 1, 0)) // max(rows, 1),
                900000,  # minimum card height ~1.0 in
            ),
            int(self.content_h * 0.75) if rows == 1 else self.content_h // 2,
        )
        positions: list[Position] = []
        for idx in range(n):
            col = idx % cols
            row = idx // cols
            positions.append(Position(
                left=self.m_left + col * (card_w + gap_h),
                top=self.content_top + row * (card_h + gap_v),
                width=card_w,
                height=card_h,
            ))
        return positions
