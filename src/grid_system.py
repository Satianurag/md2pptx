"""Grid alignment system for snapping slide elements to consistent positions."""
from __future__ import annotations
from .schemas import Position
from . import config


# Minimum gap between adjacent elements (EMU)
_MIN_GAP = 120000


class Grid:
    """Provides layout presets based on slide dimensions and margins."""

    def __init__(
        self,
        slide_w: int = config.SLIDE_WIDTH,
        slide_h: int = config.SLIDE_HEIGHT,
        m_left: int = int(config.MARGIN_LEFT),
        m_right: int = int(config.MARGIN_RIGHT),
        m_top: int = int(config.MARGIN_TOP),
        m_bottom: int = int(config.MARGIN_BOTTOM),
        content_top: int = int(config.CONTENT_TOP),
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

    # ── Single-zone presets ──────────────────────────────────────────

    def full(self) -> Position:
        """Full content area below the title."""
        return Position(
            left=self.m_left, top=self.content_top,
            width=self.content_w, height=self.content_h,
        )

    def chart(self) -> Position:
        """Slightly inset area optimized for charts."""
        inset = 200000
        return Position(
            left=self.m_left + inset,
            top=self.content_top + 100000,
            width=self.content_w - 2 * inset,
            height=self.content_h - 200000,
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


# Module-level singleton
grid = Grid()
