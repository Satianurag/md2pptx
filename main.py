#!/usr/bin/env python3
"""MD → PPTX: Convert markdown research reports into professional PowerPoint presentations."""
from __future__ import annotations
import argparse
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# Reconfigure stdout/stderr to UTF-8 BEFORE any rich imports so the Braille
# spinner characters emitted by Rich's Progress don't crash the Windows
# cp1252 codec on ``__exit__``.
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from dotenv import load_dotenv

# encoding='utf-8-sig' strips a BOM if present (saves common-mistake debugging)
load_dotenv(encoding="utf-8-sig")

from src.agent import run_pipeline
from src import config

# legacy_windows=False forces Rich to use ANSI escape codes rather than the
# Win32 console API, which avoids cp1252 charmap_encode crashes on Unicode
# spinner glyphs.
console = Console(legacy_windows=False)


def _resolve_presenter(cli_value: str = "") -> str:
    """Resolve presenter name with fallback chain.

    Priority: (1) CLI ``--presenter`` arg, (2) ``PRESENTER_NAME`` env var,
    (3) Windows ``os.getlogin()``, (4) empty string.
    """
    cli_value = (cli_value or "").strip()
    if cli_value:
        return cli_value
    env_val = os.environ.get("PRESENTER_NAME", "").strip()
    if env_val:
        return env_val
    try:
        user = os.getlogin()
        if user:
            return user
    except Exception:
        pass
    return ""


def _resolve_date(cli_value: str = "") -> str:
    """Resolve the cover slide date string. CLI override > today's date."""
    cli_value = (cli_value or "").strip()
    if cli_value:
        return cli_value
    return datetime.now().strftime("%B %d, %Y")


def _auto_slide_count(file_size_bytes: int, md_text: str | None = None) -> int:
    """Pick a target slide count (10-15) using content profiling when available,
    falling back to file-size heuristic otherwise.

    Default bias: 15 slides (2026 standard).  Only go lower if content is genuinely thin.
    """
    if md_text:
        try:
            from src.markdown_parser import parse_markdown
            from src.content_profiler import profile_content
            tree = parse_markdown(md_text)
            prof = profile_content(tree)
            # Base: 4 structural slides (cover + agenda + conclusion + thank_you)
            base = 4
            # Add 1 slide per section, capped at 11 content slides
            section_slides = min(prof.total_sections, 11)
            # Bonus slides for data-rich content
            bonus = 0
            if prof.total_tables >= 3:
                bonus += 1
            if prof.total_metrics >= 5:
                bonus += 1
            target = base + section_slides + bonus
            return max(config.MIN_SLIDES, min(config.MAX_SLIDES, target))
        except Exception:
            pass  # fall back to size-based

    # File-size heuristic — biased toward 15
    kb = file_size_bytes / 1024
    if kb < 15:
        return 10
    elif kb < 30:
        return 12
    elif kb < 60:
        return 13
    elif kb < 150:
        return 14
    else:
        return 15


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s",
        handlers=[RichHandler(console=console, show_time=False, show_path=False)],
    )
    # Suppress noisy libraries
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("httpcore").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("google").setLevel(logging.WARNING)


def process_single(
    md_path: str,
    template_path: str = "",
    output_path: str = "",
    target_slides: int = 15,
    presenter: str = "",
    date_str: str = "",
) -> bool:
    """Process a single markdown file. Returns True on success."""
    md_file = Path(md_path)
    if not md_file.exists():
        console.print(f"[red]Error: File not found: {md_path}[/red]")
        return False

    # Check file size (warn but don't reject — hackathon requires all test cases)
    size_mb = md_file.stat().st_size / (1024 * 1024)
    if size_mb > 5:
        console.print(f"[yellow]Warning: Large file ({size_mb:.1f}MB). Aggressive chunking will be applied.[/yellow]")

    panel_lines = [
        f"[bold]{md_file.name}[/bold]",
        f"Size: {size_mb:.2f} MB | Target slides: {target_slides}",
    ]
    if presenter or date_str:
        panel_lines.append(f"Presenter: {presenter or '(empty)'} | Date: {date_str or '(empty)'}")
    console.print(Panel("\n".join(panel_lines), title="Processing", border_style="blue"))

    md_text = md_file.read_text(encoding="utf-8", errors="replace")

    start = time.time()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        task = progress.add_task("Running pipeline...", total=None)

        result = run_pipeline(
            md_text=md_text,
            md_path=str(md_file),
            template_path=template_path,
            output_path=output_path,
            target_slide_count=target_slides,
            presenter=presenter,
            date_str=date_str,
        )

        progress.update(task, description="Done")

    elapsed = time.time() - start
    errors = result.get("errors", [])
    warnings = result.get("warnings", [])
    out = result.get("output_path", "")

    if errors:
        console.print(f"[red]Failed ({elapsed:.1f}s)[/red]")
        for e in errors:
            console.print(f"  [red]• {e}[/red]")
        return False

    console.print(f"[green]Success ({elapsed:.1f}s)[/green]")
    if out:
        console.print(f"  Output: [bold]{out}[/bold]")
    if warnings:
        for w in warnings:
            console.print(f"  [yellow]⚠ {w}[/yellow]")

    return True


def process_batch(
    input_dir: str,
    template_path: str = "",
    target_slides: int = 15,
    presenter: str = "",
    date_str: str = "",
) -> None:
    """Process all .md files in a directory."""
    md_dir = Path(input_dir)
    if not md_dir.is_dir():
        console.print(f"[red]Error: Not a directory: {input_dir}[/red]")
        return

    md_files = sorted(md_dir.glob("*.md"))
    if not md_files:
        console.print(f"[yellow]No .md files found in {input_dir}[/yellow]")
        return

    console.print(Panel(
        f"[bold]Batch processing {len(md_files)} files[/bold]",
        title="Batch Mode",
        border_style="cyan",
    ))

    success = 0
    failed = 0

    for i, md_file in enumerate(md_files, 1):
        console.print(f"\n[dim]── [{i}/{len(md_files)}] ──[/dim]")
        out_path = str(config.OUTPUT_DIR / f"{md_file.stem}.pptx")
        if target_slides > 0:
            slides = target_slides
        else:
            md_content = md_file.read_text(encoding="utf-8", errors="replace")
            slides = _auto_slide_count(md_file.stat().st_size, md_content)
        ok = process_single(
            md_path=str(md_file),
            template_path=template_path,
            output_path=out_path,
            target_slides=slides,
            presenter=presenter,
            date_str=date_str,
        )
        if ok:
            success += 1
        else:
            failed += 1

    console.print(f"\n[bold]Batch complete: {success} success, {failed} failed[/bold]")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert markdown research reports to PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input", "-i",
        help="Path to markdown file to convert.",
        required=True,
    )
    parser.add_argument(
        "--template", "-t",
        help="Path to a .pptx template file. Required.",
        required=True,
    )
    parser.add_argument(
        "--output", "-o",
        help="Output .pptx path (auto-generated if omitted)",
        default="",
    )
    parser.add_argument(
        "--slides", "-s",
        type=int,
        default=15,
        help="Number of slides to generate (10–15, default: 15).",
    )

    parser.add_argument(
        "--presenter", "-p",
        help="Presenter name for the cover slide (fallback: PRESENTER_NAME env var, then OS user).",
        default="",
    )
    parser.add_argument(
        "--date",
        help="Date string for the cover slide (default: today in 'Month DD, YYYY' format).",
        default="",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    setup_logging(args.verbose)

    # Validate slide count
    if not (config.MIN_SLIDES <= args.slides <= config.MAX_SLIDES):
        parser.error(
            f"--slides must be between {config.MIN_SLIDES} and {config.MAX_SLIDES} "
            f"(got {args.slides}). Default is {config.DEFAULT_SLIDE_COUNT}."
        )

    # Ensure output directory exists
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    presenter = _resolve_presenter(args.presenter)
    date_str = _resolve_date(args.date)

    ok = process_single(
        args.input, args.template, args.output, args.slides,
        presenter=presenter, date_str=date_str,
    )
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
