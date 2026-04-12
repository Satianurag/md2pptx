#!/usr/bin/env python3
"""MD → PPTX: Convert markdown research reports into professional PowerPoint presentations."""
from __future__ import annotations
import argparse
import logging
import sys
import time
from pathlib import Path

from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn
from dotenv import load_dotenv

load_dotenv()

from src.agent import run_pipeline
from src import config

console = Console()


def _auto_slide_count(file_size_bytes: int, md_text: str | None = None) -> int:
    """Pick a target slide count (10-15) using content profiling when available,
    falling back to file-size heuristic otherwise."""
    if md_text:
        try:
            from src.markdown_parser import parse_markdown
            from src.content_profiler import profile_content
            tree = parse_markdown(md_text)
            prof = profile_content(tree)
            # Base: 4 structural slides (cover + agenda + conclusion + thank_you)
            base = 4
            # Add 1 slide per section, capped at 9 content slides
            section_slides = min(prof.total_sections, 9)
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

    kb = file_size_bytes / 1024
    if kb < 30:
        return 10
    elif kb < 100:
        return 11
    elif kb < 300:
        return 12
    elif kb < 600:
        return 13
    elif kb < 1500:
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
    target_slides: int = 12,
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

    console.print(Panel(
        f"[bold]{md_file.name}[/bold]\n"
        f"Size: {size_mb:.2f} MB | Target slides: {target_slides}",
        title="Processing",
        border_style="blue",
    ))

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
    target_slides: int = 12,
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
        # Auto-calculate slide count per file if target_slides is 0 (auto mode)
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
        help="Path to markdown file (or directory for --batch)",
        required=True,
    )
    parser.add_argument(
        "--template", "-t",
        help="Path to .pptx template (auto-detected if omitted)",
        default="",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output .pptx path (auto-generated if omitted)",
        default="",
    )
    parser.add_argument(
        "--slides", "-s",
        type=int,
        default=12,
        help="Target slide count (10-15, default: 12). Use 0 for auto-detection based on file size.",
    )
    parser.add_argument(
        "--batch", "-b",
        action="store_true",
        help="Process all .md files in the input directory",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()

    setup_logging(args.verbose)

    # Clamp slide count (0 = auto mode for batch)
    if args.slides != 0:
        args.slides = max(config.MIN_SLIDES, min(config.MAX_SLIDES, args.slides))

    # Ensure output directory exists
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if args.batch:
        process_batch(args.input, args.template, args.slides)
    else:
        ok = process_single(args.input, args.template, args.output, args.slides)
        sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
