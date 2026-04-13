#!/usr/bin/env python3
"""Batch generate PPTX presentations with individual folders for each markdown file."""
from __future__ import annotations
import argparse
import logging
import shutil
import subprocess
import sys
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn

# Add parent to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from src import config

console = Console()

def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s",
    )


def process_single(md_path: Path, output_dir: Path, template_path: str = "") -> bool:
    """Process a single markdown file into its folder. Returns True on success."""
    if not md_path.exists():
        console.print(f"[red]Error: File not found: {md_path}[/red]")
        return False

    # Create folder named after the file (without extension)
    folder_name = md_path.stem
    target_folder = output_dir / folder_name
    target_folder.mkdir(parents=True, exist_ok=True)

    # Copy markdown file to target folder
    target_md = target_folder / md_path.name
    shutil.copy2(md_path, target_md)

    # Determine output PPTX path
    target_pptx = target_folder / f"{md_path.stem}.pptx"

    console.print(Panel(
        f"[bold]{folder_name}[/bold]\n"
        f"Folder: {target_folder}",
        title="Processing",
        border_style="blue",
    ))

    # Run main.py with auto slide count (slides=0)
    cmd = [
        sys.executable,
        str(Path(__file__).parent.parent / "main.py"),
        "--input", str(target_md),
        "--output", str(target_pptx),
        "--slides", "0",  # Auto-detect slide count
    ]
    if template_path:
        cmd.extend(["--template", template_path])

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,  # 5 minute timeout per file
        )

        if result.returncode != 0:
            console.print(f"[red]Failed to generate PPTX[/red]")
            if result.stderr:
                console.print(f"[dim]{result.stderr[:500]}[/dim]")
            return False

        # Check if PPTX was created
        if target_pptx.exists():
            size_mb = target_pptx.stat().st_size / (1024 * 1024)
            console.print(f"[green]Success: {target_pptx.name} ({size_mb:.1f} MB)[/green]")
            return True
        else:
            console.print(f"[red]PPTX file not created[/red]")
            return False

    except subprocess.TimeoutExpired:
        console.print(f"[red]Timeout (5 min) exceeded[/red]")
        return False
    except Exception as e:
        console.print(f"[red]Error: {e}[/red]")
        return False


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Batch generate PPTX files with individual folders",
    )
    parser.add_argument(
        "--input-dir", "-i",
        help="Directory containing markdown files",
        default=str(config.PROJECT_ROOT / "test_cases"),
    )
    parser.add_argument(
        "--output-dir", "-o",
        help="Output directory for generated folders",
        default=str(config.PROJECT_ROOT / "generated_presentations"),
    )
    parser.add_argument(
        "--template", "-t",
        help="Path to .pptx template (auto-detected if omitted)",
        default="",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging",
    )

    args = parser.parse_args()
    setup_logging(args.verbose)

    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)

    if not input_dir.exists():
        console.print(f"[red]Input directory not found: {input_dir}[/red]")
        sys.exit(1)

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    # Find all markdown files
    md_files = sorted(input_dir.glob("*.md"))
    if not md_files:
        console.print(f"[yellow]No .md files found in {input_dir}[/yellow]")
        sys.exit(0)

    console.print(Panel(
        f"[bold]Batch processing {len(md_files)} files[/bold]\n"
        f"Input: {input_dir}\n"
        f"Output: {output_dir}",
        title="Batch Mode",
        border_style="cyan",
    ))

    success = 0
    failed = 0

    for i, md_file in enumerate(md_files, 1):
        console.print(f"\n[dim]── [{i}/{len(md_files)}] ──[/dim]")
        ok = process_single(md_file, output_dir, args.template)
        if ok:
            success += 1
        else:
            failed += 1

    console.print(f"\n[bold]Batch complete: {success} success, {failed} failed[/bold]")

    if failed > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
