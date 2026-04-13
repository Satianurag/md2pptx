#!/usr/bin/env python3
"""
Capture screenshots of all slides from PPTX files in the output folder.
Uses LibreOffice headless to convert PPTX → PDF, then pdftoppm to convert PDF → PNG.
"""

import argparse
import subprocess
import sys
from pathlib import Path
from typing import List


def convert_pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> bool:
    """Convert PPTX to PDF using LibreOffice headless."""
    print(f"Converting {pptx_path.name} to PDF...")
    
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(pdf_path.parent),
        str(pptx_path)
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            print(f"LibreOffice conversion failed: {result.stderr}")
            return False
        
        # LibreOffice creates PDF with the same base name as the PPTX file + .pdf extension
        expected_pdf = pdf_path.parent / (pptx_path.stem + ".pdf")
        
        if not expected_pdf.exists():
            print(f"PDF file not found at expected path: {expected_pdf}")
            print(f"Files in directory: {list(pdf_path.parent.glob('*'))}")
            return False
        
        # Rename to the desired PDF path
        if expected_pdf != pdf_path:
            expected_pdf.rename(pdf_path)
        
        return True
        
    except subprocess.TimeoutExpired:
        print("LibreOffice conversion timed out")
        return False
    except Exception as e:
        print(f"Error during PPTX to PDF conversion: {e}")
        return False


def convert_pdf_to_png(pdf_path: Path, output_dir: Path, dpi: int = 300) -> bool:
    """Convert PDF pages to PNG using pdftoppm."""
    print(f"Converting PDF to PNG slides at {dpi} DPI...")
    
    cmd = [
        "pdftoppm",
        "-png",
        "-r", str(dpi),
        str(pdf_path),
        str(output_dir / "slide")
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120
        )
        
        if result.returncode != 0:
            print(f"pdftoppm conversion failed: {result.stderr}")
            return False
        
        return True
        
    except subprocess.TimeoutExpired:
        print("pdftoppm conversion timed out")
        return False
    except Exception as e:
        print(f"Error during PDF to PNG conversion: {e}")
        return False


def rename_slide_files(output_dir: Path) -> int:
    """Rename pdftoppm output files to slide_001.png, slide_002.png, etc."""
    slide_files = sorted(output_dir.glob("slide-*.png"))
    
    for idx, old_path in enumerate(slide_files, start=1):
        new_name = f"slide_{idx:03d}.png"
        new_path = output_dir / new_name
        old_path.rename(new_path)
        print(f"  Created: {new_name}")
    
    return len(slide_files)


def process_presentation(pptx_path: Path, screenshots_dir: Path) -> bool:
    """Process a single PPTX file: convert to slides."""
    presentation_name = pptx_path.stem
    output_folder = screenshots_dir / presentation_name
    
    # Skip if folder already exists
    if output_folder.exists():
        print(f"Skipping {presentation_name} (screenshots already exist)")
        return True
    
    # Create output folder
    output_folder.mkdir(parents=True, exist_ok=True)
    
    # Temporary PDF path
    temp_pdf = output_folder / "temp.pdf"
    
    try:
        # Step 1: PPTX → PDF
        if not convert_pptx_to_pdf(pptx_path, temp_pdf):
            return False
        
        # Step 2: PDF → PNG
        if not convert_pdf_to_png(temp_pdf, output_folder):
            return False
        
        # Step 3: Rename files
        slide_count = rename_slide_files(output_folder)
        
        print(f"✓ Generated {slide_count} slides for {presentation_name}")
        
        return True
        
    finally:
        # Clean up temporary PDF
        if temp_pdf.exists():
            temp_pdf.unlink()


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Capture screenshots of PPTX slides")
    parser.add_argument("files", nargs="*", help="Specific PPTX files to process (if not provided, processes all files in output/)")
    args = parser.parse_args()
    
    project_root = Path(__file__).parent.parent
    screenshots_dir = project_root / "screenshots"
    
    # Create screenshots directory
    screenshots_dir.mkdir(exist_ok=True)
    
    # Determine which files to process
    if args.files:
        pptx_files = [Path(f) for f in args.files]
    else:
        output_dir = project_root / "output"
        pptx_files = sorted(output_dir.glob("*.pptx"))
    
    if not pptx_files:
        print("No PPTX files to process")
        sys.exit(0)
    
    print(f"Found {len(pptx_files)} PPTX file(s) to process\n")
    
    # Process each presentation
    success_count = 0
    for pptx_path in pptx_files:
        if not pptx_path.exists():
            print(f"File not found: {pptx_path}")
            continue
            
        print(f"\n{'='*60}")
        print(f"Processing: {pptx_path.name}")
        print('='*60)
        
        if process_presentation(pptx_path, screenshots_dir):
            success_count += 1
    
    print(f"\n{'='*60}")
    print(f"Complete: {success_count}/{len(pptx_files)} presentations processed")
    print(f"Screenshots saved to: {screenshots_dir}")
    print('='*60)


if __name__ == "__main__":
    main()
