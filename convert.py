#!/usr/bin/env python3
"""
PDF to DOCX Converter - Command Line Interface

A high-fidelity PDF to DOCX converter that preserves formatting, images, and tables.

Usage:
    python convert.py input.pdf                    # Convert single file
    python convert.py input.pdf -o output.docx     # Specify output path
    python convert.py input.pdf --pages 0,1,2      # Convert specific pages
    python convert.py *.pdf                        # Convert multiple files
    python convert.py --dir /path/to/pdfs          # Convert directory

Author: PDF to DOCX Converter
License: MIT
"""

import os
import sys
import argparse
from pathlib import Path
from typing import List, Optional

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

from pdf_converter import PDFConverter


def parse_pages(pages_str: str) -> List[int]:
    """
    Parse page specification string into a list of page numbers.

    Supports formats:
    - Single pages: "0,1,2,5"
    - Ranges: "0-5"
    - Mixed: "0,1,3-5,7"

    Args:
        pages_str: String specification of pages.

    Returns:
        List of page numbers (0-indexed).
    """
    pages = []

    for part in pages_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))

    return sorted(set(pages))


def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def print_banner():
    """Print application banner."""
    banner = """
╔══════════════════════════════════════════════════════════╗
║           PDF to DOCX Converter v1.0.0                   ║
║     High-Fidelity Conversion with Format Preservation    ║
╚══════════════════════════════════════════════════════════╝
"""
    print(banner)


def print_pdf_info(converter: PDFConverter, pdf_path: str) -> None:
    """Print detailed information about a PDF file."""
    try:
        info = converter.get_pdf_info(pdf_path)
        print(f"\nPDF Information:")
        print(f"  File: {info['path']}")
        print(f"  Pages: {info['page_count']}")

        if info['metadata']:
            meta = info['metadata']
            if meta.get('title'):
                print(f"  Title: {meta['title']}")
            if meta.get('author'):
                print(f"  Author: {meta['author']}")
            if meta.get('creator'):
                print(f"  Creator: {meta['creator']}")

        # Count images
        total_images = sum(p['image_count'] for p in info['pages'])
        if total_images:
            print(f"  Total Images: {total_images}")

        print()
    except Exception as e:
        print(f"Could not read PDF info: {e}")


def convert_single_file(
    converter: PDFConverter,
    pdf_path: str,
    output_path: Optional[str],
    pages: Optional[List[int]],
    start_page: int,
    end_page: Optional[int],
    show_info: bool
) -> bool:
    """Convert a single PDF file."""
    if show_info:
        print_pdf_info(converter, pdf_path)

    result = converter.convert(
        pdf_path,
        output_path,
        pages=pages,
        start_page=start_page,
        end_page=end_page
    )

    if result.success:
        output_size = os.path.getsize(result.output_path)
        print(f"\n[SUCCESS] Converted successfully!")
        print(f"  Output: {result.output_path}")
        print(f"  Pages: {result.pages_converted}")
        print(f"  Size: {format_size(output_size)}")

        if result.warnings:
            print(f"  Warnings:")
            for warning in result.warnings:
                print(f"    - {warning}")

        return True
    else:
        print(f"\n[ERROR] Conversion failed!")
        print(f"  Error: {result.error_message}")
        return False


def convert_multiple_files(
    converter: PDFConverter,
    pdf_paths: List[str],
    output_dir: Optional[str]
) -> bool:
    """Convert multiple PDF files."""
    print(f"\nBatch converting {len(pdf_paths)} files...")

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        print(f"Output directory: {output_dir}")

    if HAS_TQDM:
        progress_bar = tqdm(pdf_paths, desc="Converting", unit="file")
    else:
        progress_bar = pdf_paths

    successful = 0
    failed = 0

    for pdf_path in progress_bar:
        if output_dir:
            docx_name = Path(pdf_path).with_suffix('.docx').name
            docx_path = os.path.join(output_dir, docx_name)
        else:
            docx_path = None

        result = converter.convert(pdf_path, docx_path)

        if result.success:
            successful += 1
        else:
            failed += 1
            if not HAS_TQDM:
                print(f"  [FAILED] {pdf_path}: {result.error_message}")

    print(f"\n{'='*50}")
    print(f"Batch Conversion Complete")
    print(f"  Successful: {successful}")
    print(f"  Failed: {failed}")
    print(f"  Total: {len(pdf_paths)}")
    print(f"{'='*50}")

    return failed == 0


def convert_directory(
    converter: PDFConverter,
    input_dir: str,
    output_dir: Optional[str],
    recursive: bool
) -> bool:
    """Convert all PDFs in a directory."""
    print(f"\nScanning directory: {input_dir}")
    print(f"Recursive: {recursive}")

    results = converter.convert_directory(
        input_dir,
        output_dir,
        recursive=recursive
    )

    if not results:
        print("No PDF files found in directory.")
        return True

    successful = sum(1 for r in results if r.success)
    failed = len(results) - successful

    print(f"\n{'='*50}")
    print(f"Directory Conversion Complete")
    print(f"  Successful: {successful}")
    print(f"  Failed: {failed}")
    print(f"  Total: {len(results)}")
    print(f"{'='*50}")

    # List failed files
    if failed > 0:
        print("\nFailed files:")
        for r in results:
            if not r.success:
                print(f"  - {r.input_path}: {r.error_message}")

    return failed == 0


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Convert PDF files to DOCX format with high fidelity.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.pdf                     Convert single file
  %(prog)s document.pdf -o output.docx      Specify output filename
  %(prog)s document.pdf --pages 0,1,2       Convert specific pages (0-indexed)
  %(prog)s document.pdf --pages 0-5         Convert page range
  %(prog)s *.pdf                            Convert multiple files
  %(prog)s --dir /path/to/pdfs              Convert all PDFs in directory
  %(prog)s --dir /path/to/pdfs -r           Convert recursively
  %(prog)s document.pdf --info              Show PDF info before converting

Features:
  - Preserves text formatting (font, size, bold, italic, color)
  - Maintains image positioning and quality
  - Recreates table structures accurately
  - Handles multi-column layouts
  - Supports batch processing
        """
    )

    # Input options
    parser.add_argument(
        'input_files',
        nargs='*',
        help='PDF file(s) to convert'
    )

    parser.add_argument(
        '-d', '--dir',
        dest='input_dir',
        help='Directory containing PDF files to convert'
    )

    # Output options
    parser.add_argument(
        '-o', '--output',
        dest='output_path',
        help='Output file path (for single file) or directory (for batch)'
    )

    # Page selection
    parser.add_argument(
        '-p', '--pages',
        help='Specific pages to convert (e.g., "0,1,2" or "0-5" or "0,2-4,6")'
    )

    parser.add_argument(
        '--start',
        type=int,
        default=0,
        help='Start page (0-indexed, default: 0)'
    )

    parser.add_argument(
        '--end',
        type=int,
        default=None,
        help='End page (0-indexed, exclusive, default: last page)'
    )

    # Processing options
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Process subdirectories recursively (with --dir)'
    )

    parser.add_argument(
        '--info',
        action='store_true',
        help='Show PDF information before converting'
    )

    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress progress output'
    )

    parser.add_argument(
        '-v', '--version',
        action='version',
        version='%(prog)s 1.0.0'
    )

    args = parser.parse_args()

    # Validate arguments
    if not args.input_files and not args.input_dir:
        parser.print_help()
        print("\nError: Please provide input file(s) or use --dir to specify a directory.")
        sys.exit(1)

    # Print banner unless quiet mode
    if not args.quiet:
        print_banner()

    # Create converter
    converter = PDFConverter(verbose=not args.quiet)

    # Parse pages if specified
    pages = None
    if args.pages:
        try:
            pages = parse_pages(args.pages)
            if not args.quiet:
                print(f"Selected pages: {pages}")
        except ValueError as e:
            print(f"Error parsing pages: {e}")
            sys.exit(1)

    # Process based on input type
    success = True

    if args.input_dir:
        # Directory mode
        success = convert_directory(
            converter,
            args.input_dir,
            args.output_path,
            args.recursive
        )
    elif len(args.input_files) == 1:
        # Single file mode
        success = convert_single_file(
            converter,
            args.input_files[0],
            args.output_path,
            pages,
            args.start,
            args.end,
            args.info
        )
    else:
        # Multiple files mode
        success = convert_multiple_files(
            converter,
            args.input_files,
            args.output_path
        )

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
