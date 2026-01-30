#!/usr/bin/env python3
"""
PDF to DOCX Converter - Advanced Command Line Interface

A high-fidelity PDF to DOCX converter with:
- Multi-column layout preservation
- Structure analysis and validation
- Post-processing cleanup
- Content comparison

Usage:
    python convert.py input.pdf                     # Standard conversion
    python convert.py input.pdf --advanced          # Advanced layout preservation
    python convert.py input.pdf --analyze           # Analyze PDF structure only
    python convert.py input.pdf --validate          # Convert and validate
    python convert.py input.pdf --cleanup           # Convert and cleanup output

Author: PDF to DOCX Converter v2.0
License: MIT
"""

import os
import sys
import argparse
from pathlib import Path
from typing import List, Optional

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from pdf_converter import PDFConverter, AdvancedPDFConverter
from pdf_converter.analyzer import PDFAnalyzer, LayoutType
from pdf_converter.postprocessor import PostProcessor, ConversionValidator
from pdf_converter.advanced_converter import ConversionQuality


def print_banner():
    """Print application banner."""
    banner = """
╔══════════════════════════════════════════════════════════════╗
║          PDF to DOCX Converter v2.0 - Advanced               ║
║   Multi-Column Layout Preservation & Content Validation      ║
╚══════════════════════════════════════════════════════════════╝
"""
    print(banner)


def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def analyze_pdf(pdf_path: str, detailed: bool = False):
    """Analyze PDF structure and print report."""
    print(f"\nAnalyzing: {pdf_path}")
    print("-" * 60)

    with PDFAnalyzer(pdf_path) as analyzer:
        structure = analyzer.analyze()

        print(f"Total Pages: {structure.page_count}")
        print(f"Dominant Layout: {structure.dominant_layout.value}")
        print(f"Consistent Layout: {'Yes' if structure.has_consistent_layout else 'No (mixed layouts)'}")

        # Layout distribution
        layout_counts = {}
        for page in structure.pages:
            layout = page.layout_type.value
            layout_counts[layout] = layout_counts.get(layout, 0) + 1

        print(f"\nLayout Distribution:")
        for layout, count in sorted(layout_counts.items(), key=lambda x: -x[1]):
            pct = count / structure.page_count * 100
            print(f"  {layout}: {count} pages ({pct:.0f}%)")

        # Detailed per-page analysis
        if detailed:
            print(f"\nDetailed Page Analysis:")
            for page in structure.pages:
                print(f"\n  Page {page.page_num + 1}:")
                print(f"    Layout: {page.layout_type.value}")
                print(f"    Columns: {len(page.columns)}")
                print(f"    Text Blocks: {len(page.text_blocks)}")
                print(f"    Images: {len(page.images)}")

                if page.columns and len(page.columns) > 1:
                    print(f"    Column Widths: {[f'{c.width:.0f}pt' for c in page.columns]}")

        # Summary statistics
        total_text_blocks = sum(len(p.text_blocks) for p in structure.pages)
        total_images = sum(len(p.images) for p in structure.pages)
        multi_col_pages = sum(1 for p in structure.pages if len(p.columns) > 1)

        print(f"\nSummary:")
        print(f"  Total Text Blocks: {total_text_blocks}")
        print(f"  Total Images: {total_images}")
        print(f"  Multi-Column Pages: {multi_col_pages}")

        # Recommendations
        print(f"\nRecommendations:")
        if multi_col_pages > 0:
            print(f"  - Use --advanced mode for best multi-column preservation")
        if total_images > 0:
            print(f"  - Images will be extracted and positioned in output")
        if not structure.has_consistent_layout:
            print(f"  - Mixed layouts detected - consider page-by-page review")


def convert_standard(
    pdf_path: str,
    output_path: Optional[str],
    pages: Optional[List[int]],
    quiet: bool
) -> bool:
    """Standard conversion using pdf2docx."""
    converter = PDFConverter(verbose=not quiet)
    result = converter.convert(pdf_path, output_path, pages=pages)

    if result.success:
        if not quiet:
            output_size = os.path.getsize(result.output_path)
            print(f"\n[SUCCESS] Standard conversion complete!")
            print(f"  Output: {result.output_path}")
            print(f"  Pages: {result.pages_converted}")
            print(f"  Size: {format_size(output_size)}")
        return True
    else:
        print(f"\n[ERROR] Conversion failed: {result.error_message}")
        return False


def convert_advanced(
    pdf_path: str,
    output_path: Optional[str],
    pages: Optional[List[int]],
    quality: str,
    validate: bool,
    cleanup: bool,
    quiet: bool
) -> bool:
    """Advanced conversion with layout preservation."""
    quality_map = {
        "draft": ConversionQuality.DRAFT,
        "standard": ConversionQuality.STANDARD,
        "high": ConversionQuality.HIGH
    }

    converter = AdvancedPDFConverter(quality=quality_map.get(quality, ConversionQuality.HIGH))

    try:
        out_path, stats, validation = converter.convert(
            pdf_path,
            output_path,
            pages=pages,
            validate=validate
        )

        if not quiet:
            output_size = os.path.getsize(out_path)
            print(f"\n{'='*60}")
            print(f"CONVERSION COMPLETE")
            print(f"{'='*60}")
            print(f"Output: {out_path}")
            print(f"Size: {format_size(output_size)}")
            print(f"\nStatistics:")
            print(f"  Pages Processed: {stats.pages_processed}")
            print(f"  Text Blocks: {stats.text_blocks_converted}")
            print(f"  Images Extracted: {stats.images_extracted}")
            print(f"  Columns Preserved: {stats.columns_preserved}")

            if stats.warnings:
                print(f"\nWarnings:")
                for warning in stats.warnings:
                    print(f"  - {warning}")

            if validation:
                status = "PASSED" if validation.is_valid else "NEEDS REVIEW"
                print(f"\nValidation: {status}")
                print(f"  Text Match: {validation.text_match_ratio:.1%}")
                if validation.issues:
                    print(f"  Issues:")
                    for issue in validation.issues:
                        print(f"    - {issue}")
                if validation.suggestions:
                    print(f"  Suggestions:")
                    for suggestion in validation.suggestions:
                        print(f"    - {suggestion}")

        # Post-processing cleanup if requested
        if cleanup:
            print(f"\nRunning post-processing cleanup...")
            post_processor = PostProcessor(verbose=not quiet)
            cleanup_result = post_processor.cleanup_document(out_path)

            if not quiet and cleanup_result.success:
                print(f"  Blank paragraphs removed: {cleanup_result.blank_paragraphs_removed}")
                print(f"  Spacing issues fixed: {cleanup_result.duplicate_spaces_fixed}")

        return True

    except Exception as e:
        print(f"\n[ERROR] Advanced conversion failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def validate_existing(pdf_path: str, docx_path: str, quiet: bool) -> bool:
    """Validate an existing conversion."""
    print(f"\nValidating conversion...")
    print(f"  PDF: {pdf_path}")
    print(f"  DOCX: {docx_path}")

    validator = ConversionValidator()
    report = validator.full_validation(pdf_path, docx_path)

    if not quiet:
        validator.print_validation_report(report)

    return report["valid"]


def compare_content(pdf_path: str, docx_path: str, quiet: bool) -> bool:
    """Compare content between PDF and DOCX."""
    post_processor = PostProcessor(verbose=not quiet)
    comparison = post_processor.compare_content(pdf_path, docx_path)

    if not quiet:
        post_processor.print_comparison_report(comparison)

    return comparison.match_percentage >= 85


def cleanup_docx(docx_path: str, output_path: Optional[str], quiet: bool) -> bool:
    """Clean up an existing DOCX file."""
    print(f"\nCleaning up: {docx_path}")

    post_processor = PostProcessor(verbose=not quiet)
    result = post_processor.cleanup_document(docx_path, output_path)

    if not quiet:
        print(f"\nCleanup Results:")
        print(f"  Original paragraphs: {result.original_paragraphs}")
        print(f"  Final paragraphs: {result.final_paragraphs}")
        print(f"  Blank removed: {result.blank_paragraphs_removed}")
        print(f"  Spacing fixed: {result.duplicate_spaces_fixed}")

        if result.issues_fixed:
            print(f"\nIssues Fixed:")
            for issue in result.issues_fixed:
                print(f"  - {issue}")

    return result.success


def parse_pages(pages_str: str) -> List[int]:
    """Parse page specification string."""
    pages = []
    for part in pages_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))
    return sorted(set(pages))


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Convert PDF files to DOCX with high fidelity and layout preservation.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.pdf                      Standard conversion
  %(prog)s document.pdf --advanced           Advanced layout preservation
  %(prog)s document.pdf --analyze            Analyze PDF structure only
  %(prog)s document.pdf --analyze --detailed Full structure analysis
  %(prog)s document.pdf --validate           Convert and validate output
  %(prog)s document.pdf --cleanup            Convert and cleanup output
  %(prog)s document.pdf --compare doc.docx   Compare PDF with existing DOCX

Modes:
  Standard:  Uses pdf2docx library (fast, good for simple layouts)
  Advanced:  Uses custom analyzer with multi-column preservation
             (slower, better for complex documents)

Quality Levels (for --advanced):
  draft:     Fast conversion, basic formatting
  standard:  Balanced quality and speed
  high:      Maximum fidelity (default)
        """
    )

    # Input
    parser.add_argument(
        'input_file',
        nargs='?',
        help='PDF file to convert'
    )

    # Output
    parser.add_argument(
        '-o', '--output',
        dest='output_path',
        help='Output file path'
    )

    # Mode selection
    mode_group = parser.add_argument_group('Conversion Modes')
    mode_group.add_argument(
        '--advanced',
        action='store_true',
        help='Use advanced converter with layout preservation'
    )
    mode_group.add_argument(
        '--analyze',
        action='store_true',
        help='Analyze PDF structure only (no conversion)'
    )
    mode_group.add_argument(
        '--detailed',
        action='store_true',
        help='Show detailed analysis (with --analyze)'
    )

    # Validation and cleanup
    val_group = parser.add_argument_group('Validation & Cleanup')
    val_group.add_argument(
        '--validate',
        action='store_true',
        help='Validate conversion against original PDF'
    )
    val_group.add_argument(
        '--cleanup',
        action='store_true',
        help='Run post-processing cleanup on output'
    )
    val_group.add_argument(
        '--compare',
        metavar='DOCX',
        help='Compare PDF with existing DOCX file'
    )
    val_group.add_argument(
        '--cleanup-only',
        metavar='DOCX',
        help='Cleanup existing DOCX file (no conversion)'
    )

    # Quality settings
    parser.add_argument(
        '--quality',
        choices=['draft', 'standard', 'high'],
        default='high',
        help='Conversion quality (default: high)'
    )

    # Page selection
    parser.add_argument(
        '-p', '--pages',
        help='Specific pages to convert (e.g., "0,1,2" or "0-5")'
    )

    # Other options
    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Suppress progress output'
    )
    parser.add_argument(
        '-v', '--version',
        action='version',
        version='%(prog)s 2.0.0'
    )

    args = parser.parse_args()

    # Handle cleanup-only mode
    if args.cleanup_only:
        if not args.quiet:
            print_banner()
        success = cleanup_docx(args.cleanup_only, args.output_path, args.quiet)
        sys.exit(0 if success else 1)

    # Require input file for other modes
    if not args.input_file:
        parser.print_help()
        print("\nError: Please provide an input PDF file.")
        sys.exit(1)

    pdf_path = args.input_file

    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    if not args.quiet:
        print_banner()

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

    success = True

    # Handle different modes
    if args.analyze:
        # Analysis mode
        analyze_pdf(pdf_path, args.detailed)

    elif args.compare:
        # Compare mode
        success = compare_content(pdf_path, args.compare, args.quiet)

    elif args.advanced:
        # Advanced conversion
        success = convert_advanced(
            pdf_path,
            args.output_path,
            pages,
            args.quality,
            args.validate,
            args.cleanup,
            args.quiet
        )

    else:
        # Standard conversion
        success = convert_standard(pdf_path, args.output_path, pages, args.quiet)

        # Optional validation after standard conversion
        if success and args.validate:
            output_path = args.output_path or str(Path(pdf_path).with_suffix('.docx'))
            validate_existing(pdf_path, output_path, args.quiet)

        # Optional cleanup after standard conversion
        if success and args.cleanup:
            output_path = args.output_path or str(Path(pdf_path).with_suffix('.docx'))
            cleanup_docx(output_path, None, args.quiet)

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
