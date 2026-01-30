#!/usr/bin/env python3
"""
PDF to DOCX Converter - Powered by Claude Vision

A single-command PDF to DOCX converter that:
1. Analyzes each page visually using Claude Vision AI
2. Understands exact layout, columns, text formatting
3. Recreates the document structure accurately in DOCX
4. Validates output automatically

Usage:
    python convert.py document.pdf
    python convert.py document.pdf -o output.docx
    python convert.py document.pdf --pages 0-5

Environment:
    ANTHROPIC_API_KEY - Your Anthropic API key for Claude Vision

Author: PDF to DOCX Converter v3.0
"""

import os
import sys
import argparse
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))


def print_banner():
    """Print application banner."""
    print("""
╔══════════════════════════════════════════════════════════════════╗
║        PDF to DOCX Converter v3.0 - Claude Vision Powered        ║
║                                                                  ║
║  Intelligent conversion with AI-powered layout analysis          ║
║  Multi-column • Tables • Images • Exact formatting               ║
╚══════════════════════════════════════════════════════════════════╝
""")


def parse_pages(pages_str: str) -> list:
    """Parse page specification string into list of page numbers."""
    pages = []
    for part in pages_str.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))
    return sorted(set(pages))


def check_api_key() -> bool:
    """Check if Anthropic API key is available."""
    if os.environ.get("ANTHROPIC_API_KEY"):
        return True

    print("\n[ERROR] ANTHROPIC_API_KEY environment variable not set!")
    print("\nTo use this converter, you need an Anthropic API key:")
    print("  1. Get your API key from https://console.anthropic.com/")
    print("  2. Set the environment variable:")
    print("     export ANTHROPIC_API_KEY='your-key-here'")
    print("\nThe API key is used ONLY for visual analysis of PDF pages,")
    print("not for any data extraction or storage.")
    return False


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Convert PDF to DOCX with Claude Vision-powered layout analysis",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.pdf                    Convert with auto-naming
  %(prog)s document.pdf -o output.docx     Specify output file
  %(prog)s document.pdf --pages 0-5        Convert specific pages

The converter automatically:
  • Analyzes page layout using Claude Vision AI
  • Detects columns, headers, footers, tables
  • Preserves text formatting and positioning
  • Handles multi-column layouts accurately
  • Extracts and positions images
  • Validates conversion quality

Requires ANTHROPIC_API_KEY environment variable.
        """
    )

    parser.add_argument(
        'pdf_file',
        help='PDF file to convert'
    )

    parser.add_argument(
        '-o', '--output',
        dest='output_path',
        help='Output DOCX file path (default: same name as PDF)'
    )

    parser.add_argument(
        '-p', '--pages',
        help='Pages to convert (e.g., "0,1,2" or "0-5"). Default: all pages'
    )

    parser.add_argument(
        '-k', '--api-key',
        dest='api_key',
        help='Anthropic API key (alternative to environment variable)'
    )

    parser.add_argument(
        '-q', '--quiet',
        action='store_true',
        help='Minimal output'
    )

    parser.add_argument(
        '-v', '--version',
        action='version',
        version='%(prog)s 3.0.0'
    )

    args = parser.parse_args()

    # Print banner unless quiet
    if not args.quiet:
        print_banner()

    # Check input file
    if not os.path.exists(args.pdf_file):
        print(f"[ERROR] File not found: {args.pdf_file}")
        sys.exit(1)

    # Check API key
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        if not check_api_key():
            sys.exit(1)

    # Parse pages if specified
    pages = None
    if args.pages:
        try:
            pages = parse_pages(args.pages)
            if not args.quiet:
                print(f"Converting pages: {pages}")
        except ValueError as e:
            print(f"[ERROR] Invalid page specification: {e}")
            sys.exit(1)

    # Import converter (delayed to allow early error messages)
    try:
        from pdf_converter.intelligent_converter import convert_pdf
    except ImportError as e:
        print(f"[ERROR] Failed to import converter: {e}")
        print("\nMake sure all dependencies are installed:")
        print("  pip install -r requirements.txt")
        sys.exit(1)

    # Convert
    try:
        result = convert_pdf(
            args.pdf_file,
            args.output_path,
            api_key=api_key
        )

        if result.success:
            if not args.quiet:
                print(f"\n[SUCCESS] Conversion complete!")
                print(f"Output saved to: {result.output_path}")
            sys.exit(0)
        else:
            print(f"\n[ERROR] Conversion failed")
            for note in result.notes:
                print(f"  - {note}")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\n[CANCELLED] Conversion interrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()
