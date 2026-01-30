#!/usr/bin/env python3
"""
PDF to DOCX Converter

Convert PDF to DOCX with accurate layout preservation.

How it works:
1. Claude Vision analyzes page layout (columns, structure) - NOT content
2. Python (PyMuPDF) extracts all text and images
3. DOCX is generated matching the original structure

Usage:
    python convert.py document.pdf
    python convert.py document.pdf -o output.docx

Set ANTHROPIC_API_KEY for layout analysis, or it will use default single-column.
"""

import os
import sys
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to DOCX with layout preservation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python convert.py document.pdf              Convert to document.docx
  python convert.py input.pdf -o output.docx  Specify output path

Environment:
  ANTHROPIC_API_KEY  Set for better multi-column layout detection
                     (Without it, uses basic single-column layout)
        """
    )

    parser.add_argument('pdf_file', help='PDF file to convert')
    parser.add_argument('-o', '--output', help='Output DOCX path')
    parser.add_argument('-k', '--api-key', help='Anthropic API key')
    parser.add_argument('-v', '--version', action='version', version='4.0.0')

    args = parser.parse_args()

    if not os.path.exists(args.pdf_file):
        print(f"Error: File not found: {args.pdf_file}")
        sys.exit(1)

    # Import here for faster startup on errors
    from pdf_converter import convert

    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")

    if not api_key:
        print("\nNote: No ANTHROPIC_API_KEY set.")
        print("      Layout analysis disabled - using basic conversion.")
        print("      Set the key for better multi-column support.\n")

    result = convert(args.pdf_file, args.output, api_key)

    if result.success:
        print(f"\nOutput: {result.output_path}")
        sys.exit(0)
    else:
        print(f"\nConversion failed:")
        for err in result.errors:
            print(f"  - {err}")
        sys.exit(1)


if __name__ == '__main__':
    main()
