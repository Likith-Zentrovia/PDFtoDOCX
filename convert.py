#!/usr/bin/env python3
"""
PDF to DOCX Converter

Convert PDF to DOCX with accurate layout preservation.

How it works:
1. (Optional) Claude Vision provides layout hints - NOT required
2. Python (PyMuPDF) extracts all text, images, tables
3. Automatic column detection from text positions
4. DOCX is generated matching the original structure

Usage:
    python convert.py document.pdf
    python convert.py document.pdf -o output.docx

Works WITHOUT API key! API key only provides optional layout hints.
"""

import os
import sys
from pathlib import Path

def convert_pdf_to_docx(pdf_path: str, docx_path: str = None) -> bool:
    """
    Convert PDF to DOCX with maximum fidelity.

    Uses pdf2docx library which is specifically designed for
    accurate PDF to DOCX conversion.
    """
    from pdf2docx import Converter

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to DOCX with layout preservation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python convert.py document.pdf              Convert to document.docx
  python convert.py input.pdf -o output.docx  Specify output path

The converter works WITHOUT an API key.
Set ANTHROPIC_API_KEY only if you want optional layout hints.
        """
    )
    
    parser.add_argument('pdf_file', help='PDF file to convert')
    parser.add_argument('-o', '--output', help='Output DOCX path')
    parser.add_argument('-k', '--api-key', help='Anthropic API key (optional)')
    parser.add_argument('-v', '--version', action='version', version='4.1.0')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.pdf_file):
        print(f"Error: File not found: {args.pdf_file}")
        sys.exit(1)
    
    # Import here for faster startup on errors
    from pdf_converter import convert
    
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
    
    if not api_key:
        print("\nNote: No ANTHROPIC_API_KEY set.")
        print("      Conversion will work fine using automatic layout detection.")
        print("      API key only provides optional hints for complex layouts.\n")
    
    result = convert(args.pdf_file, args.output, api_key)
    
    if result.success:
        print(f"\nOutput: {result.output_path}")
        sys.exit(0)
    else:
        print(f"\nConversion failed:")
        for err in result.errors:
            print(f"  - {err}")
        sys.exit(1)

    # Check for output flag
    docx_path = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            docx_path = sys.argv[idx + 1]

    success = convert_pdf_to_docx(pdf_path, docx_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
