#!/usr/bin/env python3
"""
PDF to DOCX Converter

High-fidelity conversion using pdf2docx library with optimized settings
for multi-column layouts.

Usage:
    python convert.py document.pdf
    python convert.py document.pdf -o output.docx
"""

import os
import sys
from pathlib import Path


def convert_pdf_to_docx(pdf_path: str, docx_path: str = None) -> bool:
    """
    Convert PDF to DOCX with maximum fidelity.
    """
    from pdf2docx import Converter

    pdf_path = str(Path(pdf_path).resolve())

    if docx_path is None:
        docx_path = str(Path(pdf_path).with_suffix('.docx'))
    else:
        docx_path = str(Path(docx_path).resolve())

    print(f"\nConverting: {pdf_path}")
    print(f"Output:     {docx_path}")

    try:
        cv = Converter(pdf_path)
        page_count = len(cv.pages)
        print(f"Pages:      {page_count}")
        print("\nProcessing...")

        # Convert with optimized settings for better layout handling
        cv.convert(
            docx_path,
            # Process all pages
            start=0,
            end=None,
            # Keep tables and preserve structure
            kwargs={
                'connected_border': False,  # Don't connect table borders
                'line_overlap_threshold': 0.9,  # Higher = stricter line detection
                'line_break_width_ratio': 0.5,  # Column break detection
                'line_break_free_space_ratio': 0.1,  # Space between columns
                'line_separate_threshold': 5.0,  # Separate lines threshold
                'new_paragraph_threshold': 0.5,  # New paragraph detection
            }
        )
        cv.close()

        if os.path.exists(docx_path):
            size_kb = os.path.getsize(docx_path) / 1024
            print(f"\nSuccess!")
            print(f"Output: {docx_path}")
            print(f"Size:   {size_kb:.1f} KB")
            return True
        else:
            print("\nError: Output file not created")
            return False

    except TypeError:
        # If kwargs not supported, try without
        print("Retrying with default settings...")
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()

            if os.path.exists(docx_path):
                size_kb = os.path.getsize(docx_path) / 1024
                print(f"\nSuccess!")
                print(f"Output: {docx_path}")
                print(f"Size:   {size_kb:.1f} KB")
                return True
            return False
        except Exception as e:
            print(f"\nError: {e}")
            return False

    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    if len(sys.argv) < 2:
        print("Usage: python convert.py <pdf_file> [-o output.docx]")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    docx_path = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            docx_path = sys.argv[idx + 1]

    success = convert_pdf_to_docx(pdf_path, docx_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
