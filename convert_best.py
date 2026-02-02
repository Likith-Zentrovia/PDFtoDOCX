#!/usr/bin/env python3
"""
Best-effort PDF to DOCX conversion.

Runs two conversion paths and saves both outputs so you can compare:
1. pdf2docx (convert.py) – best for most PDFs; uses fixed multi-column kwargs.
2. pdf_converter (PyMuPDF + python-docx) – alternative for complex layouts.

Usage:
    python convert_best.py document.pdf
    python convert_best.py document.pdf -o output.docx   # pdf2docx only to -o path
    python convert_best.py document.pdf --both         # write both versions
"""

import os
import sys
from pathlib import Path


def convert_pdf2docx(pdf_path: str, docx_path: str) -> bool:
    """Convert using pdf2docx with multi-column-optimized settings."""
    from pdf2docx import Converter

    cv = Converter(pdf_path)
    try:
        cv.convert(
            docx_path,
            start=0,
            end=None,
            kwargs={
                'connected_border_tolerance': 0.5,
                'line_break_width_ratio': 0.45,
                'line_break_free_space_ratio': 0.15,
                'line_separate_threshold': 8.0,
                'new_paragraph_free_space_ratio': 0.8,
                'line_overlap_threshold': 0.85,
                'delete_end_line_hyphen': True,
            }
        )
        return os.path.exists(docx_path)
    except Exception as e:
        print(f"  pdf2docx error: {e}")
        return False
    finally:
        cv.close()


def convert_pdf_converter(pdf_path: str, docx_path: str) -> bool:
    """Convert using local pdf_converter (PyMuPDF + python-docx)."""
    try:
        from pdf_converter import convert
        result = convert(pdf_path, output_path=docx_path)
        return result.success
    except Exception as e:
        print(f"  pdf_converter error: {e}")
        return False


def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_best.py <pdf_file> [-o output.docx] [--both]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    base = Path(pdf_path).resolve()
    out_arg = None
    if "-o" in sys.argv:
        idx = sys.argv.index("-o")
        if idx + 1 < len(sys.argv):
            out_arg = sys.argv[idx + 1]
    do_both = "--both" in sys.argv

    # Default: single output with pdf2docx
    if out_arg:
        docx_primary = str(Path(out_arg).resolve())
    else:
        docx_primary = str(base.with_suffix(".docx"))

    print(f"Input:  {pdf_path}")
    print(f"Output: {docx_primary}")
    if do_both:
        docx_alt = str(base.with_name(base.stem + "_pdf_converter.docx"))
        print(f"Also:   {docx_alt} (pdf_converter)")
    print()

    # 1) pdf2docx
    print("[1] Converting with pdf2docx (multi-column optimized)...")
    ok1 = convert_pdf2docx(pdf_path, docx_primary)
    if ok1:
        print(f"      Saved: {docx_primary}")
    else:
        print("      Failed.")

    if do_both:
        print("[2] Converting with pdf_converter...")
        ok2 = convert_pdf_converter(pdf_path, docx_alt)
        if ok2:
            print(f"      Saved: {docx_alt}")
        else:
            print("      Failed.")
        print()
        print("Compare both DOCX files and keep the one that looks better.")
    else:
        if not ok1:
            print("\nTip: Try with --both to also generate pdf_converter version.")
            sys.exit(1)

    sys.exit(0)


if __name__ == "__main__":
    main()
