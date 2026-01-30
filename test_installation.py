#!/usr/bin/env python3
"""Test installation of PDF to DOCX converter."""

import sys
import os


def main():
    print("Checking dependencies...")

    deps = {'fitz': 'PyMuPDF', 'docx': 'python-docx'}
    missing = []

    for module, name in deps.items():
        try:
            __import__(module)
            print(f"  [OK] {name}")
        except ImportError:
            print(f"  [MISSING] {name}")
            missing.append(name)

    # Optional
    try:
        __import__('anthropic')
        print(f"  [OK] anthropic (optional)")
    except ImportError:
        print(f"  [--] anthropic (optional, for layout analysis)")

    if missing:
        print(f"\nInstall missing: pip install {' '.join(missing)}")
        sys.exit(1)

    print("\nChecking converter...")
    try:
        from pdf_converter import convert
        print("  [OK] Converter loaded")
    except ImportError as e:
        print(f"  [ERROR] {e}")
        sys.exit(1)

    print("\nReady! Usage:")
    print("  python convert.py document.pdf")

    if len(sys.argv) > 1:
        print(f"\nTesting with: {sys.argv[1]}")
        from pdf_converter import convert
        result = convert(sys.argv[1])
        print(f"Result: {'SUCCESS' if result.success else 'FAILED'}")


if __name__ == '__main__':
    main()
