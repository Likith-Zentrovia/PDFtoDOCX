#!/usr/bin/env python3
"""
Test script to verify PDF to DOCX converter installation (v3.0).

Usage:
    python test_installation.py
    python test_installation.py /path/to/test.pdf
"""

import sys
import os


def check_dependencies():
    """Check if all required dependencies are installed."""
    print("Checking dependencies...")
    print("-" * 40)

    dependencies = {
        'anthropic': 'anthropic',
        'fitz': 'PyMuPDF',
        'docx': 'python-docx',
        'PIL': 'Pillow',
    }

    all_ok = True

    for module, package in dependencies.items():
        try:
            __import__(module)
            print(f"  [OK] {package}")
        except ImportError:
            print(f"  [MISSING] {package}")
            all_ok = False

    print("-" * 40)

    if all_ok:
        print("All dependencies installed!")
    else:
        print("\nInstall missing dependencies:")
        print("  pip install -r requirements.txt")

    return all_ok


def check_api_key():
    """Check if API key is configured."""
    print("\nChecking API key...")
    print("-" * 40)

    if os.environ.get("ANTHROPIC_API_KEY"):
        key = os.environ.get("ANTHROPIC_API_KEY")
        masked = key[:8] + "..." + key[-4:] if len(key) > 12 else "***"
        print(f"  [OK] ANTHROPIC_API_KEY is set ({masked})")
        return True
    else:
        print("  [MISSING] ANTHROPIC_API_KEY not set")
        print("\n  To set it:")
        print("    export ANTHROPIC_API_KEY='your-key-here'")
        return False


def check_converter():
    """Check if converter modules can be imported."""
    print("\nChecking converter modules...")
    print("-" * 40)

    try:
        from pdf_converter import convert_pdf, IntelligentConverter
        print("  [OK] Converter modules loaded")
        return True
    except ImportError as e:
        print(f"  [ERROR] {e}")
        return False


def test_conversion(pdf_path: str):
    """Test conversion with a real PDF."""
    print(f"\nTesting conversion: {pdf_path}")
    print("-" * 40)

    if not os.path.exists(pdf_path):
        print(f"  [ERROR] File not found: {pdf_path}")
        return False

    try:
        from pdf_converter import convert_pdf

        result = convert_pdf(pdf_path)

        if result.success:
            print(f"\n  [SUCCESS] Conversion complete!")
            print(f"    Output: {result.output_path}")
            print(f"    Pages: {result.pages_converted}")
            print(f"    Text elements: {result.text_elements_processed}")
            print(f"    Quality: {result.validation_score:.0%}")
            return True
        else:
            print(f"  [ERROR] Conversion failed")
            for note in result.notes:
                print(f"    - {note}")
            return False

    except Exception as e:
        print(f"  [ERROR] {e}")
        return False


def main():
    print("=" * 50)
    print("  PDF to DOCX Converter v3.0 - Installation Test")
    print("=" * 50)

    # Check dependencies
    deps_ok = check_dependencies()
    if not deps_ok:
        sys.exit(1)

    # Check API key
    api_ok = check_api_key()

    # Check converter
    converter_ok = check_converter()
    if not converter_ok:
        sys.exit(1)

    # Test with PDF if provided
    if len(sys.argv) > 1 and api_ok:
        test_conversion(sys.argv[1])
    elif len(sys.argv) > 1:
        print("\nSkipping conversion test (API key not set)")
    else:
        print("\nTo test conversion:")
        print("  python test_installation.py /path/to/test.pdf")

    print("\n" + "=" * 50)
    print("  Installation check complete!")
    print("=" * 50)

    if api_ok:
        print("\nReady to convert:")
        print("  python convert.py document.pdf")
    else:
        print("\nSet API key, then convert:")
        print("  export ANTHROPIC_API_KEY='your-key'")
        print("  python convert.py document.pdf")


if __name__ == '__main__':
    main()
