#!/usr/bin/env python3
"""
Test script to verify PDF to DOCX converter installation.

This script checks that all required dependencies are installed
and provides a simple test of the conversion functionality.

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
        'pdf2docx': 'pdf2docx',
        'fitz': 'PyMuPDF',
        'docx': 'python-docx',
        'click': 'click',
        'PIL': 'Pillow',
    }

    all_ok = True

    for module, package in dependencies.items():
        try:
            __import__(module)
            print(f"  [OK] {package}")
        except ImportError:
            print(f"  [MISSING] {package} - install with: pip install {package}")
            all_ok = False

    # Optional dependency
    try:
        __import__('tqdm')
        print(f"  [OK] tqdm (optional)")
    except ImportError:
        print(f"  [INFO] tqdm (optional) - not installed, progress bars disabled")

    print("-" * 40)

    if all_ok:
        print("All required dependencies are installed!")
    else:
        print("\nSome dependencies are missing. Install them with:")
        print("  pip install -r requirements.txt")

    return all_ok


def check_converter_module():
    """Check if the converter module can be imported."""
    print("\nChecking converter module...")
    print("-" * 40)

    try:
        from pdf_converter import PDFConverter
        print("  [OK] PDFConverter imported successfully")

        from pdf_converter.converter import convert_pdf_to_docx
        print("  [OK] convert_pdf_to_docx imported successfully")

        # Create instance
        converter = PDFConverter(verbose=False)
        print("  [OK] PDFConverter instance created")

        return True
    except ImportError as e:
        print(f"  [ERROR] Failed to import: {e}")
        return False
    except Exception as e:
        print(f"  [ERROR] Unexpected error: {e}")
        return False


def test_conversion(pdf_path: str):
    """Test conversion with a real PDF file."""
    print(f"\nTesting conversion with: {pdf_path}")
    print("-" * 40)

    if not os.path.exists(pdf_path):
        print(f"  [ERROR] File not found: {pdf_path}")
        return False

    try:
        from pdf_converter import PDFConverter

        converter = PDFConverter(verbose=True)

        # Get PDF info
        print("\n  Getting PDF info...")
        info = converter.get_pdf_info(pdf_path)
        print(f"    Pages: {info['page_count']}")
        total_images = sum(p['image_count'] for p in info['pages'])
        print(f"    Images: {total_images}")

        # Convert
        print("\n  Converting...")
        result = converter.convert(pdf_path)

        if result.success:
            output_size = os.path.getsize(result.output_path)
            print(f"\n  [SUCCESS] Conversion complete!")
            print(f"    Output: {result.output_path}")
            print(f"    Pages converted: {result.pages_converted}")
            print(f"    Output size: {output_size / 1024:.1f} KB")
            return True
        else:
            print(f"\n  [ERROR] Conversion failed: {result.error_message}")
            return False

    except Exception as e:
        print(f"  [ERROR] {e}")
        return False


def main():
    """Main test function."""
    print("=" * 50)
    print("  PDF to DOCX Converter - Installation Test")
    print("=" * 50)
    print()

    # Check dependencies
    deps_ok = check_dependencies()

    if not deps_ok:
        sys.exit(1)

    # Check converter module
    module_ok = check_converter_module()

    if not module_ok:
        print("\nConverter module check failed.")
        sys.exit(1)

    # Test with actual PDF if provided
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        test_conversion(pdf_path)
    else:
        print("\nTo test with an actual PDF file, run:")
        print("  python test_installation.py /path/to/your/file.pdf")

    print("\n" + "=" * 50)
    print("  Installation test complete!")
    print("=" * 50)
    print("\nYou can now use the converter:")
    print("  python convert.py your_document.pdf")


if __name__ == '__main__':
    main()
