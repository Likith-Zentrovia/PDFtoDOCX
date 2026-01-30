#!/usr/bin/env python3
"""
Test script to verify PDF to DOCX converter installation (v2.0).

This script checks that all required dependencies are installed
and provides a test of the advanced conversion functionality.

Usage:
    python test_installation.py
    python test_installation.py /path/to/test.pdf
    python test_installation.py /path/to/test.pdf --advanced
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


def check_converter_modules():
    """Check if all converter modules can be imported."""
    print("\nChecking converter modules...")
    print("-" * 40)

    modules = [
        ("pdf_converter", "Main package"),
        ("pdf_converter.converter", "Standard converter"),
        ("pdf_converter.analyzer", "PDF analyzer"),
        ("pdf_converter.advanced_converter", "Advanced converter"),
        ("pdf_converter.postprocessor", "Post-processor"),
    ]

    all_ok = True

    for module_name, description in modules:
        try:
            __import__(module_name)
            print(f"  [OK] {description}")
        except ImportError as e:
            print(f"  [ERROR] {description}: {e}")
            all_ok = False

    # Test specific imports
    try:
        from pdf_converter import PDFConverter, AdvancedPDFConverter, PDFAnalyzer
        from pdf_converter import PostProcessor, ConversionValidator
        print(f"  [OK] All classes imported successfully")
    except ImportError as e:
        print(f"  [ERROR] Failed to import classes: {e}")
        all_ok = False

    return all_ok


def test_standard_conversion(pdf_path: str):
    """Test standard conversion."""
    print(f"\nTesting STANDARD conversion...")
    print("-" * 40)

    try:
        from pdf_converter import PDFConverter

        converter = PDFConverter(verbose=True)

        # Get PDF info
        print("  Getting PDF info...")
        info = converter.get_pdf_info(pdf_path)
        print(f"    Pages: {info['page_count']}")

        # Convert
        print("  Converting...")
        result = converter.convert(pdf_path)

        if result.success:
            output_size = os.path.getsize(result.output_path)
            print(f"\n  [SUCCESS] Standard conversion complete!")
            print(f"    Output: {result.output_path}")
            print(f"    Pages: {result.pages_converted}")
            print(f"    Size: {output_size / 1024:.1f} KB")
            return result.output_path
        else:
            print(f"\n  [ERROR] Conversion failed: {result.error_message}")
            return None

    except Exception as e:
        print(f"  [ERROR] {e}")
        return None


def test_advanced_conversion(pdf_path: str):
    """Test advanced conversion with layout preservation."""
    print(f"\nTesting ADVANCED conversion...")
    print("-" * 40)

    try:
        from pdf_converter import AdvancedPDFConverter

        converter = AdvancedPDFConverter()

        # Convert with validation
        output_path, stats, validation = converter.convert(
            pdf_path,
            validate=True
        )

        output_size = os.path.getsize(output_path)
        print(f"\n  [SUCCESS] Advanced conversion complete!")
        print(f"    Output: {output_path}")
        print(f"    Text blocks: {stats.text_blocks_converted}")
        print(f"    Columns preserved: {stats.columns_preserved}")
        print(f"    Images: {stats.images_extracted}")
        print(f"    Size: {output_size / 1024:.1f} KB")

        if validation:
            status = "PASSED" if validation.is_valid else "NEEDS REVIEW"
            print(f"    Validation: {status} ({validation.text_match_ratio:.1%} match)")

        return output_path

    except Exception as e:
        print(f"  [ERROR] {e}")
        import traceback
        traceback.print_exc()
        return None


def test_analyzer(pdf_path: str):
    """Test PDF structure analyzer."""
    print(f"\nTesting PDF ANALYZER...")
    print("-" * 40)

    try:
        from pdf_converter import PDFAnalyzer

        with PDFAnalyzer(pdf_path) as analyzer:
            structure = analyzer.analyze()

            print(f"  Pages: {structure.page_count}")
            print(f"  Dominant layout: {structure.dominant_layout.value}")
            print(f"  Consistent layout: {structure.has_consistent_layout}")

            # Per-page summary
            for page in structure.pages[:3]:  # First 3 pages
                print(f"    Page {page.page_num + 1}: {page.layout_type.value}, "
                      f"{len(page.columns)} columns, {len(page.text_blocks)} blocks")

            if len(structure.pages) > 3:
                print(f"    ... and {len(structure.pages) - 3} more pages")

        print(f"\n  [SUCCESS] Analyzer working correctly!")
        return True

    except Exception as e:
        print(f"  [ERROR] {e}")
        return False


def main():
    """Main test function."""
    print("=" * 60)
    print("  PDF to DOCX Converter v2.0 - Installation Test")
    print("=" * 60)
    print()

    # Check dependencies
    deps_ok = check_dependencies()
    if not deps_ok:
        sys.exit(1)

    # Check modules
    modules_ok = check_converter_modules()
    if not modules_ok:
        print("\nModule check failed.")
        sys.exit(1)

    # Test with PDF if provided
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]

        if not os.path.exists(pdf_path):
            print(f"\nError: File not found: {pdf_path}")
            sys.exit(1)

        use_advanced = "--advanced" in sys.argv

        # Test analyzer
        test_analyzer(pdf_path)

        # Test conversion
        if use_advanced:
            test_advanced_conversion(pdf_path)
        else:
            test_standard_conversion(pdf_path)

    else:
        print("\nTo test with a PDF file, run:")
        print("  python test_installation.py /path/to/your/file.pdf")
        print("  python test_installation.py /path/to/your/file.pdf --advanced")

    print("\n" + "=" * 60)
    print("  Installation test complete!")
    print("=" * 60)
    print("\nUsage:")
    print("  python convert.py document.pdf              # Standard conversion")
    print("  python convert.py document.pdf --advanced   # Advanced (multi-column)")
    print("  python convert.py document.pdf --analyze    # Analyze structure")


if __name__ == '__main__':
    main()
