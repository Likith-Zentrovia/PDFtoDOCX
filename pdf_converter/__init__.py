"""
PDF to DOCX Converter Package

Converts PDF to DOCX with accurate layout preservation:
- Claude Vision analyzes layout structure (columns, headers, tables, images)
- Python extracts all text, images, and tables from PDF
- DOCX is generated matching the original structure with proper element ordering

Key Features:
- Accurate element positioning (text, images, tables interleaved by Y-position)
- Table detection and extraction
- Multi-column layout support
- Proper spacing preservation

Usage:
    from pdf_converter import convert
    result = convert("document.pdf")
"""

from .converter import convert, PDFtoDOCXConverter, ConversionResult

__version__ = "4.0.0"
__all__ = ["convert", "PDFtoDOCXConverter", "ConversionResult"]
