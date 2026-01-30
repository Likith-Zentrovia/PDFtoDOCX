"""
PDF to DOCX Converter Package

Converts PDF to DOCX with accurate layout preservation:
- Claude Vision analyzes layout structure (columns, headers, etc.)
- Python extracts all text and images from PDF
- DOCX is generated matching the original structure

Usage:
    from pdf_converter import convert
    result = convert("document.pdf")
"""

from .converter import convert, PDFtoDOCXConverter, ConversionResult

__version__ = "3.0.0"
__all__ = ["convert", "PDFtoDOCXConverter", "ConversionResult"]
