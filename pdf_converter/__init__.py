"""
PDF to DOCX Converter Package

Converts PDF to DOCX with accurate layout preservation.

Key Features:
- Works WITHOUT AI - pure Python extraction and conversion
- Automatic column detection from text positions  
- Tables extracted and rendered properly
- Images placed at correct positions
- Font formatting preserved

Optional: Provide ANTHROPIC_API_KEY for layout hints (1 API call max)

Usage:
    from pdf_converter import convert
    result = convert("document.pdf")
"""

from .converter import convert, PDFtoDOCXConverter, ConversionResult

__version__ = "4.1.0"
__all__ = ["convert", "PDFtoDOCXConverter", "ConversionResult"]
