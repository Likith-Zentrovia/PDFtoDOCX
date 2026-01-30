"""
PDF to DOCX Converter Package

A high-fidelity PDF to DOCX converter that preserves:
- Text formatting (font, size, weight, style, color)
- Images with accurate positioning
- Tables with proper structure
- Multi-column layouts
- Page layout and structure
"""

from .converter import PDFConverter, convert_pdf_to_docx
from .analyzer import PDFAnalyzer, PageLayout, LayoutType
from .advanced_converter import AdvancedPDFConverter, convert_with_layout_preservation
from .postprocessor import PostProcessor, ConversionValidator

__version__ = "2.0.0"
__all__ = [
    "PDFConverter",
    "convert_pdf_to_docx",
    "PDFAnalyzer",
    "PageLayout",
    "LayoutType",
    "AdvancedPDFConverter",
    "convert_with_layout_preservation",
    "PostProcessor",
    "ConversionValidator"
]
