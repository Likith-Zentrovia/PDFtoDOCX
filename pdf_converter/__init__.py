"""
PDF to DOCX Converter Package

A high-fidelity PDF to DOCX converter that preserves:
- Text formatting (font, size, weight, style, color)
- Images with accurate positioning
- Tables with proper structure
- Page layout and structure
"""

from .converter import PDFConverter

__version__ = "1.0.0"
__all__ = ["PDFConverter"]
