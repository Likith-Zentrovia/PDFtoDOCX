"""
PDF to DOCX Converter Package - v3.0

Intelligent PDF to DOCX conversion powered by Claude Vision AI.

Features:
- Visual analysis of each page using Claude Vision
- Multi-column layout detection and preservation
- Accurate text formatting (font, size, style, color)
- Image extraction and positioning
- Table detection and recreation
- Automatic validation

Usage:
    from pdf_converter import convert_pdf

    result = convert_pdf("document.pdf")
    print(f"Saved to: {result.output_path}")
"""

from .intelligent_converter import IntelligentConverter, convert_pdf, ConversionResult
from .vision_analyzer import ClaudeVisionAnalyzer, DocumentAnalysis, PageAnalysis

__version__ = "3.0.0"
__all__ = [
    "convert_pdf",
    "IntelligentConverter",
    "ConversionResult",
    "ClaudeVisionAnalyzer",
    "DocumentAnalysis",
    "PageAnalysis"
]
