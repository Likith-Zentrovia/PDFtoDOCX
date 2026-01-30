"""
PDF Layout Analyzer (Optional)

This module provides OPTIONAL layout hints using Claude Vision.
The main extraction works WITHOUT this - it's only used for:
- Confirming column detection in complex documents
- Getting hints about unusual layouts

All heavy lifting is done by pdf_extractor.py using pure Python.
"""

import os
import base64
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field
import fitz  # PyMuPDF


@dataclass
class LayoutHints:
    """Optional hints from Vision analysis."""
    num_columns: int = 1
    has_complex_layout: bool = False
    confidence: float = 0.0


@dataclass
class DocumentLayout:
    """Layout analysis for entire document (for backward compatibility)."""
    page_count: int
    pages: List[Any]  # Keeping for compatibility
    dominant_columns: int
    is_consistent: bool


class LayoutAnalyzer:
    """
    OPTIONAL layout analyzer using Claude Vision.
    
    Only used when API key is provided and even then
    just provides hints to supplement the Python extraction.
    """
    
    LAYOUT_PROMPT = """Look at this PDF page and answer in JSON format:
{
    "num_columns": <1, 2, or 3>,
    "has_complex_layout": <true if tables/figures/sidebars present, false otherwise>
}

Only JSON, no explanation."""
    
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        self.client = None
        
        if self.api_key:
            try:
                import anthropic
                self.client = anthropic.Anthropic(api_key=self.api_key)
            except Exception:
                self.client = None
    
    def analyze_document(self, pdf_path: str, sample_pages: int = 1) -> Optional[DocumentLayout]:
        """
        Quick layout analysis - just samples first page.
        
        Returns None if no API key or analysis fails.
        """
        if not self.client:
            return None
        
        try:
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            
            # Only analyze first page for speed
            hints = self._analyze_page(doc, 0)
            doc.close()
            
            if hints:
                return DocumentLayout(
                    page_count=total_pages,
                    pages=[],  # Not used
                    dominant_columns=hints.num_columns,
                    is_consistent=True
                )
            
        except Exception:
            pass
        
        return None
    
    def _analyze_page(self, doc: fitz.Document, page_num: int) -> Optional[LayoutHints]:
        """Analyze a single page for layout hints."""
        try:
            page = doc[page_num]
            
            # Low resolution for speed
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
            img_bytes = pix.tobytes("png")
            img_b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
            
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=100,  # Minimal tokens needed
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                        {"type": "text", "text": self.LAYOUT_PROMPT}
                    ]
                }]
            )
            
            text = response.content[0].text.strip()
            
            # Parse JSON
            if text.startswith("```"):
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]
            
            import json
            data = json.loads(text.strip())
            
            return LayoutHints(
                num_columns=data.get("num_columns", 1),
                has_complex_layout=data.get("has_complex_layout", False),
                confidence=0.8
            )
            
        except Exception:
            return None


def get_layout_hints(pdf_path: str, api_key: Optional[str] = None) -> Optional[LayoutHints]:
    """
    Convenience function to get quick layout hints.
    
    Returns None if no API key or analysis fails.
    Does NOT block the main conversion process.
    """
    if not api_key:
        return None
    
    try:
        analyzer = LayoutAnalyzer(api_key)
        result = analyzer.analyze_document(pdf_path, sample_pages=1)
        if result:
            return LayoutHints(
                num_columns=result.dominant_columns,
                has_complex_layout=False,
                confidence=0.8
            )
    except Exception:
        pass
    
    return None
