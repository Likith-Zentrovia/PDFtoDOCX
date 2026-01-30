"""
Claude Vision-Based PDF Page Analyzer

Uses Claude's vision capabilities to analyze PDF pages and understand:
- Page layout (single/multi-column)
- Text block positions and reading order
- Font styles and formatting
- Image and table locations
- Headers, footers, and page numbers
"""

import os
import io
import json
import base64
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
import fitz  # PyMuPDF
from PIL import Image
import anthropic


@dataclass
class TextElement:
    """A text element identified by Claude Vision."""
    text: str
    bbox: Dict[str, float]  # x, y, width, height as percentages
    font_size: str  # "small", "medium", "large", "xlarge"
    font_weight: str  # "normal", "bold"
    font_style: str  # "normal", "italic"
    color: str  # hex color or description
    is_header: bool = False
    is_footer: bool = False
    is_page_number: bool = False
    column: int = 1
    reading_order: int = 0


@dataclass
class ImageElement:
    """An image element identified by Claude Vision."""
    bbox: Dict[str, float]
    description: str
    is_background: bool = False


@dataclass
class TableElement:
    """A table identified by Claude Vision."""
    bbox: Dict[str, float]
    rows: int
    cols: int
    has_header: bool
    content_summary: str


@dataclass
class PageAnalysis:
    """Complete analysis of a single page."""
    page_num: int
    width: float
    height: float
    layout_type: str  # "single_column", "two_column", "three_column", "complex"
    num_columns: int
    column_positions: List[Dict[str, float]]  # Column boundaries
    text_elements: List[TextElement]
    images: List[ImageElement]
    tables: List[TableElement]
    reading_order: List[int]  # Indices of text_elements in reading order
    has_header: bool
    has_footer: bool
    background_color: str
    special_notes: str = ""


@dataclass
class DocumentAnalysis:
    """Complete document analysis."""
    page_count: int
    pages: List[PageAnalysis]
    dominant_layout: str
    consistent_style: bool
    document_type: str  # "report", "article", "form", etc.


class ClaudeVisionAnalyzer:
    """
    Analyzes PDF pages using Claude Vision API.

    This analyzer:
    1. Converts PDF pages to high-quality images
    2. Sends each page to Claude Vision for analysis
    3. Extracts detailed layout and content information
    4. Returns structured data for accurate DOCX conversion
    """

    ANALYSIS_PROMPT = """Analyze this PDF page image and provide a detailed JSON response with the following structure:

{
    "layout_type": "single_column" | "two_column" | "three_column" | "complex",
    "num_columns": <number>,
    "column_positions": [
        {"x_start": <0-100>, "x_end": <0-100>}
    ],
    "background_color": "<hex or description>",
    "has_header": <boolean>,
    "has_footer": <boolean>,
    "text_elements": [
        {
            "text": "<actual text content>",
            "bbox": {"x": <0-100>, "y": <0-100>, "width": <0-100>, "height": <0-100>},
            "font_size": "small" | "medium" | "large" | "xlarge",
            "font_weight": "normal" | "bold",
            "font_style": "normal" | "italic",
            "color": "<hex or description like 'white', 'black'>",
            "is_header": <boolean>,
            "is_footer": <boolean>,
            "is_page_number": <boolean>,
            "column": <1-based column number>,
            "reading_order": <sequence number for reading order>
        }
    ],
    "images": [
        {
            "bbox": {"x": <0-100>, "y": <0-100>, "width": <0-100>, "height": <0-100>},
            "description": "<what the image shows>",
            "is_background": <boolean>
        }
    ],
    "tables": [
        {
            "bbox": {"x": <0-100>, "y": <0-100>, "width": <0-100>, "height": <0-100>},
            "rows": <number>,
            "cols": <number>,
            "has_header": <boolean>,
            "content_summary": "<brief description of table content>"
        }
    ],
    "special_notes": "<any important observations about layout, styling, or conversion challenges>"
}

IMPORTANT INSTRUCTIONS:
1. For multi-column layouts, identify EACH column and assign text to the correct column
2. Reading order should follow: left column top-to-bottom, then right column top-to-bottom
3. All bbox coordinates are PERCENTAGES (0-100) relative to page dimensions
4. Extract ALL visible text, preserving the exact wording
5. Note any special formatting like colored backgrounds, sidebars, callout boxes
6. For headers/footers, check if they appear at consistent positions
7. Identify page numbers and mark them appropriately

Respond ONLY with the JSON object, no additional text."""

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize the analyzer.

        Args:
            api_key: Anthropic API key. If not provided, uses ANTHROPIC_API_KEY env var.
        """
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError(
                "Anthropic API key required. Set ANTHROPIC_API_KEY environment variable "
                "or pass api_key parameter."
            )
        self.client = anthropic.Anthropic(api_key=self.api_key)

    def analyze_document(
        self,
        pdf_path: str,
        pages: Optional[List[int]] = None,
        dpi: int = 150,
        progress_callback: Optional[callable] = None
    ) -> DocumentAnalysis:
        """
        Analyze entire PDF document.

        Args:
            pdf_path: Path to PDF file
            pages: Specific pages to analyze (0-indexed). None = all pages.
            dpi: Resolution for page rendering
            progress_callback: Optional callback(current, total) for progress updates

        Returns:
            DocumentAnalysis with complete document structure
        """
        doc = fitz.open(pdf_path)

        if pages is None:
            pages = list(range(len(doc)))

        page_analyses = []
        layout_types = []

        for i, page_num in enumerate(pages):
            if progress_callback:
                progress_callback(i + 1, len(pages))

            print(f"  Analyzing page {page_num + 1}/{len(doc)} with Claude Vision...")

            # Render page to image
            page = doc[page_num]
            page_image = self._render_page(page, dpi)

            # Analyze with Claude Vision
            analysis = self._analyze_page(page_image, page_num, page.rect.width, page.rect.height)
            page_analyses.append(analysis)
            layout_types.append(analysis.layout_type)

        doc.close()

        # Determine dominant layout
        layout_counts = {}
        for lt in layout_types:
            layout_counts[lt] = layout_counts.get(lt, 0) + 1
        dominant_layout = max(layout_counts, key=layout_counts.get)

        # Check style consistency
        consistent_style = len(set(layout_types)) <= 2

        # Determine document type
        doc_type = self._infer_document_type(page_analyses)

        return DocumentAnalysis(
            page_count=len(pages),
            pages=page_analyses,
            dominant_layout=dominant_layout,
            consistent_style=consistent_style,
            document_type=doc_type
        )

    def analyze_single_page(
        self,
        pdf_path: str,
        page_num: int,
        dpi: int = 150
    ) -> PageAnalysis:
        """Analyze a single page."""
        doc = fitz.open(pdf_path)
        page = doc[page_num]

        page_image = self._render_page(page, dpi)
        analysis = self._analyze_page(page_image, page_num, page.rect.width, page.rect.height)

        doc.close()
        return analysis

    def _render_page(self, page: fitz.Page, dpi: int) -> bytes:
        """Render a PDF page to PNG image bytes."""
        # Calculate zoom factor for desired DPI
        zoom = dpi / 72
        matrix = fitz.Matrix(zoom, zoom)

        # Render page
        pix = page.get_pixmap(matrix=matrix)

        # Convert to PNG bytes
        img_bytes = pix.tobytes("png")

        return img_bytes

    def _analyze_page(
        self,
        image_bytes: bytes,
        page_num: int,
        width: float,
        height: float
    ) -> PageAnalysis:
        """Analyze a single page image using Claude Vision."""
        # Encode image to base64
        image_b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

        # Call Claude Vision API
        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=8000,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": image_b64
                                }
                            },
                            {
                                "type": "text",
                                "text": self.ANALYSIS_PROMPT
                            }
                        ]
                    }
                ]
            )

            # Parse JSON response
            response_text = response.content[0].text

            # Clean up response if needed (remove markdown code blocks)
            if response_text.startswith("```"):
                response_text = response_text.split("```")[1]
                if response_text.startswith("json"):
                    response_text = response_text[4:]
            response_text = response_text.strip()

            data = json.loads(response_text)

        except json.JSONDecodeError as e:
            print(f"    Warning: Failed to parse Claude response, using fallback analysis")
            data = self._fallback_analysis()
        except Exception as e:
            print(f"    Warning: Claude API error: {e}, using fallback analysis")
            data = self._fallback_analysis()

        # Convert to PageAnalysis
        return self._parse_analysis(data, page_num, width, height)

    def _parse_analysis(
        self,
        data: Dict[str, Any],
        page_num: int,
        width: float,
        height: float
    ) -> PageAnalysis:
        """Parse Claude's JSON response into PageAnalysis."""
        # Parse text elements
        text_elements = []
        for te in data.get("text_elements", []):
            text_elements.append(TextElement(
                text=te.get("text", ""),
                bbox=te.get("bbox", {}),
                font_size=te.get("font_size", "medium"),
                font_weight=te.get("font_weight", "normal"),
                font_style=te.get("font_style", "normal"),
                color=te.get("color", "black"),
                is_header=te.get("is_header", False),
                is_footer=te.get("is_footer", False),
                is_page_number=te.get("is_page_number", False),
                column=te.get("column", 1),
                reading_order=te.get("reading_order", 0)
            ))

        # Parse images
        images = []
        for img in data.get("images", []):
            images.append(ImageElement(
                bbox=img.get("bbox", {}),
                description=img.get("description", ""),
                is_background=img.get("is_background", False)
            ))

        # Parse tables
        tables = []
        for tbl in data.get("tables", []):
            tables.append(TableElement(
                bbox=tbl.get("bbox", {}),
                rows=tbl.get("rows", 0),
                cols=tbl.get("cols", 0),
                has_header=tbl.get("has_header", False),
                content_summary=tbl.get("content_summary", "")
            ))

        # Build reading order
        reading_order = sorted(
            range(len(text_elements)),
            key=lambda i: text_elements[i].reading_order
        )

        return PageAnalysis(
            page_num=page_num,
            width=width,
            height=height,
            layout_type=data.get("layout_type", "single_column"),
            num_columns=data.get("num_columns", 1),
            column_positions=data.get("column_positions", []),
            text_elements=text_elements,
            images=images,
            tables=tables,
            reading_order=reading_order,
            has_header=data.get("has_header", False),
            has_footer=data.get("has_footer", False),
            background_color=data.get("background_color", "white"),
            special_notes=data.get("special_notes", "")
        )

    def _fallback_analysis(self) -> Dict[str, Any]:
        """Return a basic fallback analysis structure."""
        return {
            "layout_type": "single_column",
            "num_columns": 1,
            "column_positions": [{"x_start": 0, "x_end": 100}],
            "background_color": "white",
            "has_header": False,
            "has_footer": False,
            "text_elements": [],
            "images": [],
            "tables": [],
            "special_notes": "Fallback analysis used"
        }

    def _infer_document_type(self, pages: List[PageAnalysis]) -> str:
        """Infer the type of document from analysis."""
        # Check for common patterns
        has_toc = any("content" in p.special_notes.lower() or "table of" in p.special_notes.lower() for p in pages)
        multi_column = sum(1 for p in pages if p.num_columns > 1) > len(pages) / 2
        has_many_tables = sum(len(p.tables) for p in pages) > len(pages)

        if multi_column:
            return "article"
        elif has_toc:
            return "report"
        elif has_many_tables:
            return "form"
        else:
            return "document"
