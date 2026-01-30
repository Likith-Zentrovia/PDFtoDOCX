"""
PDF Layout Analyzer using Claude Vision

Uses Claude Vision to analyze page layout complexity:
- Number of columns
- Column boundaries (percentages)
- Header/footer regions
- Tables with boundaries
- Image regions
- Element positioning for accurate reconstruction

Does NOT extract any text content - that's done by Python.
"""

import os
import base64
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field
import fitz  # PyMuPDF
import anthropic


@dataclass
class TableRegion:
    """Information about a detected table region."""
    bbox_pct: Dict[str, float]  # {x_start, y_start, x_end, y_end} as percentages
    approx_rows: int
    approx_cols: int
    has_header_row: bool


@dataclass
class ImageRegion:
    """Information about a detected image region."""
    bbox_pct: Dict[str, float]  # {x_start, y_start, x_end, y_end} as percentages
    position_type: str  # "inline", "float_left", "float_right", "full_width", "centered"
    near_text_above: bool  # Is there text immediately above?
    near_text_below: bool  # Is there text immediately below?


@dataclass
class LayoutInfo:
    """Layout information for a single page."""
    page_num: int
    num_columns: int
    column_boundaries: List[Dict[str, float]]  # [{x_start: %, x_end: %}]
    has_header: bool
    header_height_pct: float  # Percentage of page height
    has_footer: bool
    footer_height_pct: float
    has_sidebar: bool
    sidebar_position: str  # "left", "right", "none"
    complexity: str  # "simple", "moderate", "complex"
    tables: List[TableRegion] = field(default_factory=list)
    images: List[ImageRegion] = field(default_factory=list)
    element_order: List[Dict[str, Any]] = field(default_factory=list)  # Ordered list of element regions


@dataclass
class DocumentLayout:
    """Layout analysis for entire document."""
    page_count: int
    pages: List[LayoutInfo]
    dominant_columns: int
    is_consistent: bool


class LayoutAnalyzer:
    """
    Analyzes PDF layout using Claude Vision.

    ONLY determines structure/layout - does NOT extract content.
    This is a quick analysis to guide the Python extraction.
    """

    LAYOUT_PROMPT = """Analyze this PDF page layout structure. Do NOT extract or read any text content.

Return a JSON object with this structure:
{
    "num_columns": <1, 2, or 3>,
    "column_boundaries": [
        {"x_start": <0-100>, "x_end": <0-100>}
    ],
    "has_header": <true/false>,
    "header_height_pct": <0-20>,
    "has_footer": <true/false>,
    "footer_height_pct": <0-15>,
    "has_sidebar": <true/false>,
    "sidebar_position": "left" | "right" | "none",
    "complexity": "simple" | "moderate" | "complex",
    "tables": [
        {
            "bbox_pct": {"x_start": <0-100>, "y_start": <0-100>, "x_end": <0-100>, "y_end": <0-100>},
            "approx_rows": <number>,
            "approx_cols": <number>,
            "has_header_row": <true/false>
        }
    ],
    "images": [
        {
            "bbox_pct": {"x_start": <0-100>, "y_start": <0-100>, "x_end": <0-100>, "y_end": <0-100>},
            "position_type": "inline" | "float_left" | "float_right" | "full_width" | "centered",
            "near_text_above": <true/false>,
            "near_text_below": <true/false>
        }
    ],
    "element_order": [
        {"type": "text", "y_start_pct": <0-100>, "y_end_pct": <0-100>, "column": <1-3 or 0 for full width>},
        {"type": "image", "y_start_pct": <0-100>, "y_end_pct": <0-100>, "index": <0-based index in images array>},
        {"type": "table", "y_start_pct": <0-100>, "y_end_pct": <0-100>, "index": <0-based index in tables array>}
    ]
}

Rules:
- num_columns: Count main content columns (1, 2, or 3)
- column_boundaries: X positions as percentages (0=left edge, 100=right edge)
- tables: List ALL visible tables with their bounding boxes and structure
- images: List ALL visible images/figures/charts with their positions
- element_order: List elements from TOP to BOTTOM of page in visual reading order
  - Include text regions, images, and tables
  - This helps reconstruct the exact page layout
- bbox_pct values are percentages of page dimensions (0-100)
- position_type for images:
  - "full_width": spans entire content width
  - "centered": centered on page, may have text above/below
  - "float_left"/"float_right": text wraps around
  - "inline": small image within text flow
- complexity: "simple"=single column no tables, "moderate"=2 columns or few tables, "complex"=3+ columns or many elements

Respond with ONLY the JSON, no explanation."""

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY required")
        self.client = anthropic.Anthropic(api_key=self.api_key)

    def analyze_document(self, pdf_path: str, sample_pages: int = 3) -> DocumentLayout:
        """
        Analyze document layout by sampling a few pages.

        Args:
            pdf_path: Path to PDF
            sample_pages: Number of pages to sample (default 3 for speed)
        """
        doc = fitz.open(pdf_path)
        total_pages = len(doc)

        # Sample pages: first, middle, and one from later section
        if total_pages <= sample_pages:
            pages_to_analyze = list(range(total_pages))
        else:
            pages_to_analyze = [
                0,  # First page
                total_pages // 2,  # Middle
                min(total_pages - 1, total_pages // 2 + 2)  # A bit after middle
            ]

        print(f"  Analyzing layout of {len(pages_to_analyze)} sample pages...")

        page_layouts = []
        column_counts = []

        for page_num in pages_to_analyze:
            layout = self._analyze_page(doc, page_num)
            page_layouts.append(layout)
            column_counts.append(layout.num_columns)

        doc.close()

        # Determine dominant layout
        dominant_cols = max(set(column_counts), key=column_counts.count)
        is_consistent = len(set(column_counts)) == 1

        # Extend layout to all pages using dominant pattern
        all_layouts = []
        template = page_layouts[0] if page_layouts else self._default_layout(0)

        for i in range(total_pages):
            if i in pages_to_analyze:
                idx = pages_to_analyze.index(i)
                all_layouts.append(page_layouts[idx])
            else:
                # Use template with correct page number (tables/images/element_order empty for non-analyzed pages)
                all_layouts.append(LayoutInfo(
                    page_num=i,
                    num_columns=dominant_cols,
                    column_boundaries=template.column_boundaries,
                    has_header=template.has_header,
                    header_height_pct=template.header_height_pct,
                    has_footer=template.has_footer,
                    footer_height_pct=template.footer_height_pct,
                    has_sidebar=template.has_sidebar,
                    sidebar_position=template.sidebar_position,
                    complexity=template.complexity,
                    tables=[],  # Can't know tables for non-analyzed pages
                    images=[],  # Can't know images for non-analyzed pages
                    element_order=[]  # Will be determined from extraction
                ))

        return DocumentLayout(
            page_count=total_pages,
            pages=all_layouts,
            dominant_columns=dominant_cols,
            is_consistent=is_consistent
        )

    def _analyze_page(self, doc: fitz.Document, page_num: int) -> LayoutInfo:
        """Analyze a single page layout."""
        page = doc[page_num]

        # Render at higher resolution for better layout analysis
        zoom = 1.5  # ~108 DPI for better detail
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img_bytes = pix.tobytes("png")
        img_b64 = base64.standard_b64encode(img_bytes).decode("utf-8")

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=2000,  # More tokens for detailed layout info
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                        {"type": "text", "text": self.LAYOUT_PROMPT}
                    ]
                }]
            )

            text = response.content[0].text.strip()
            if text.startswith("```"):
                text = text.split("```")[1]
                if text.startswith("json"):
                    text = text[4:]

            import json
            data = json.loads(text.strip())

            # Parse table regions
            tables = []
            for t in data.get("tables", []):
                tables.append(TableRegion(
                    bbox_pct=t.get("bbox_pct", {"x_start": 0, "y_start": 0, "x_end": 100, "y_end": 100}),
                    approx_rows=t.get("approx_rows", 1),
                    approx_cols=t.get("approx_cols", 1),
                    has_header_row=t.get("has_header_row", False)
                ))

            # Parse image regions
            images = []
            for img in data.get("images", []):
                images.append(ImageRegion(
                    bbox_pct=img.get("bbox_pct", {"x_start": 0, "y_start": 0, "x_end": 100, "y_end": 100}),
                    position_type=img.get("position_type", "centered"),
                    near_text_above=img.get("near_text_above", True),
                    near_text_below=img.get("near_text_below", True)
                ))

            return LayoutInfo(
                page_num=page_num,
                num_columns=data.get("num_columns", 1),
                column_boundaries=data.get("column_boundaries", [{"x_start": 5, "x_end": 95}]),
                has_header=data.get("has_header", False),
                header_height_pct=data.get("header_height_pct", 0),
                has_footer=data.get("has_footer", False),
                footer_height_pct=data.get("footer_height_pct", 0),
                has_sidebar=data.get("has_sidebar", False),
                sidebar_position=data.get("sidebar_position", "none"),
                complexity=data.get("complexity", "simple"),
                tables=tables,
                images=images,
                element_order=data.get("element_order", [])
            )

        except Exception as e:
            print(f"    Warning: Vision analysis failed for page {page_num + 1}: {e}")
            return self._default_layout(page_num)

    def _default_layout(self, page_num: int) -> LayoutInfo:
        """Return default single-column layout."""
        return LayoutInfo(
            page_num=page_num,
            num_columns=1,
            column_boundaries=[{"x_start": 5, "x_end": 95}],
            has_header=False,
            header_height_pct=0,
            has_footer=False,
            footer_height_pct=0,
            has_sidebar=False,
            sidebar_position="none",
            complexity="simple",
            tables=[],
            images=[],
            element_order=[]
        )
