"""
PDF Layout Analyzer using Claude Vision

Uses Claude Vision ONLY to analyze page layout complexity:
- Number of columns
- Column boundaries (percentages)
- Header/footer regions
- General structure

Does NOT extract any text content - that's done by Python.
"""

import os
import base64
from typing import Optional, List, Dict, Any
from dataclasses import dataclass
import fitz  # PyMuPDF
import anthropic


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

    LAYOUT_PROMPT = """Analyze this PDF page layout ONLY. Do NOT extract or read any text content.

Return a JSON object with ONLY this structure:
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
    "complexity": "simple" | "moderate" | "complex"
}

Rules:
- num_columns: Count main content columns (1, 2, or 3)
- column_boundaries: X positions as percentages (0=left edge, 100=right edge)
- For single column: [{"x_start": 5, "x_end": 95}]
- For two columns: [{"x_start": 5, "x_end": 48}, {"x_start": 52, "x_end": 95}]
- header_height_pct: Height of header area as % of page (0 if no header)
- footer_height_pct: Height of footer area as % of page (0 if no footer)
- complexity: "simple"=single column, "moderate"=2 columns, "complex"=3+ or irregular

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
                # Use template with correct page number
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
                    complexity=template.complexity
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

        # Render at low resolution for quick analysis
        zoom = 1.0  # 72 DPI is enough for layout analysis
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        img_bytes = pix.tobytes("png")
        img_b64 = base64.standard_b64encode(img_bytes).decode("utf-8")

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=500,
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
                complexity=data.get("complexity", "simple")
            )

        except Exception as e:
            print(f"    Warning: Vision analysis failed for page {page_num + 1}, using default")
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
            complexity="simple"
        )
