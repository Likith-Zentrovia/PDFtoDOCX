"""
PDF Content Extractor using PyMuPDF

Extracts all content from PDF with accurate positioning and formatting:
- Text blocks with font info (name, size, bold, italic, color)
- Images with positions
- Reading order based on layout analysis

This does the actual extraction work - Vision only provides layout hints.
"""

import fitz  # PyMuPDF
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass, field
from .layout_analyzer import LayoutInfo


@dataclass
class TextSpan:
    """A text span with formatting."""
    text: str
    font_name: str
    font_size: float
    is_bold: bool
    is_italic: bool
    color: Tuple[int, int, int]  # RGB
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1


@dataclass
class TextBlock:
    """A block of text (paragraph-like unit)."""
    spans: List[TextSpan]
    bbox: Tuple[float, float, float, float]
    column: int = 1
    is_header: bool = False
    is_footer: bool = False

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.spans)

    @property
    def primary_font_size(self) -> float:
        if not self.spans:
            return 11.0
        sizes = [s.font_size for s in self.spans]
        return max(set(sizes), key=sizes.count)

    @property
    def is_bold(self) -> bool:
        if not self.spans:
            return False
        return any(s.is_bold for s in self.spans)

    @property
    def is_italic(self) -> bool:
        if not self.spans:
            return False
        return any(s.is_italic for s in self.spans)


@dataclass
class ImageInfo:
    """Information about an extracted image."""
    bbox: Tuple[float, float, float, float]
    data: bytes
    ext: str  # "png", "jpeg", etc.
    width: float
    height: float


@dataclass
class PageContent:
    """All content from a single page."""
    page_num: int
    width: float
    height: float
    text_blocks: List[TextBlock]
    images: List[ImageInfo]
    reading_order: List[int]  # Indices of text_blocks in reading order


class PDFExtractor:
    """
    Extracts content from PDF using PyMuPDF.

    Uses layout information to properly assign columns and reading order.
    """

    def __init__(self):
        self.doc: Optional[fitz.Document] = None

    def extract_document(
        self,
        pdf_path: str,
        layout_info: List[LayoutInfo]
    ) -> List[PageContent]:
        """
        Extract all content from PDF.

        Args:
            pdf_path: Path to PDF file
            layout_info: Layout analysis for each page

        Returns:
            List of PageContent for each page
        """
        self.doc = fitz.open(pdf_path)
        pages = []

        for page_num in range(len(self.doc)):
            layout = layout_info[page_num] if page_num < len(layout_info) else None
            page_content = self._extract_page(page_num, layout)
            pages.append(page_content)

        self.doc.close()
        return pages

    def _extract_page(self, page_num: int, layout: Optional[LayoutInfo]) -> PageContent:
        """Extract content from a single page."""
        page = self.doc[page_num]
        width = page.rect.width
        height = page.rect.height

        # Extract text blocks
        text_blocks = self._extract_text_blocks(page, layout, width, height)

        # Assign columns based on layout
        if layout and layout.num_columns > 1:
            text_blocks = self._assign_columns(text_blocks, layout, width)

        # Determine reading order
        reading_order = self._determine_reading_order(text_blocks, layout)

        # Extract images
        images = self._extract_images(page)

        return PageContent(
            page_num=page_num,
            width=width,
            height=height,
            text_blocks=text_blocks,
            images=images,
            reading_order=reading_order
        )

    def _extract_text_blocks(
        self,
        page: fitz.Page,
        layout: Optional[LayoutInfo],
        page_width: float,
        page_height: float
    ) -> List[TextBlock]:
        """Extract text blocks with formatting information."""
        blocks = []

        # Get detailed text with formatting
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:  # Skip non-text blocks
                continue

            bbox = tuple(block["bbox"])
            spans = []

            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text:
                        continue

                    font = span.get("font", "")
                    flags = span.get("flags", 0)

                    # Parse flags for bold/italic
                    is_bold = bool(flags & 2**4) or "bold" in font.lower()
                    is_italic = bool(flags & 2**1) or "italic" in font.lower() or "oblique" in font.lower()

                    # Parse color
                    color_int = span.get("color", 0)
                    if isinstance(color_int, int):
                        r = (color_int >> 16) & 0xFF
                        g = (color_int >> 8) & 0xFF
                        b = color_int & 0xFF
                    else:
                        r, g, b = 0, 0, 0

                    spans.append(TextSpan(
                        text=text,
                        font_name=font,
                        font_size=span.get("size", 11),
                        is_bold=is_bold,
                        is_italic=is_italic,
                        color=(r, g, b),
                        bbox=tuple(span.get("bbox", bbox))
                    ))

            if spans:
                # Check if header/footer based on layout
                is_header = False
                is_footer = False

                if layout:
                    y_pct = (bbox[1] / page_height) * 100
                    y_end_pct = (bbox[3] / page_height) * 100

                    if layout.has_header and y_pct < layout.header_height_pct:
                        is_header = True
                    if layout.has_footer and y_end_pct > (100 - layout.footer_height_pct):
                        is_footer = True

                blocks.append(TextBlock(
                    spans=spans,
                    bbox=bbox,
                    is_header=is_header,
                    is_footer=is_footer
                ))

        return blocks

    def _assign_columns(
        self,
        blocks: List[TextBlock],
        layout: LayoutInfo,
        page_width: float
    ) -> List[TextBlock]:
        """Assign each block to a column based on its position."""
        for block in blocks:
            if block.is_header or block.is_footer:
                block.column = 0  # Full width
                continue

            # Calculate block center X as percentage
            x_center = ((block.bbox[0] + block.bbox[2]) / 2) / page_width * 100

            # Find which column this belongs to
            for i, col in enumerate(layout.column_boundaries):
                x_start = col.get("x_start", 0)
                x_end = col.get("x_end", 100)

                if x_start <= x_center <= x_end:
                    block.column = i + 1
                    break
            else:
                # Default to closest column
                block.column = 1

        return blocks

    def _determine_reading_order(
        self,
        blocks: List[TextBlock],
        layout: Optional[LayoutInfo]
    ) -> List[int]:
        """Determine the correct reading order of blocks."""
        if not blocks:
            return []

        # Separate headers, content, footers
        headers = [(i, b) for i, b in enumerate(blocks) if b.is_header]
        footers = [(i, b) for i, b in enumerate(blocks) if b.is_footer]
        content = [(i, b) for i, b in enumerate(blocks) if not b.is_header and not b.is_footer]

        reading_order = []

        # Headers first (top to bottom)
        headers.sort(key=lambda x: x[1].bbox[1])
        reading_order.extend([i for i, _ in headers])

        # Content: by column, then by Y position
        if layout and layout.num_columns > 1:
            # Group by column
            columns: Dict[int, List[Tuple[int, TextBlock]]] = {}
            for i, block in content:
                col = block.column
                if col not in columns:
                    columns[col] = []
                columns[col].append((i, block))

            # Sort each column by Y, then add in column order
            for col_num in sorted(columns.keys()):
                col_blocks = columns[col_num]
                col_blocks.sort(key=lambda x: x[1].bbox[1])
                reading_order.extend([i for i, _ in col_blocks])
        else:
            # Single column: just sort by Y position
            content.sort(key=lambda x: x[1].bbox[1])
            reading_order.extend([i for i, _ in content])

        # Footers last
        footers.sort(key=lambda x: x[1].bbox[1])
        reading_order.extend([i for i, _ in footers])

        return reading_order

    def _extract_images(self, page: fitz.Page) -> List[ImageInfo]:
        """Extract images from page."""
        images = []
        image_list = page.get_images(full=True)

        for img_info in image_list:
            xref = img_info[0]

            # Get image rectangle
            rects = page.get_image_rects(xref)
            if not rects:
                continue

            bbox = tuple(rects[0])

            try:
                base_image = self.doc.extract_image(xref)
                img_data = base_image.get("image", b"")
                img_ext = base_image.get("ext", "png")

                if img_data:
                    images.append(ImageInfo(
                        bbox=bbox,
                        data=img_data,
                        ext=img_ext,
                        width=bbox[2] - bbox[0],
                        height=bbox[3] - bbox[1]
                    ))
            except Exception:
                continue

        return images
