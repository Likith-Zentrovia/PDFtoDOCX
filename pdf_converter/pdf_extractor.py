"""
PDF Content Extractor using PyMuPDF

Extracts all content from PDF with accurate positioning and formatting:
- Text blocks with font info (name, size, bold, italic, color)
- Tables with cell structure and formatting
- Images with positions
- Unified element list ordered by Y-position for accurate layout reconstruction

This does the actual extraction work - Vision only provides layout hints.
"""

import fitz  # PyMuPDF
from typing import List, Dict, Tuple, Optional, Any, Union
from dataclasses import dataclass, field
from enum import Enum
from .layout_analyzer import LayoutInfo


class ElementType(Enum):
    """Type of page element."""
    TEXT = "text"
    IMAGE = "image"
    TABLE = "table"


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
    line_spacing: float = 1.0  # Relative line spacing

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

    @property
    def y_position(self) -> float:
        """Top Y position of this block."""
        return self.bbox[1]


@dataclass
class TableCell:
    """A single cell in a table."""
    text: str
    bbox: Tuple[float, float, float, float]
    row: int
    col: int
    rowspan: int = 1
    colspan: int = 1
    is_header: bool = False
    font_size: float = 11.0
    is_bold: bool = False
    alignment: str = "left"  # "left", "center", "right"


@dataclass
class TableInfo:
    """Information about an extracted table."""
    bbox: Tuple[float, float, float, float]
    cells: List[TableCell]
    num_rows: int
    num_cols: int
    has_header_row: bool = False

    @property
    def y_position(self) -> float:
        """Top Y position of this table."""
        return self.bbox[1]


@dataclass
class ImageInfo:
    """Information about an extracted image."""
    bbox: Tuple[float, float, float, float]
    data: bytes
    ext: str  # "png", "jpeg", etc.
    width: float
    height: float
    position_type: str = "block"  # "block", "inline", "float_left", "float_right"

    @property
    def y_position(self) -> float:
        """Top Y position of this image."""
        return self.bbox[1]


@dataclass
class PageElement:
    """A unified page element for proper ordering."""
    element_type: ElementType
    element: Union[TextBlock, ImageInfo, TableInfo]
    y_position: float  # For sorting
    x_position: float  # For column detection
    column: int = 1


@dataclass
class PageContent:
    """All content from a single page."""
    page_num: int
    width: float
    height: float
    text_blocks: List[TextBlock]
    images: List[ImageInfo]
    tables: List[TableInfo]
    reading_order: List[int]  # Indices of text_blocks in reading order
    # Unified list of all elements in Y-order for proper layout reconstruction
    elements: List[PageElement] = field(default_factory=list)


class PDFExtractor:
    """
    Extracts content from PDF using PyMuPDF.

    Uses layout information to properly assign columns and reading order.
    Creates unified element list for accurate layout reconstruction.
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

        # Extract tables first (so we can exclude table regions from text extraction)
        tables = self._extract_tables(page, layout)
        table_regions = [t.bbox for t in tables]

        # Extract text blocks (excluding table regions to avoid duplication)
        text_blocks = self._extract_text_blocks(page, layout, width, height, table_regions)

        # Assign columns based on layout
        if layout and layout.num_columns > 1:
            text_blocks = self._assign_columns(text_blocks, layout, width)

        # Determine reading order for text blocks
        reading_order = self._determine_reading_order(text_blocks, layout)

        # Extract images
        images = self._extract_images(page)

        # Create unified element list ordered by Y-position
        elements = self._create_element_list(text_blocks, images, tables, layout, width)

        return PageContent(
            page_num=page_num,
            width=width,
            height=height,
            text_blocks=text_blocks,
            images=images,
            tables=tables,
            reading_order=reading_order,
            elements=elements
        )

    def _extract_tables(self, page: fitz.Page, layout: Optional[LayoutInfo]) -> List[TableInfo]:
        """Extract tables from page using PyMuPDF's table detection."""
        tables = []

        try:
            # Use PyMuPDF's built-in table finder
            tab_finder = page.find_tables()

            for tab in tab_finder.tables:
                bbox = tuple(tab.bbox)
                cells = []

                # Extract table data
                table_data = tab.extract()
                num_rows = len(table_data)
                num_cols = len(table_data[0]) if table_data else 0

                # Get cell information
                for row_idx, row in enumerate(table_data):
                    for col_idx, cell_text in enumerate(row):
                        if cell_text is None:
                            cell_text = ""

                        # Try to get cell bbox (approximate based on table structure)
                        cell_width = (bbox[2] - bbox[0]) / max(num_cols, 1)
                        cell_height = (bbox[3] - bbox[1]) / max(num_rows, 1)
                        cell_bbox = (
                            bbox[0] + col_idx * cell_width,
                            bbox[1] + row_idx * cell_height,
                            bbox[0] + (col_idx + 1) * cell_width,
                            bbox[1] + (row_idx + 1) * cell_height
                        )

                        cells.append(TableCell(
                            text=str(cell_text).strip(),
                            bbox=cell_bbox,
                            row=row_idx,
                            col=col_idx,
                            is_header=(row_idx == 0),  # Assume first row is header
                            font_size=10.0,
                            is_bold=(row_idx == 0),
                            alignment="left"
                        ))

                # Determine if first row is a header (often has different formatting)
                has_header = num_rows > 1

                tables.append(TableInfo(
                    bbox=bbox,
                    cells=cells,
                    num_rows=num_rows,
                    num_cols=num_cols,
                    has_header_row=has_header
                ))

        except Exception as e:
            # Table extraction is optional, continue without tables
            pass

        return tables

    def _is_in_table_region(self, bbox: Tuple[float, float, float, float], table_regions: List[Tuple]) -> bool:
        """Check if a bbox is inside any table region."""
        for table_bbox in table_regions:
            # Check if centers overlap significantly
            block_center_y = (bbox[1] + bbox[3]) / 2
            table_top = table_bbox[1]
            table_bottom = table_bbox[3]

            if table_top <= block_center_y <= table_bottom:
                # Check horizontal overlap too
                block_center_x = (bbox[0] + bbox[2]) / 2
                table_left = table_bbox[0]
                table_right = table_bbox[2]

                if table_left <= block_center_x <= table_right:
                    return True
        return False

    def _create_element_list(
        self,
        text_blocks: List[TextBlock],
        images: List[ImageInfo],
        tables: List[TableInfo],
        layout: Optional[LayoutInfo],
        page_width: float
    ) -> List[PageElement]:
        """Create unified element list ordered by Y-position for proper layout reconstruction."""
        elements = []

        # Add text blocks
        for block in text_blocks:
            elements.append(PageElement(
                element_type=ElementType.TEXT,
                element=block,
                y_position=block.y_position,
                x_position=block.bbox[0],
                column=block.column
            ))

        # Add images
        for img in images:
            # Determine column for image based on x-position
            col = 1
            if layout and layout.num_columns > 1:
                x_pct = ((img.bbox[0] + img.bbox[2]) / 2) / page_width * 100
                for i, bound in enumerate(layout.column_boundaries):
                    if bound.get("x_start", 0) <= x_pct <= bound.get("x_end", 100):
                        col = i + 1
                        break

            elements.append(PageElement(
                element_type=ElementType.IMAGE,
                element=img,
                y_position=img.y_position,
                x_position=img.bbox[0],
                column=col
            ))

        # Add tables
        for table in tables:
            elements.append(PageElement(
                element_type=ElementType.TABLE,
                element=table,
                y_position=table.y_position,
                x_position=table.bbox[0],
                column=0  # Tables typically span full width or are treated specially
            ))

        # Sort by Y-position, then by column for same Y
        elements.sort(key=lambda e: (e.y_position, e.column, e.x_position))

        return elements

    def _extract_text_blocks(
        self,
        page: fitz.Page,
        layout: Optional[LayoutInfo],
        page_width: float,
        page_height: float,
        table_regions: List[Tuple] = None
    ) -> List[TextBlock]:
        """Extract text blocks with formatting information."""
        blocks = []
        table_regions = table_regions or []

        # Get detailed text with formatting
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:  # Skip non-text blocks
                continue

            bbox = tuple(block["bbox"])

            # Skip text that's inside table regions (already extracted as table)
            if table_regions and self._is_in_table_region(bbox, table_regions):
                continue

            spans = []
            line_heights = []

            for line in block.get("lines", []):
                line_bbox = line.get("bbox", bbox)
                line_heights.append(line_bbox[3] - line_bbox[1])

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

                # Calculate average line spacing
                avg_line_height = sum(line_heights) / len(line_heights) if line_heights else 12
                primary_font_size = max([s.font_size for s in spans], default=11)
                line_spacing = avg_line_height / primary_font_size if primary_font_size > 0 else 1.0

                blocks.append(TextBlock(
                    spans=spans,
                    bbox=bbox,
                    is_header=is_header,
                    is_footer=is_footer,
                    line_spacing=max(1.0, min(3.0, line_spacing))  # Clamp to reasonable range
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
