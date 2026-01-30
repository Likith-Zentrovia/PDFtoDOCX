"""
Advanced PDF Structure Analyzer

This module provides detailed analysis of PDF structure including:
- Column detection and layout analysis
- Text block clustering and reading order
- Header/footer detection
- Image and table positioning
"""

import fitz  # PyMuPDF
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any
from enum import Enum
import statistics


class LayoutType(Enum):
    """Page layout types."""
    SINGLE_COLUMN = "single_column"
    TWO_COLUMN = "two_column"
    THREE_COLUMN = "three_column"
    MULTI_COLUMN = "multi_column"
    MIXED = "mixed"
    COMPLEX = "complex"


@dataclass
class TextBlock:
    """Represents a text block in the PDF."""
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1
    text: str
    font_name: str = ""
    font_size: float = 12.0
    font_flags: int = 0  # bold, italic, etc.
    color: Tuple[float, float, float] = (0, 0, 0)
    block_no: int = 0
    line_count: int = 1

    @property
    def x0(self) -> float:
        return self.bbox[0]

    @property
    def y0(self) -> float:
        return self.bbox[1]

    @property
    def x1(self) -> float:
        return self.bbox[2]

    @property
    def y1(self) -> float:
        return self.bbox[3]

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def is_bold(self) -> bool:
        return bool(self.font_flags & 2**4)

    @property
    def is_italic(self) -> bool:
        return bool(self.font_flags & 2**1)


@dataclass
class ImageBlock:
    """Represents an image in the PDF."""
    bbox: Tuple[float, float, float, float]
    image_data: bytes = b""
    image_ext: str = "png"
    xref: int = 0

    @property
    def width(self) -> float:
        return self.bbox[2] - self.bbox[0]

    @property
    def height(self) -> float:
        return self.bbox[3] - self.bbox[1]


@dataclass
class Column:
    """Represents a column in the layout."""
    x0: float
    x1: float
    blocks: List[TextBlock] = field(default_factory=list)

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def center(self) -> float:
        return (self.x0 + self.x1) / 2


@dataclass
class PageLayout:
    """Analyzed layout of a single page."""
    page_num: int
    width: float
    height: float
    layout_type: LayoutType
    columns: List[Column]
    text_blocks: List[TextBlock]
    images: List[ImageBlock]
    header_blocks: List[TextBlock] = field(default_factory=list)
    footer_blocks: List[TextBlock] = field(default_factory=list)
    margins: Dict[str, float] = field(default_factory=dict)
    reading_order: List[int] = field(default_factory=list)  # block indices in reading order


@dataclass
class DocumentStructure:
    """Complete document structure analysis."""
    page_count: int
    pages: List[PageLayout]
    dominant_layout: LayoutType
    has_consistent_layout: bool
    metadata: Dict[str, Any] = field(default_factory=dict)


class PDFAnalyzer:
    """
    Advanced PDF structure analyzer.

    Analyzes PDF documents to detect:
    - Column layouts (single, double, triple, etc.)
    - Reading order of text blocks
    - Headers and footers
    - Image positions
    - Page margins
    """

    def __init__(self, pdf_path: str):
        """Initialize analyzer with PDF file."""
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self._structure: Optional[DocumentStructure] = None

    def close(self):
        """Close the PDF document."""
        if self.doc:
            self.doc.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def analyze(self) -> DocumentStructure:
        """
        Perform complete document structure analysis.

        Returns:
            DocumentStructure with detailed layout information.
        """
        pages = []
        layout_types = []

        for page_num in range(len(self.doc)):
            page_layout = self._analyze_page(page_num)
            pages.append(page_layout)
            layout_types.append(page_layout.layout_type)

        # Determine dominant layout
        layout_counts = {}
        for lt in layout_types:
            layout_counts[lt] = layout_counts.get(lt, 0) + 1
        dominant_layout = max(layout_counts, key=layout_counts.get)

        # Check layout consistency
        has_consistent = layout_counts[dominant_layout] / len(layout_types) > 0.7

        self._structure = DocumentStructure(
            page_count=len(self.doc),
            pages=pages,
            dominant_layout=dominant_layout,
            has_consistent_layout=has_consistent,
            metadata=dict(self.doc.metadata) if self.doc.metadata else {}
        )

        return self._structure

    def _analyze_page(self, page_num: int) -> PageLayout:
        """Analyze a single page structure."""
        page = self.doc[page_num]
        width = page.rect.width
        height = page.rect.height

        # Extract text blocks with detailed information
        text_blocks = self._extract_text_blocks(page)

        # Extract images
        images = self._extract_images(page)

        # Detect margins
        margins = self._detect_margins(text_blocks, width, height)

        # Detect headers and footers
        header_blocks, footer_blocks, content_blocks = self._separate_header_footer(
            text_blocks, height, margins
        )

        # Detect columns from content blocks
        columns, layout_type = self._detect_columns(content_blocks, width, margins)

        # Determine reading order
        reading_order = self._determine_reading_order(content_blocks, columns, layout_type)

        return PageLayout(
            page_num=page_num,
            width=width,
            height=height,
            layout_type=layout_type,
            columns=columns,
            text_blocks=text_blocks,
            images=images,
            header_blocks=header_blocks,
            footer_blocks=footer_blocks,
            margins=margins,
            reading_order=reading_order
        )

    def _extract_text_blocks(self, page: fitz.Page) -> List[TextBlock]:
        """Extract text blocks with formatting information."""
        blocks = []

        # Get detailed text information
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        block_no = 0
        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:  # Skip non-text blocks
                continue

            bbox = tuple(block["bbox"])

            # Collect all text and font info from lines
            full_text = []
            fonts = []
            sizes = []
            flags_list = []
            colors = []
            line_count = 0

            for line in block.get("lines", []):
                line_count += 1
                for span in line.get("spans", []):
                    full_text.append(span.get("text", ""))
                    fonts.append(span.get("font", ""))
                    sizes.append(span.get("size", 12))
                    flags_list.append(span.get("flags", 0))
                    color = span.get("color", 0)
                    # Convert color int to RGB tuple
                    if isinstance(color, int):
                        r = (color >> 16) & 0xFF
                        g = (color >> 8) & 0xFF
                        b = color & 0xFF
                        colors.append((r/255, g/255, b/255))
                    else:
                        colors.append((0, 0, 0))

            text = "".join(full_text).strip()
            if not text:
                continue

            # Use most common font info
            font_name = max(set(fonts), key=fonts.count) if fonts else ""
            font_size = statistics.median(sizes) if sizes else 12
            font_flags = max(set(flags_list), key=flags_list.count) if flags_list else 0
            color = colors[0] if colors else (0, 0, 0)

            blocks.append(TextBlock(
                bbox=bbox,
                text=text,
                font_name=font_name,
                font_size=font_size,
                font_flags=font_flags,
                color=color,
                block_no=block_no,
                line_count=line_count
            ))
            block_no += 1

        return blocks

    def _extract_images(self, page: fitz.Page) -> List[ImageBlock]:
        """Extract images from page."""
        images = []

        image_list = page.get_images(full=True)

        for img_index, img in enumerate(image_list):
            xref = img[0]

            # Get image bounding box
            img_rects = page.get_image_rects(xref)
            if not img_rects:
                continue

            bbox = tuple(img_rects[0])

            # Extract image data
            try:
                base_image = self.doc.extract_image(xref)
                image_data = base_image.get("image", b"")
                image_ext = base_image.get("ext", "png")
            except Exception:
                image_data = b""
                image_ext = "png"

            images.append(ImageBlock(
                bbox=bbox,
                image_data=image_data,
                image_ext=image_ext,
                xref=xref
            ))

        return images

    def _detect_margins(
        self,
        blocks: List[TextBlock],
        page_width: float,
        page_height: float
    ) -> Dict[str, float]:
        """Detect page margins from text block positions."""
        if not blocks:
            return {"left": 72, "right": 72, "top": 72, "bottom": 72}

        # Find the extent of text blocks
        left_edges = [b.x0 for b in blocks]
        right_edges = [b.x1 for b in blocks]
        top_edges = [b.y0 for b in blocks]
        bottom_edges = [b.y1 for b in blocks]

        # Use percentile to avoid outliers
        left_margin = min(left_edges) if left_edges else 72
        right_margin = page_width - max(right_edges) if right_edges else 72
        top_margin = min(top_edges) if top_edges else 72
        bottom_margin = page_height - max(bottom_edges) if bottom_edges else 72

        return {
            "left": max(left_margin, 20),
            "right": max(right_margin, 20),
            "top": max(top_margin, 20),
            "bottom": max(bottom_margin, 20)
        }

    def _separate_header_footer(
        self,
        blocks: List[TextBlock],
        page_height: float,
        margins: Dict[str, float]
    ) -> Tuple[List[TextBlock], List[TextBlock], List[TextBlock]]:
        """Separate header and footer blocks from content."""
        header_threshold = margins["top"] + 50  # Within 50pt of top margin
        footer_threshold = page_height - margins["bottom"] - 50

        headers = []
        footers = []
        content = []

        for block in blocks:
            if block.y1 < header_threshold:
                headers.append(block)
            elif block.y0 > footer_threshold:
                footers.append(block)
            else:
                content.append(block)

        return headers, footers, content

    def _detect_columns(
        self,
        blocks: List[TextBlock],
        page_width: float,
        margins: Dict[str, float]
    ) -> Tuple[List[Column], LayoutType]:
        """Detect column layout from text blocks."""
        if not blocks:
            return [Column(x0=margins["left"], x1=page_width - margins["right"])], LayoutType.SINGLE_COLUMN

        # Content area width
        content_left = margins["left"]
        content_right = page_width - margins["right"]
        content_width = content_right - content_left

        # Analyze block x-positions to detect columns
        # Group blocks by their horizontal position
        x_centers = [b.center_x for b in blocks]

        if not x_centers:
            return [Column(x0=content_left, x1=content_right)], LayoutType.SINGLE_COLUMN

        # Use clustering to detect column centers
        column_positions = self._cluster_positions(x_centers, threshold=content_width * 0.15)

        num_columns = len(column_positions)

        if num_columns == 1:
            layout_type = LayoutType.SINGLE_COLUMN
        elif num_columns == 2:
            layout_type = LayoutType.TWO_COLUMN
        elif num_columns == 3:
            layout_type = LayoutType.THREE_COLUMN
        else:
            layout_type = LayoutType.MULTI_COLUMN

        # Create column objects
        columns = []
        sorted_positions = sorted(column_positions)

        for i, center in enumerate(sorted_positions):
            # Estimate column boundaries
            if i == 0:
                x0 = content_left
            else:
                x0 = (sorted_positions[i-1] + center) / 2

            if i == len(sorted_positions) - 1:
                x1 = content_right
            else:
                x1 = (center + sorted_positions[i+1]) / 2

            col = Column(x0=x0, x1=x1)

            # Assign blocks to column
            for block in blocks:
                if x0 <= block.center_x <= x1:
                    col.blocks.append(block)

            columns.append(col)

        # Verify layout - check if blocks actually span multiple columns
        wide_blocks = [b for b in blocks if b.width > content_width * 0.6]
        if wide_blocks and num_columns > 1:
            # Has wide blocks spanning columns - mixed layout
            layout_type = LayoutType.MIXED

        return columns, layout_type

    def _cluster_positions(self, positions: List[float], threshold: float) -> List[float]:
        """Cluster positions to find distinct column centers."""
        if not positions:
            return []

        sorted_pos = sorted(positions)
        clusters = []
        current_cluster = [sorted_pos[0]]

        for pos in sorted_pos[1:]:
            if pos - current_cluster[-1] < threshold:
                current_cluster.append(pos)
            else:
                clusters.append(statistics.mean(current_cluster))
                current_cluster = [pos]

        if current_cluster:
            clusters.append(statistics.mean(current_cluster))

        return clusters

    def _determine_reading_order(
        self,
        blocks: List[TextBlock],
        columns: List[Column],
        layout_type: LayoutType
    ) -> List[int]:
        """Determine the correct reading order of blocks."""
        if not blocks:
            return []

        if layout_type == LayoutType.SINGLE_COLUMN:
            # Simple top-to-bottom ordering
            sorted_blocks = sorted(enumerate(blocks), key=lambda x: (x[1].y0, x[1].x0))
            return [idx for idx, _ in sorted_blocks]

        # Multi-column: read column by column, top to bottom within each column
        reading_order = []
        block_indices = {id(b): i for i, b in enumerate(blocks)}

        for col in columns:
            # Sort blocks in this column by y position
            col_blocks = sorted(col.blocks, key=lambda b: b.y0)
            for block in col_blocks:
                idx = block_indices.get(id(block))
                if idx is not None and idx not in reading_order:
                    reading_order.append(idx)

        # Add any blocks that weren't assigned to columns
        for i, block in enumerate(blocks):
            if i not in reading_order:
                reading_order.append(i)

        return reading_order

    def get_page_text_in_reading_order(self, page_num: int) -> str:
        """Get page text in correct reading order."""
        if not self._structure:
            self.analyze()

        page_layout = self._structure.pages[page_num]
        ordered_text = []

        # Add header text
        for block in sorted(page_layout.header_blocks, key=lambda b: (b.y0, b.x0)):
            ordered_text.append(block.text)

        # Add content in reading order
        content_blocks = [b for b in page_layout.text_blocks
                        if b not in page_layout.header_blocks
                        and b not in page_layout.footer_blocks]

        for idx in page_layout.reading_order:
            if idx < len(content_blocks):
                ordered_text.append(content_blocks[idx].text)

        # Add footer text
        for block in sorted(page_layout.footer_blocks, key=lambda b: (b.y0, b.x0)):
            ordered_text.append(block.text)

        return "\n\n".join(ordered_text)

    def print_analysis_summary(self):
        """Print a summary of the document analysis."""
        if not self._structure:
            self.analyze()

        print(f"\n{'='*60}")
        print(f"PDF Analysis Summary: {self.pdf_path}")
        print(f"{'='*60}")
        print(f"Total Pages: {self._structure.page_count}")
        print(f"Dominant Layout: {self._structure.dominant_layout.value}")
        print(f"Consistent Layout: {self._structure.has_consistent_layout}")

        for page in self._structure.pages:
            print(f"\n  Page {page.page_num + 1}:")
            print(f"    Layout: {page.layout_type.value}")
            print(f"    Columns: {len(page.columns)}")
            print(f"    Text Blocks: {len(page.text_blocks)}")
            print(f"    Images: {len(page.images)}")
            print(f"    Headers: {len(page.header_blocks)}")
            print(f"    Footers: {len(page.footer_blocks)}")
