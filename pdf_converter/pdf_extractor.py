"""
PDF Content Extractor using PyMuPDF

Robust extraction with:
- Visual ordering based on actual positions (not PDF internal order)
- Automatic column detection from text positions
- Proper table extraction with structure
- Images with exact bounding boxes
- Line-level text extraction for accuracy

NO AI DEPENDENCY - pure Python analysis.
"""

import fitz  # PyMuPDF
from typing import List, Dict, Tuple, Optional, Any, Union
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict
import re


class ElementType(Enum):
    """Type of page element."""
    TEXT = "text"
    IMAGE = "image"
    TABLE = "table"


@dataclass
class TextLine:
    """A single line of text with formatting."""
    text: str
    bbox: Tuple[float, float, float, float]  # x0, y0, x1, y1
    font_name: str
    font_size: float
    is_bold: bool
    is_italic: bool
    color: Tuple[int, int, int]
    
    @property
    def y_center(self) -> float:
        return (self.bbox[1] + self.bbox[3]) / 2
    
    @property
    def x_center(self) -> float:
        return (self.bbox[0] + self.bbox[2]) / 2


@dataclass 
class TextBlock:
    """A block of text (one or more lines that belong together)."""
    lines: List[TextLine]
    bbox: Tuple[float, float, float, float]
    column: int = 0  # 0 = full width, 1 = left, 2 = right, etc.
    is_header: bool = False
    is_footer: bool = False
    
    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)
    
    @property
    def y_position(self) -> float:
        return self.bbox[1]
    
    @property
    def x_position(self) -> float:
        return self.bbox[0]
    
    @property
    def primary_font_size(self) -> float:
        if not self.lines:
            return 11.0
        sizes = [line.font_size for line in self.lines]
        return max(set(sizes), key=sizes.count)
    
    @property
    def primary_font_name(self) -> str:
        if not self.lines:
            return ""
        return self.lines[0].font_name
    
    @property
    def is_bold(self) -> bool:
        return any(line.is_bold for line in self.lines)
    
    @property
    def is_italic(self) -> bool:
        return any(line.is_italic for line in self.lines)
    
    @property
    def color(self) -> Tuple[int, int, int]:
        if self.lines:
            return self.lines[0].color
        return (0, 0, 0)


@dataclass
class TableCell:
    """A single cell in a table."""
    text: str
    bbox: Tuple[float, float, float, float]
    row: int
    col: int
    rowspan: int = 1
    colspan: int = 1
    is_bold: bool = False
    font_size: float = 10.0
    background_color: Optional[Tuple[int, int, int]] = None


@dataclass
class TableInfo:
    """Information about an extracted table."""
    bbox: Tuple[float, float, float, float]
    cells: List[List[TableCell]]  # 2D array [row][col]
    num_rows: int
    num_cols: int
    
    @property
    def y_position(self) -> float:
        return self.bbox[1]


@dataclass
class ImageInfo:
    """Information about an extracted image."""
    bbox: Tuple[float, float, float, float]
    data: bytes
    ext: str
    width: float
    height: float
    
    @property
    def y_position(self) -> float:
        return self.bbox[1]


@dataclass
class PageElement:
    """A unified page element for ordering."""
    element_type: ElementType
    element: Union[TextBlock, ImageInfo, TableInfo]
    bbox: Tuple[float, float, float, float]
    column: int = 0  # 0 = full width
    
    @property
    def y_position(self) -> float:
        return self.bbox[1]
    
    @property
    def x_position(self) -> float:
        return self.bbox[0]


@dataclass
class ColumnInfo:
    """Information about detected columns."""
    num_columns: int
    boundaries: List[Tuple[float, float]]  # [(x_start, x_end), ...]
    gutter_positions: List[float]  # X positions of gutters between columns


@dataclass
class PageContent:
    """All content from a single page."""
    page_num: int
    width: float
    height: float
    elements: List[PageElement]  # All elements in reading order
    column_info: ColumnInfo
    text_blocks: List[TextBlock] = field(default_factory=list)  # For backward compatibility
    images: List[ImageInfo] = field(default_factory=list)
    tables: List[TableInfo] = field(default_factory=list)


class PDFExtractor:
    """
    Extracts content from PDF with accurate visual ordering.
    
    Key features:
    - Detects columns automatically from text positions
    - Extracts text line by line for accuracy
    - Handles tables properly
    - Places images at correct positions
    """
    
    # Thresholds for layout detection
    LINE_SPACING_THRESHOLD = 1.5  # Multiple of font size for paragraph break
    COLUMN_GAP_THRESHOLD = 30  # Minimum gap to consider as column separator
    HEADER_REGION_PCT = 0.08  # Top 8% of page
    FOOTER_REGION_PCT = 0.08  # Bottom 8% of page
    
    def __init__(self):
        self.doc: Optional[fitz.Document] = None
    
    def extract_document(self, pdf_path: str, layout_hints: List = None) -> List[PageContent]:
        """
        Extract all content from PDF.
        
        Args:
            pdf_path: Path to PDF file
            layout_hints: Optional layout hints (ignored - we detect automatically)
        
        Returns:
            List of PageContent for each page
        """
        self.doc = fitz.open(pdf_path)
        pages = []
        
        for page_num in range(len(self.doc)):
            page_content = self._extract_page(page_num)
            pages.append(page_content)
        
        self.doc.close()
        return pages
    
    def _extract_page(self, page_num: int) -> PageContent:
        """Extract all content from a single page."""
        page = self.doc[page_num]
        width = page.rect.width
        height = page.rect.height
        
        # Step 1: Extract all text lines with formatting
        text_lines = self._extract_text_lines(page)
        
        # Step 2: Detect table regions FIRST (so we can exclude them from text)
        tables, table_bboxes = self._extract_tables(page)
        
        # Step 3: Filter text lines that are inside tables
        text_lines = [line for line in text_lines 
                      if not self._is_inside_any_bbox(line.bbox, table_bboxes)]
        
        # Step 4: Detect columns from text positions
        column_info = self._detect_columns(text_lines, width)
        
        # Step 5: Assign columns to text lines
        text_lines = self._assign_columns_to_lines(text_lines, column_info)
        
        # Step 6: Group text lines into blocks
        text_blocks = self._group_lines_into_blocks(text_lines, height, column_info)
        
        # Step 7: Mark headers and footers
        text_blocks = self._mark_headers_footers(text_blocks, height)
        
        # Step 8: Extract images
        images = self._extract_images(page)
        
        # Step 9: Create unified element list in reading order
        elements = self._create_reading_order(text_blocks, images, tables, column_info, width)
        
        return PageContent(
            page_num=page_num,
            width=width,
            height=height,
            elements=elements,
            column_info=column_info,
            text_blocks=text_blocks,
            images=images,
            tables=tables
        )
    
    def _extract_text_lines(self, page: fitz.Page) -> List[TextLine]:
        """Extract text as individual lines with formatting."""
        lines = []
        
        # Use "dict" mode for detailed text info
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        
        for block in blocks:
            if block.get("type") != 0:  # Skip non-text blocks
                continue
            
            for line in block.get("lines", []):
                line_text = ""
                line_bbox = list(line["bbox"])
                font_name = ""
                font_size = 11.0
                is_bold = False
                is_italic = False
                color = (0, 0, 0)
                
                for span in line.get("spans", []):
                    span_text = span.get("text", "")
                    if not span_text.strip():
                        continue
                    
                    line_text += span_text
                    
                    # Get formatting from first non-empty span
                    if not font_name:
                        font_name = span.get("font", "")
                        font_size = span.get("size", 11.0)
                        flags = span.get("flags", 0)
                        is_bold = bool(flags & 2**4) or "bold" in font_name.lower()
                        is_italic = bool(flags & 2**1) or "italic" in font_name.lower()
                        
                        color_int = span.get("color", 0)
                        if isinstance(color_int, int):
                            color = (
                                (color_int >> 16) & 0xFF,
                                (color_int >> 8) & 0xFF,
                                color_int & 0xFF
                            )
                
                if line_text.strip():
                    lines.append(TextLine(
                        text=line_text,
                        bbox=tuple(line_bbox),
                        font_name=font_name,
                        font_size=font_size,
                        is_bold=is_bold,
                        is_italic=is_italic,
                        color=color
                    ))
        
        return lines
    
    def _extract_tables(self, page: fitz.Page) -> Tuple[List[TableInfo], List[Tuple]]:
        """Extract tables using PyMuPDF's table detection."""
        tables = []
        table_bboxes = []
        
        try:
            tab_finder = page.find_tables()
            
            for tab in tab_finder.tables:
                bbox = tuple(tab.bbox)
                table_bboxes.append(bbox)
                
                # Extract table data
                data = tab.extract()
                if not data:
                    continue
                
                num_rows = len(data)
                num_cols = max(len(row) for row in data) if data else 0
                
                if num_rows == 0 or num_cols == 0:
                    continue
                
                # Create cell grid
                cells = []
                cell_height = (bbox[3] - bbox[1]) / num_rows
                cell_width = (bbox[2] - bbox[0]) / num_cols
                
                for row_idx, row in enumerate(data):
                    row_cells = []
                    for col_idx in range(num_cols):
                        cell_text = row[col_idx] if col_idx < len(row) else ""
                        if cell_text is None:
                            cell_text = ""
                        
                        cell_bbox = (
                            bbox[0] + col_idx * cell_width,
                            bbox[1] + row_idx * cell_height,
                            bbox[0] + (col_idx + 1) * cell_width,
                            bbox[1] + (row_idx + 1) * cell_height
                        )
                        
                        row_cells.append(TableCell(
                            text=str(cell_text).strip(),
                            bbox=cell_bbox,
                            row=row_idx,
                            col=col_idx,
                            is_bold=(row_idx == 0),  # Header row
                            font_size=10.0
                        ))
                    cells.append(row_cells)
                
                tables.append(TableInfo(
                    bbox=bbox,
                    cells=cells,
                    num_rows=num_rows,
                    num_cols=num_cols
                ))
        
        except Exception as e:
            pass  # Table extraction is optional
        
        return tables, table_bboxes
    
    def _is_inside_any_bbox(self, inner: Tuple, outers: List[Tuple]) -> bool:
        """Check if inner bbox is inside any of the outer bboxes."""
        for outer in outers:
            # Check if center of inner is inside outer
            center_x = (inner[0] + inner[2]) / 2
            center_y = (inner[1] + inner[3]) / 2
            
            if (outer[0] <= center_x <= outer[2] and 
                outer[1] <= center_y <= outer[3]):
                return True
        return False
    
    def _detect_columns(self, lines: List[TextLine], page_width: float) -> ColumnInfo:
        """
        Detect column structure from text line positions.
        
        This is the KEY function for accurate layout detection.
        """
        if not lines:
            return ColumnInfo(
                num_columns=1,
                boundaries=[(0, page_width)],
                gutter_positions=[]
            )
        
        # Collect all X positions (left edges of text)
        x_positions = [line.bbox[0] for line in lines]
        x_right_positions = [line.bbox[2] for line in lines]
        
        # Find gaps in horizontal coverage
        # Sort all x_start positions and find clusters
        x_starts = sorted(set(round(x, 0) for x in x_positions))
        
        # Look for significant gaps that could be column gutters
        margin_left = min(x_positions) if x_positions else 50
        margin_right = page_width - max(x_right_positions) if x_right_positions else 50
        
        # Calculate content width
        content_left = margin_left
        content_right = page_width - margin_right
        content_width = content_right - content_left
        
        # Analyze horizontal distribution of text
        # Divide page into bins and count text in each bin
        num_bins = 20
        bin_width = page_width / num_bins
        bin_counts = [0] * num_bins
        
        for line in lines:
            center_x = (line.bbox[0] + line.bbox[2]) / 2
            bin_idx = min(int(center_x / bin_width), num_bins - 1)
            bin_counts[bin_idx] += 1
        
        # Find gutters (bins with very low counts surrounded by higher counts)
        gutters = []
        for i in range(2, num_bins - 2):
            # Check if this bin is a gutter (low count with content on both sides)
            if bin_counts[i] <= 2:  # Low count
                left_has_content = any(bin_counts[j] > 3 for j in range(max(0, i-3), i))
                right_has_content = any(bin_counts[j] > 3 for j in range(i+1, min(num_bins, i+4)))
                
                if left_has_content and right_has_content:
                    gutter_x = (i + 0.5) * bin_width
                    # Verify this is actually a gap in the text
                    is_gap = True
                    for line in lines:
                        if line.bbox[0] < gutter_x < line.bbox[2]:
                            is_gap = False
                            break
                    if is_gap:
                        gutters.append(gutter_x)
        
        # Merge nearby gutters
        merged_gutters = []
        for g in sorted(gutters):
            if not merged_gutters or g - merged_gutters[-1] > self.COLUMN_GAP_THRESHOLD:
                merged_gutters.append(g)
        
        # Determine number of columns
        if not merged_gutters:
            return ColumnInfo(
                num_columns=1,
                boundaries=[(content_left, content_right)],
                gutter_positions=[]
            )
        
        # Create column boundaries
        num_columns = len(merged_gutters) + 1
        boundaries = []
        
        prev_x = content_left
        for gutter in merged_gutters:
            boundaries.append((prev_x, gutter))
            prev_x = gutter
        boundaries.append((prev_x, content_right))
        
        return ColumnInfo(
            num_columns=num_columns,
            boundaries=boundaries,
            gutter_positions=merged_gutters
        )
    
    def _assign_columns_to_lines(self, lines: List[TextLine], column_info: ColumnInfo) -> List[TextLine]:
        """Assign each line to a column based on its X position."""
        if column_info.num_columns == 1:
            return lines
        
        for line in lines:
            center_x = (line.bbox[0] + line.bbox[2]) / 2
            
            # Find which column this line belongs to
            for col_idx, (x_start, x_end) in enumerate(column_info.boundaries):
                if x_start <= center_x <= x_end:
                    # Store column in a way we can access later
                    line._column = col_idx + 1
                    break
            else:
                line._column = 1
        
        return lines
    
    def _group_lines_into_blocks(
        self, 
        lines: List[TextLine], 
        page_height: float,
        column_info: ColumnInfo
    ) -> List[TextBlock]:
        """
        Group text lines into logical blocks (paragraphs).
        
        Lines are grouped if they:
        - Are in the same column
        - Are vertically close (within LINE_SPACING_THRESHOLD * font_size)
        - Have similar X alignment
        """
        if not lines:
            return []
        
        # Sort lines by column, then by Y position
        def sort_key(line):
            col = getattr(line, '_column', 1)
            return (col, line.bbox[1], line.bbox[0])
        
        sorted_lines = sorted(lines, key=sort_key)
        
        blocks = []
        current_block_lines = []
        current_column = None
        
        for line in sorted_lines:
            line_column = getattr(line, '_column', 1)
            
            if not current_block_lines:
                current_block_lines = [line]
                current_column = line_column
                continue
            
            prev_line = current_block_lines[-1]
            
            # Check if we should start a new block
            start_new_block = False
            
            # Different column = new block
            if line_column != current_column:
                start_new_block = True
            
            # Large vertical gap = new block
            vertical_gap = line.bbox[1] - prev_line.bbox[3]
            threshold = prev_line.font_size * self.LINE_SPACING_THRESHOLD
            if vertical_gap > threshold:
                start_new_block = True
            
            # Very different X position (not continuation) = might be new block
            x_diff = abs(line.bbox[0] - prev_line.bbox[0])
            if x_diff > 50 and vertical_gap > prev_line.font_size * 0.5:
                start_new_block = True
            
            if start_new_block:
                # Save current block
                if current_block_lines:
                    blocks.append(self._create_block(current_block_lines, current_column))
                current_block_lines = [line]
                current_column = line_column
            else:
                current_block_lines.append(line)
        
        # Don't forget the last block
        if current_block_lines:
            blocks.append(self._create_block(current_block_lines, current_column))
        
        return blocks
    
    def _create_block(self, lines: List[TextLine], column: int) -> TextBlock:
        """Create a TextBlock from a list of lines."""
        # Calculate bounding box
        x0 = min(line.bbox[0] for line in lines)
        y0 = min(line.bbox[1] for line in lines)
        x1 = max(line.bbox[2] for line in lines)
        y1 = max(line.bbox[3] for line in lines)
        
        return TextBlock(
            lines=lines,
            bbox=(x0, y0, x1, y1),
            column=column
        )
    
    def _mark_headers_footers(self, blocks: List[TextBlock], page_height: float) -> List[TextBlock]:
        """Mark blocks that are likely headers or footers."""
        header_threshold = page_height * self.HEADER_REGION_PCT
        footer_threshold = page_height * (1 - self.FOOTER_REGION_PCT)
        
        for block in blocks:
            if block.bbox[3] < header_threshold:  # Bottom of block in header region
                block.is_header = True
            elif block.bbox[1] > footer_threshold:  # Top of block in footer region
                block.is_footer = True
        
        return blocks
    
    def _extract_images(self, page: fitz.Page) -> List[ImageInfo]:
        """Extract images with their positions."""
        images = []
        
        try:
            image_list = page.get_images(full=True)
            
            for img_info in image_list:
                xref = img_info[0]
                
                # Get image rectangle on page
                rects = page.get_image_rects(xref)
                if not rects:
                    continue
                
                bbox = tuple(rects[0])
                
                # Extract image data
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
        
        except Exception:
            pass
        
        return images
    
    def _create_reading_order(
        self,
        text_blocks: List[TextBlock],
        images: List[ImageInfo],
        tables: List[TableInfo],
        column_info: ColumnInfo,
        page_width: float
    ) -> List[PageElement]:
        """
        Create a list of all elements in proper reading order.
        
        Reading order rules:
        1. Headers first (top to bottom)
        2. For multi-column: read left column top to bottom, then right column
        3. For single column: top to bottom
        4. Images and tables inserted at their Y positions
        5. Footers last
        """
        elements = []
        
        # Separate headers, footers, and content
        headers = [b for b in text_blocks if b.is_header]
        footers = [b for b in text_blocks if b.is_footer]
        content = [b for b in text_blocks if not b.is_header and not b.is_footer]
        
        # Sort headers by Y
        headers.sort(key=lambda b: b.y_position)
        for block in headers:
            elements.append(PageElement(
                element_type=ElementType.TEXT,
                element=block,
                bbox=block.bbox,
                column=0
            ))
        
        # Handle content based on column structure
        if column_info.num_columns == 1:
            # Single column: combine all elements and sort by Y
            all_content = []
            
            for block in content:
                all_content.append(('text', block.y_position, block, block.bbox))
            
            for img in images:
                all_content.append(('image', img.y_position, img, img.bbox))
            
            for table in tables:
                all_content.append(('table', table.y_position, table, table.bbox))
            
            # Sort by Y position
            all_content.sort(key=lambda x: x[1])
            
            for item_type, _, item, bbox in all_content:
                if item_type == 'text':
                    elements.append(PageElement(
                        element_type=ElementType.TEXT,
                        element=item,
                        bbox=bbox,
                        column=0
                    ))
                elif item_type == 'image':
                    elements.append(PageElement(
                        element_type=ElementType.IMAGE,
                        element=item,
                        bbox=bbox,
                        column=0
                    ))
                elif item_type == 'table':
                    elements.append(PageElement(
                        element_type=ElementType.TABLE,
                        element=item,
                        bbox=bbox,
                        column=0
                    ))
        
        else:
            # Multi-column: process column by column
            # First, assign images and tables to columns
            def get_column(bbox, col_info):
                center_x = (bbox[0] + bbox[2]) / 2
                for col_idx, (x_start, x_end) in enumerate(col_info.boundaries):
                    if x_start <= center_x <= x_end:
                        return col_idx + 1
                return 0  # Full width
            
            # Organize content by column
            column_content = defaultdict(list)
            full_width_items = []
            
            for block in content:
                col = block.column
                if col == 0:
                    full_width_items.append(('text', block.y_position, block))
                else:
                    column_content[col].append(('text', block.y_position, block))
            
            for img in images:
                # Check if image spans multiple columns (wide image)
                img_col = get_column(img.bbox, column_info)
                img_width_ratio = img.width / page_width
                if img_width_ratio > 0.6:  # Wide image
                    full_width_items.append(('image', img.y_position, img))
                else:
                    column_content[img_col].append(('image', img.y_position, img))
            
            for table in tables:
                # Tables usually span full width or are in a specific column
                table_col = get_column(table.bbox, column_info)
                table_width = table.bbox[2] - table.bbox[0]
                if table_width / page_width > 0.6:
                    full_width_items.append(('table', table.y_position, table))
                else:
                    column_content[table_col].append(('table', table.y_position, table))
            
            # Sort each column by Y
            for col in column_content:
                column_content[col].sort(key=lambda x: x[1])
            full_width_items.sort(key=lambda x: x[1])
            
            # Interleave columns with full-width items
            # Process in Y order, switching between columns at appropriate points
            
            # Get all Y positions where we need to insert full-width items
            fw_positions = [item[1] for item in full_width_items]
            
            current_y = 0
            for fw_y in fw_positions + [float('inf')]:
                # Add column content up to this Y position
                for col in sorted(column_content.keys()):
                    for item_type, y_pos, item in column_content[col]:
                        if current_y <= y_pos < fw_y:
                            bbox = item.bbox if hasattr(item, 'bbox') else (0, y_pos, page_width, y_pos + 10)
                            elements.append(PageElement(
                                element_type=ElementType[item_type.upper()],
                                element=item,
                                bbox=bbox,
                                column=col
                            ))
                
                # Add full-width item at this position
                for item_type, y_pos, item in full_width_items:
                    if y_pos == fw_y:
                        bbox = item.bbox
                        elements.append(PageElement(
                            element_type=ElementType[item_type.upper()],
                            element=item,
                            bbox=bbox,
                            column=0
                        ))
                
                current_y = fw_y
        
        # Add footers last
        footers.sort(key=lambda b: b.y_position)
        for block in footers:
            elements.append(PageElement(
                element_type=ElementType.TEXT,
                element=block,
                bbox=block.bbox,
                column=0
            ))
        
        return elements
