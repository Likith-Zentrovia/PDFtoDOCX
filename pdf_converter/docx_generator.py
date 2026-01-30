"""
DOCX Generator

Creates accurate DOCX from extracted PDF content using layout information.
Preserves:
- Exact element ordering (text, images, tables interleaved by Y-position)
- Multi-column layouts (using tables)
- Text formatting (font size, bold, italic, color)
- Images with proper positioning
- Tables with structure and formatting
- Reading order and spacing
"""

import io
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

from .layout_analyzer import LayoutInfo, DocumentLayout
from .pdf_extractor import (
    PageContent, TextBlock, ImageInfo, TableInfo, TableCell,
    PageElement, ElementType
)


@dataclass
class GenerationResult:
    """Result of DOCX generation."""
    success: bool
    output_path: str
    pages_processed: int
    text_blocks_written: int
    images_added: int
    tables_added: int
    errors: List[str]


class DOCXGenerator:
    """
    Generates DOCX from extracted PDF content.

    Uses unified element list to create accurate layout matching the original PDF.
    Elements are placed in exact Y-order for proper document flow.
    """

    # Map PDF font sizes to reasonable DOCX sizes
    MIN_FONT_SIZE = 8
    MAX_FONT_SIZE = 36

    def __init__(self):
        self.doc: Optional[Document] = None
        self._text_count = 0
        self._image_count = 0
        self._table_count = 0
        self._errors: List[str] = []

    def generate(
        self,
        pages: List[PageContent],
        layout: DocumentLayout,
        output_path: str
    ) -> GenerationResult:
        """
        Generate DOCX from extracted content.

        Args:
            pages: Extracted content for each page
            layout: Layout analysis
            output_path: Output file path

        Returns:
            GenerationResult with details
        """
        self.doc = Document()
        self._text_count = 0
        self._image_count = 0
        self._table_count = 0
        self._errors = []

        # Setup document
        self._setup_document(pages[0] if pages else None)

        # Process each page
        for i, page in enumerate(pages):
            page_layout = layout.pages[i] if layout and i < len(layout.pages) else None

            if i > 0:
                self.doc.add_page_break()

            self._generate_page(page, page_layout)

        # Save
        try:
            self.doc.save(output_path)
            return GenerationResult(
                success=True,
                output_path=output_path,
                pages_processed=len(pages),
                text_blocks_written=self._text_count,
                images_added=self._image_count,
                tables_added=self._table_count,
                errors=self._errors
            )
        except Exception as e:
            return GenerationResult(
                success=False,
                output_path=output_path,
                pages_processed=len(pages),
                text_blocks_written=self._text_count,
                images_added=self._image_count,
                tables_added=self._table_count,
                errors=[str(e)]
            )

    def _setup_document(self, first_page: Optional[PageContent]):
        """Setup document properties."""
        if not first_page:
            return

        section = self.doc.sections[0]

        # Set page size based on PDF
        section.page_width = Twips(first_page.width * 20)
        section.page_height = Twips(first_page.height * 20)

        # Set narrower margins for better layout fidelity
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    def _generate_page(self, page: PageContent, layout: Optional[LayoutInfo]):
        """Generate content for a single page using unified element list."""
        num_columns = layout.num_columns if layout else 1

        if num_columns == 1 or not layout:
            self._generate_single_column_page(page)
        else:
            self._generate_multi_column_page(page, layout)

    def _generate_single_column_page(self, page: PageContent):
        """
        Generate single-column page with elements in Y-order.
        This is the key fix - elements are interleaved by position, not type.
        """
        # Use the unified element list - already sorted by Y-position
        if page.elements:
            prev_y = 0
            for elem in page.elements:
                # Add spacing based on vertical gap
                gap = elem.y_position - prev_y
                if prev_y > 0 and gap > 20:  # Significant gap
                    self._add_spacing(min(gap / 10, 24))  # Max 24pt spacing

                if elem.element_type == ElementType.TEXT:
                    self._add_text_block(elem.element)
                elif elem.element_type == ElementType.IMAGE:
                    self._add_image(elem.element, page.width)
                elif elem.element_type == ElementType.TABLE:
                    self._add_table(elem.element, page.width)

                prev_y = elem.y_position
        else:
            # Fallback to old method if no unified elements
            self._generate_single_column_fallback(page)

    def _generate_single_column_fallback(self, page: PageContent):
        """Fallback single-column generation (legacy behavior)."""
        # Combine all elements with their Y positions
        all_elements = []

        for idx in page.reading_order:
            if idx < len(page.text_blocks):
                block = page.text_blocks[idx]
                all_elements.append(('text', block.bbox[1], block))

        for img in page.images:
            all_elements.append(('image', img.bbox[1], img))

        for table in page.tables:
            all_elements.append(('table', table.bbox[1], table))

        # Sort by Y position
        all_elements.sort(key=lambda x: x[1])

        # Add elements in order
        prev_y = 0
        for elem_type, y_pos, elem in all_elements:
            # Add spacing based on gap
            gap = y_pos - prev_y
            if prev_y > 0 and gap > 20:
                self._add_spacing(min(gap / 10, 24))

            if elem_type == 'text':
                self._add_text_block(elem)
            elif elem_type == 'image':
                self._add_image(elem, page.width)
            elif elem_type == 'table':
                self._add_table(elem, page.width)

            prev_y = y_pos

    def _generate_multi_column_page(self, page: PageContent, layout: LayoutInfo):
        """Generate multi-column page content."""
        num_cols = layout.num_columns

        # Separate elements by type and position
        headers = []
        footers = []
        column_elements: Dict[int, List[PageElement]] = {i: [] for i in range(num_cols)}
        full_width_elements = []  # Tables and large images

        if page.elements:
            for elem in page.elements:
                if elem.element_type == ElementType.TEXT:
                    block = elem.element
                    if block.is_header:
                        headers.append(elem)
                    elif block.is_footer:
                        footers.append(elem)
                    else:
                        col = min(elem.column - 1, num_cols - 1)
                        col = max(0, col)
                        column_elements[col].append(elem)
                elif elem.element_type == ElementType.TABLE:
                    full_width_elements.append(elem)
                elif elem.element_type == ElementType.IMAGE:
                    # Decide if image goes in column or full width
                    img = elem.element
                    img_width_ratio = img.width / page.width
                    if img_width_ratio > 0.6:  # Wide image
                        full_width_elements.append(elem)
                    else:
                        col = min(elem.column - 1, num_cols - 1)
                        col = max(0, col)
                        column_elements[col].append(elem)
        else:
            # Fallback - use text_blocks and images directly
            for block in page.text_blocks:
                if block.is_header:
                    headers.append(PageElement(ElementType.TEXT, block, block.y_position, block.bbox[0], 0))
                elif block.is_footer:
                    footers.append(PageElement(ElementType.TEXT, block, block.y_position, block.bbox[0], 0))
                else:
                    col = min(block.column - 1, num_cols - 1)
                    col = max(0, col)
                    column_elements[col].append(
                        PageElement(ElementType.TEXT, block, block.y_position, block.bbox[0], col + 1)
                    )

            for img in page.images:
                full_width_elements.append(
                    PageElement(ElementType.IMAGE, img, img.y_position, img.bbox[0], 0)
                )

            for table in page.tables:
                full_width_elements.append(
                    PageElement(ElementType.TABLE, table, table.y_position, table.bbox[0], 0)
                )

        # Sort all element lists by Y position
        headers.sort(key=lambda e: e.y_position)
        footers.sort(key=lambda e: e.y_position)
        full_width_elements.sort(key=lambda e: e.y_position)
        for col in column_elements:
            column_elements[col].sort(key=lambda e: e.y_position)

        # Add headers first
        for elem in headers:
            self._add_element(elem, page.width)

        # Check if we need to interleave full-width elements with column content
        if full_width_elements:
            self._generate_interleaved_columns(column_elements, full_width_elements, layout, page)
        else:
            # Simple column table
            self._generate_column_table(column_elements, layout, page)

        # Add footers last
        for elem in footers:
            self._add_element(elem, page.width)

    def _generate_column_table(
        self,
        column_elements: Dict[int, List[PageElement]],
        layout: LayoutInfo,
        page: PageContent
    ):
        """Generate a table-based column layout."""
        num_cols = layout.num_columns

        # Create table for column layout
        table = self.doc.add_table(rows=1, cols=num_cols)
        table.autofit = False

        # Calculate column widths
        total_width = 7.0  # Content width in inches
        col_widths = self._calculate_column_widths(layout.column_boundaries, total_width, num_cols)

        # Fill columns
        for col_idx in range(num_cols):
            cell = table.rows[0].cells[col_idx]
            cell.width = Inches(col_widths[col_idx])

            for elem in column_elements.get(col_idx, []):
                if elem.element_type == ElementType.TEXT:
                    self._add_text_block_to_cell(cell, elem.element)
                elif elem.element_type == ElementType.IMAGE:
                    self._add_image_to_cell(cell, elem.element, page.width / num_cols)

        # Remove table borders
        self._remove_table_borders(table)

    def _generate_interleaved_columns(
        self,
        column_elements: Dict[int, List[PageElement]],
        full_width_elements: List[PageElement],
        layout: LayoutInfo,
        page: PageContent
    ):
        """Generate columns with full-width elements interspersed."""
        num_cols = layout.num_columns

        # Get all column content Y positions to determine split points
        all_y_positions = set()
        for col in column_elements.values():
            for elem in col:
                all_y_positions.add(elem.y_position)

        # Add full-width element positions as split points
        split_points = sorted([elem.y_position for elem in full_width_elements])

        if not split_points:
            self._generate_column_table(column_elements, layout, page)
            return

        # Generate content in segments
        current_y = 0

        for split_y in split_points:
            # Get column content before this split point
            segment_cols: Dict[int, List[PageElement]] = {i: [] for i in range(num_cols)}
            for col_idx, elems in column_elements.items():
                segment_cols[col_idx] = [e for e in elems if current_y <= e.y_position < split_y]

            # Add column segment if not empty
            if any(segment_cols.values()):
                self._generate_column_table(segment_cols, layout, page)

            # Add the full-width element at this position
            for elem in full_width_elements:
                if abs(elem.y_position - split_y) < 1:
                    self._add_element(elem, page.width)

            current_y = split_y

        # Add remaining column content after last split
        remaining_cols: Dict[int, List[PageElement]] = {i: [] for i in range(num_cols)}
        for col_idx, elems in column_elements.items():
            remaining_cols[col_idx] = [e for e in elems if e.y_position >= current_y]

        if any(remaining_cols.values()):
            self._generate_column_table(remaining_cols, layout, page)

    def _add_element(self, elem: PageElement, page_width: float):
        """Add any element type to the document."""
        if elem.element_type == ElementType.TEXT:
            self._add_text_block(elem.element)
        elif elem.element_type == ElementType.IMAGE:
            self._add_image(elem.element, page_width)
        elif elem.element_type == ElementType.TABLE:
            self._add_table(elem.element, page_width)

    def _add_spacing(self, points: float):
        """Add vertical spacing."""
        para = self.doc.add_paragraph()
        para.paragraph_format.space_before = Pt(points)
        para.paragraph_format.space_after = Pt(0)

    def _calculate_column_widths(
        self,
        boundaries: List[Dict[str, float]],
        total_width: float,
        num_cols: int = 2
    ) -> List[float]:
        """Calculate column widths from boundaries."""
        widths = []
        for bound in boundaries:
            x_start = bound.get("x_start", 0)
            x_end = bound.get("x_end", 100)
            width_pct = (x_end - x_start) / 100
            widths.append(width_pct * total_width)

        # Ensure we have exactly num_cols widths
        if len(widths) < num_cols:
            default_width = total_width / num_cols
            while len(widths) < num_cols:
                widths.append(default_width)
        elif len(widths) > num_cols:
            widths = widths[:num_cols]

        if not widths:
            return [total_width / num_cols] * num_cols

        return widths

    def _add_text_block(self, block: TextBlock):
        """Add a text block as a paragraph."""
        para = self.doc.add_paragraph()
        self._format_paragraph(para, block)
        self._text_count += 1

    def _add_text_block_to_cell(self, cell, block: TextBlock):
        """Add a text block to a table cell."""
        para = cell.add_paragraph()
        self._format_paragraph(para, block)
        self._text_count += 1

    def _format_paragraph(self, para, block: TextBlock):
        """Format a paragraph based on text block."""
        # Add text with formatting from spans
        for span in block.spans:
            run = para.add_run(span.text)

            # Font size (clamp to reasonable range)
            size = max(self.MIN_FONT_SIZE, min(self.MAX_FONT_SIZE, span.font_size))
            run.font.size = Pt(size)

            # Bold/Italic
            run.font.bold = span.is_bold
            run.font.italic = span.is_italic

            # Color (only if not black)
            if span.color != (0, 0, 0):
                run.font.color.rgb = RGBColor(*span.color)

            # Font name (clean it up)
            font_name = self._clean_font_name(span.font_name)
            if font_name:
                run.font.name = font_name

        # Paragraph spacing - preserve line spacing from original
        para.paragraph_format.space_after = Pt(4)
        para.paragraph_format.space_before = Pt(0)

        # Apply line spacing if available
        if hasattr(block, 'line_spacing') and block.line_spacing > 1.0:
            para.paragraph_format.line_spacing = block.line_spacing

    def _clean_font_name(self, font_name: str) -> str:
        """Clean up font name from PDF."""
        if not font_name:
            return ""

        # Remove common prefixes (subset font names)
        name = font_name
        if "+" in name:
            name = name.split("+", 1)[-1]

        # Map common variations
        mappings = {
            "ArialMT": "Arial",
            "Arial-BoldMT": "Arial",
            "Arial-ItalicMT": "Arial",
            "TimesNewRomanPSMT": "Times New Roman",
            "TimesNewRomanPS-BoldMT": "Times New Roman",
            "CourierNewPSMT": "Courier New",
        }

        return mappings.get(name, name.split("-")[0] if "-" in name else name)

    def _add_image(self, img: ImageInfo, page_width: float):
        """Add an image to the document at proper position."""
        if not img.data:
            return

        try:
            image_stream = io.BytesIO(img.data)

            # Calculate appropriate width while maintaining aspect ratio
            img_width_ratio = img.width / page_width
            max_width = min(img_width_ratio * 7, 6.5)  # Max 6.5 inches

            # Ensure minimum width
            max_width = max(max_width, 1.0)

            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

            run = para.add_run()
            run.add_picture(image_stream, width=Inches(max_width))

            # Minimal spacing after image
            para.paragraph_format.space_after = Pt(6)

            self._image_count += 1

        except Exception as e:
            self._errors.append(f"Failed to add image: {e}")

    def _add_image_to_cell(self, cell, img: ImageInfo, max_cell_width: float):
        """Add an image to a table cell."""
        if not img.data:
            return

        try:
            image_stream = io.BytesIO(img.data)

            # Calculate width for cell
            img_width_ratio = img.width / max_cell_width if max_cell_width > 0 else 0.5
            width = min(img_width_ratio * 3.0, 2.8)  # Max 2.8 inches in cell
            width = max(width, 0.5)  # Min 0.5 inches

            para = cell.add_paragraph()
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(width))

            self._image_count += 1

        except Exception as e:
            self._errors.append(f"Failed to add image to cell: {e}")

    def _add_table(self, table_info: TableInfo, page_width: float):
        """Add a table to the document."""
        if not table_info.cells or table_info.num_rows == 0 or table_info.num_cols == 0:
            return

        try:
            # Create the table
            table = self.doc.add_table(rows=table_info.num_rows, cols=table_info.num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

            # Calculate column widths
            table_width = min((table_info.bbox[2] - table_info.bbox[0]) / page_width * 7.0, 6.5)
            col_width = table_width / table_info.num_cols

            # Fill cells
            for cell_info in table_info.cells:
                if cell_info.row < table_info.num_rows and cell_info.col < table_info.num_cols:
                    cell = table.rows[cell_info.row].cells[cell_info.col]
                    cell.width = Inches(col_width)

                    # Set cell text
                    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                    run = para.add_run(cell_info.text)

                    # Apply formatting
                    run.font.size = Pt(cell_info.font_size)
                    run.font.bold = cell_info.is_bold or cell_info.is_header

                    # Alignment
                    if cell_info.alignment == "center":
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    elif cell_info.alignment == "right":
                        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # Style the table with borders
            self._style_table_with_borders(table)

            self._table_count += 1

        except Exception as e:
            self._errors.append(f"Failed to add table: {e}")

    def _style_table_with_borders(self, table):
        """Add borders to a data table."""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')

        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)

        tblPr.append(tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)

    def _remove_table_borders(self, table):
        """Remove all borders from table for seamless column layout."""
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')

        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'nil')
            tblBorders.append(border)

        tblPr.append(tblBorders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)
