"""
DOCX Generator

Creates accurate DOCX from extracted PDF content using layout information.
Preserves:
- Multi-column layouts (using tables)
- Text formatting (font size, bold, italic, color)
- Images with positioning
- Reading order
"""

import io
from typing import List, Dict, Optional
from dataclasses import dataclass

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .layout_analyzer import LayoutInfo, DocumentLayout
from .pdf_extractor import PageContent, TextBlock, ImageInfo


@dataclass
class GenerationResult:
    """Result of DOCX generation."""
    success: bool
    output_path: str
    pages_processed: int
    text_blocks_written: int
    images_added: int
    errors: List[str]


class DOCXGenerator:
    """
    Generates DOCX from extracted PDF content.

    Uses layout information to create accurate multi-column layouts.
    """

    # Map PDF font sizes to reasonable DOCX sizes
    MIN_FONT_SIZE = 8
    MAX_FONT_SIZE = 28

    def __init__(self):
        self.doc: Optional[Document] = None
        self._text_count = 0
        self._image_count = 0
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
        self._errors = []

        # Setup document
        self._setup_document(pages[0] if pages else None)

        # Process each page
        for i, page in enumerate(pages):
            page_layout = layout.pages[i] if i < len(layout.pages) else None

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
                errors=self._errors
            )
        except Exception as e:
            return GenerationResult(
                success=False,
                output_path=output_path,
                pages_processed=len(pages),
                text_blocks_written=self._text_count,
                images_added=self._image_count,
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

        # Set margins
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)

    def _generate_page(self, page: PageContent, layout: Optional[LayoutInfo]):
        """Generate content for a single page."""
        num_columns = layout.num_columns if layout else 1

        if num_columns == 1:
            self._generate_single_column(page)
        else:
            self._generate_multi_column(page, layout)

    def _generate_single_column(self, page: PageContent):
        """Generate single-column page content."""
        # Add blocks in reading order
        for idx in page.reading_order:
            if idx < len(page.text_blocks):
                block = page.text_blocks[idx]
                self._add_text_block(block)

        # Add images
        for img in page.images:
            self._add_image(img, page.width)

    def _generate_multi_column(self, page: PageContent, layout: LayoutInfo):
        """Generate multi-column page content using table layout."""
        num_cols = layout.num_columns

        # Separate content by type
        headers = [page.text_blocks[i] for i in page.reading_order
                   if i < len(page.text_blocks) and page.text_blocks[i].is_header]
        footers = [page.text_blocks[i] for i in page.reading_order
                   if i < len(page.text_blocks) and page.text_blocks[i].is_footer]
        content_indices = [i for i in page.reading_order
                          if i < len(page.text_blocks)
                          and not page.text_blocks[i].is_header
                          and not page.text_blocks[i].is_footer]

        # Add headers first
        for block in headers:
            self._add_text_block(block)

        # Group content by column
        columns: Dict[int, List[TextBlock]] = {i: [] for i in range(1, num_cols + 1)}

        for idx in content_indices:
            block = page.text_blocks[idx]
            col = max(1, min(block.column, num_cols))
            columns[col].append(block)

        # Create table for column layout
        table = self.doc.add_table(rows=1, cols=num_cols)
        table.autofit = False

        # Calculate column widths
        total_width = 6.0  # Content width in inches
        col_widths = self._calculate_column_widths(layout.column_boundaries, total_width, num_cols)

        # Fill columns
        for col_idx in range(num_cols):
            cell = table.rows[0].cells[col_idx]
            cell.width = Inches(col_widths[col_idx])

            for block in columns.get(col_idx + 1, []):
                self._add_text_block_to_cell(cell, block)

        # Remove table borders
        self._remove_table_borders(table)

        # Add footers
        for block in footers:
            self._add_text_block(block)

        # Add images
        for img in page.images:
            self._add_image(img, page.width)

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
            # Fill with equal widths
            default_width = total_width / num_cols
            while len(widths) < num_cols:
                widths.append(default_width)
        elif len(widths) > num_cols:
            widths = widths[:num_cols]

        # Fallback if empty
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

        # Paragraph spacing
        para.paragraph_format.space_after = Pt(6)
        para.paragraph_format.space_before = Pt(0)

    def _clean_font_name(self, font_name: str) -> str:
        """Clean up font name from PDF."""
        if not font_name:
            return ""

        # Remove common prefixes (subset font names)
        name = font_name
        for prefix in ["AAAAAA+", "BAAAAA+", "CAAAAA+", "DAAAAA+"]:
            if name.startswith(prefix):
                name = name[7:]
                break

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
        """Add an image to the document."""
        if not img.data:
            return

        try:
            image_stream = io.BytesIO(img.data)

            # Calculate appropriate width
            img_width_ratio = img.width / page_width
            max_width = min(img_width_ratio * 6, 5.5)  # Max 5.5 inches

            para = self.doc.add_paragraph()
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(max_width))

            self._image_count += 1

        except Exception as e:
            self._errors.append(f"Failed to add image: {e}")

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
