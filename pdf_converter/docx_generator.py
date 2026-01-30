"""
DOCX Generator

Creates accurate DOCX from extracted PDF content.
Key features:
- Elements placed in exact reading order
- Proper table rendering with borders
- Images at correct positions
- Multi-column support using Word sections or tables
- Formatting preservation (fonts, sizes, colors)
"""

import io
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass

from docx import Document
from docx.shared import Pt, Inches, RGBColor, Twips, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .pdf_extractor import (
    PageContent, TextBlock, ImageInfo, TableInfo, TableCell,
    PageElement, ElementType, ColumnInfo
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
    
    Creates accurate layout by:
    1. Processing elements in exact reading order
    2. Handling multi-column layouts with tables
    3. Placing images inline at correct positions
    4. Rendering tables with proper structure
    """
    
    # Font size limits
    MIN_FONT_SIZE = 6
    MAX_FONT_SIZE = 48
    
    def __init__(self):
        self.doc: Optional[Document] = None
        self._text_count = 0
        self._image_count = 0
        self._table_count = 0
        self._errors: List[str] = []
    
    def generate(
        self,
        pages: List[PageContent],
        layout_info=None,  # Optional, for backward compatibility
        output_path: str = "output.docx"
    ) -> GenerationResult:
        """
        Generate DOCX from extracted content.
        """
        self.doc = Document()
        self._text_count = 0
        self._image_count = 0
        self._table_count = 0
        self._errors = []
        
        if not pages:
            return GenerationResult(
                success=False,
                output_path=output_path,
                pages_processed=0,
                text_blocks_written=0,
                images_added=0,
                tables_added=0,
                errors=["No pages to process"]
            )
        
        # Setup document based on first page
        self._setup_document(pages[0])
        
        # Process each page
        for i, page in enumerate(pages):
            if i > 0:
                self._add_page_break()
            
            self._generate_page(page)
        
        # Save document
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
                errors=[f"Failed to save: {e}"]
            )
    
    def _setup_document(self, first_page: PageContent):
        """Setup document properties based on PDF page."""
        section = self.doc.sections[0]
        
        # Set page size (convert PDF points to DOCX units)
        section.page_width = Twips(first_page.width * 20)
        section.page_height = Twips(first_page.height * 20)
        
        # Set margins (narrower for better fidelity)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
    
    def _add_page_break(self):
        """Add a page break."""
        self.doc.add_page_break()
    
    def _generate_page(self, page: PageContent):
        """Generate content for a single page."""
        if not page.elements:
            # Fallback: use text_blocks, images, tables directly
            self._generate_page_fallback(page)
            return
        
        # Check if multi-column
        if page.column_info.num_columns > 1:
            self._generate_multicolumn_page(page)
        else:
            self._generate_singlecolumn_page(page)
    
    def _generate_singlecolumn_page(self, page: PageContent):
        """Generate single-column page content."""
        prev_y = 0
        
        for elem in page.elements:
            # Add spacing based on vertical gap
            if prev_y > 0:
                gap = elem.y_position - prev_y
                if gap > 15:
                    self._add_vertical_space(min(gap / 3, 18))
            
            if elem.element_type == ElementType.TEXT:
                self._add_text_block(elem.element, page.width)
            elif elem.element_type == ElementType.IMAGE:
                self._add_image(elem.element, page.width)
            elif elem.element_type == ElementType.TABLE:
                self._add_table(elem.element, page.width)
            
            prev_y = elem.bbox[3]  # Bottom of element
    
    def _generate_multicolumn_page(self, page: PageContent):
        """
        Generate multi-column page content.
        
        Strategy: Group consecutive same-column elements into chunks,
        then render each chunk in a table-based layout.
        """
        # Separate elements by column and track order
        chunks = []
        current_chunk = {'columns': {}, 'type': 'columns', 'y_start': 0}
        
        for elem in page.elements:
            col = elem.column
            
            if col == 0:  # Full-width element
                # Save current chunk if not empty
                if any(current_chunk['columns'].values()):
                    chunks.append(current_chunk)
                
                # Add full-width element as its own chunk
                chunks.append({
                    'type': 'full_width',
                    'element': elem
                })
                
                # Start new column chunk
                current_chunk = {'columns': {}, 'type': 'columns', 'y_start': elem.bbox[3]}
            
            else:
                # Add to current column chunk
                if col not in current_chunk['columns']:
                    current_chunk['columns'][col] = []
                current_chunk['columns'][col].append(elem)
        
        # Don't forget the last chunk
        if any(current_chunk['columns'].values()):
            chunks.append(current_chunk)
        
        # Render each chunk
        for chunk in chunks:
            if chunk['type'] == 'full_width':
                elem = chunk['element']
                if elem.element_type == ElementType.TEXT:
                    self._add_text_block(elem.element, page.width)
                elif elem.element_type == ElementType.IMAGE:
                    self._add_image(elem.element, page.width)
                elif elem.element_type == ElementType.TABLE:
                    self._add_table(elem.element, page.width)
            
            elif chunk['type'] == 'columns':
                self._render_column_chunk(chunk['columns'], page)
    
    def _render_column_chunk(self, columns: Dict[int, List[PageElement]], page: PageContent):
        """Render a chunk of multi-column content using a table."""
        if not columns:
            return
        
        num_cols = page.column_info.num_columns
        
        # Create a table with the appropriate number of columns
        table = self.doc.add_table(rows=1, cols=num_cols)
        table.autofit = False
        
        # Set column widths
        total_width = 7.0  # Total content width in inches
        col_widths = []
        for x_start, x_end in page.column_info.boundaries:
            width_pct = (x_end - x_start) / page.width
            col_widths.append(width_pct * total_width)
        
        # Fill in the columns
        for col_idx in range(num_cols):
            cell = table.rows[0].cells[col_idx]
            
            if col_idx < len(col_widths):
                cell.width = Inches(col_widths[col_idx])
            
            # Set cell properties for clean layout
            self._setup_cell(cell)
            
            # Add content to this column
            col_num = col_idx + 1
            if col_num in columns:
                for elem in columns[col_num]:
                    if elem.element_type == ElementType.TEXT:
                        self._add_text_block_to_cell(cell, elem.element)
                    elif elem.element_type == ElementType.IMAGE:
                        self._add_image_to_cell(cell, elem.element, page.width / num_cols)
        
        # Remove table borders for seamless column layout
        self._remove_table_borders(table)
    
    def _setup_cell(self, cell):
        """Setup cell properties for clean layout."""
        # Remove cell margins/padding
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        # Set cell margins to minimum
        tcMar = OxmlElement('w:tcMar')
        for margin_name in ['top', 'left', 'bottom', 'right']:
            margin = OxmlElement(f'w:{margin_name}')
            margin.set(qn('w:w'), '0')
            margin.set(qn('w:type'), 'dxa')
            tcMar.append(margin)
        tcPr.append(tcMar)
        
        # Vertical alignment
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'top')
        tcPr.append(vAlign)
    
    def _generate_page_fallback(self, page: PageContent):
        """Fallback page generation using text_blocks, images, tables directly."""
        # Combine all elements with Y positions
        all_items = []
        
        for block in page.text_blocks:
            all_items.append(('text', block.y_position, block))
        
        for img in page.images:
            all_items.append(('image', img.y_position, img))
        
        for table in page.tables:
            all_items.append(('table', table.y_position, table))
        
        # Sort by Y position
        all_items.sort(key=lambda x: x[1])
        
        # Add each item
        prev_y = 0
        for item_type, y_pos, item in all_items:
            # Add spacing
            if prev_y > 0 and y_pos - prev_y > 15:
                self._add_vertical_space(min((y_pos - prev_y) / 3, 18))
            
            if item_type == 'text':
                self._add_text_block(item, page.width)
            elif item_type == 'image':
                self._add_image(item, page.width)
            elif item_type == 'table':
                self._add_table(item, page.width)
            
            prev_y = y_pos
    
    def _add_vertical_space(self, points: float):
        """Add vertical space between elements."""
        if points < 3:
            return
        para = self.doc.add_paragraph()
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(points)
        pf = para.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    
    def _add_text_block(self, block: TextBlock, page_width: float):
        """Add a text block as paragraph(s)."""
        # Create paragraph
        para = self.doc.add_paragraph()
        
        # Add text with formatting
        for line in block.lines:
            if line != block.lines[0]:
                # Add line break between lines (not new paragraph)
                para.add_run('\n')
            
            run = para.add_run(line.text)
            
            # Apply formatting
            self._apply_run_formatting(run, line)
        
        # Paragraph formatting
        self._apply_paragraph_formatting(para, block)
        
        self._text_count += 1
    
    def _add_text_block_to_cell(self, cell, block: TextBlock):
        """Add a text block to a table cell."""
        para = cell.add_paragraph()
        
        for line in block.lines:
            if line != block.lines[0]:
                para.add_run('\n')
            
            run = para.add_run(line.text)
            self._apply_run_formatting(run, line)
        
        self._apply_paragraph_formatting(para, block)
        self._text_count += 1
    
    def _apply_run_formatting(self, run, line):
        """Apply formatting to a run based on TextLine properties."""
        # Font size
        size = max(self.MIN_FONT_SIZE, min(self.MAX_FONT_SIZE, line.font_size))
        run.font.size = Pt(size)
        
        # Bold/Italic
        run.font.bold = line.is_bold
        run.font.italic = line.is_italic
        
        # Color (skip black)
        if line.color != (0, 0, 0):
            run.font.color.rgb = RGBColor(*line.color)
        
        # Font name
        font_name = self._clean_font_name(line.font_name)
        if font_name:
            run.font.name = font_name
    
    def _apply_paragraph_formatting(self, para, block: TextBlock):
        """Apply paragraph-level formatting."""
        pf = para.paragraph_format
        
        # Minimal spacing
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        
        # Line spacing
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        # Detect alignment from block position
        # (Could be enhanced based on X positions)
    
    def _clean_font_name(self, font_name: str) -> str:
        """Clean PDF font name for Word."""
        if not font_name:
            return ""
        
        # Remove subset prefix (e.g., "ABCDEF+")
        if "+" in font_name:
            font_name = font_name.split("+", 1)[-1]
        
        # Common mappings
        mappings = {
            "ArialMT": "Arial",
            "Arial-BoldMT": "Arial",
            "Arial-ItalicMT": "Arial",
            "Arial-BoldItalicMT": "Arial",
            "TimesNewRomanPSMT": "Times New Roman",
            "TimesNewRomanPS-BoldMT": "Times New Roman",
            "TimesNewRomanPS-ItalicMT": "Times New Roman",
            "CourierNewPSMT": "Courier New",
            "Helvetica": "Arial",
            "Helvetica-Bold": "Arial",
        }
        
        if font_name in mappings:
            return mappings[font_name]
        
        # Remove style suffixes
        for suffix in ["-Bold", "-Italic", "-BoldItalic", "MT", "PS"]:
            if font_name.endswith(suffix):
                font_name = font_name[:-len(suffix)]
        
        return font_name
    
    def _add_image(self, img: ImageInfo, page_width: float):
        """Add an image to the document."""
        if not img.data:
            return
        
        try:
            image_stream = io.BytesIO(img.data)
            
            # Calculate size - preserve aspect ratio, fit within page
            img_width_ratio = img.width / page_width
            max_width_inches = min(img_width_ratio * 7.5, 7.0)  # Max 7 inches
            max_width_inches = max(max_width_inches, 1.0)  # Min 1 inch
            
            # Create paragraph and add image
            para = self.doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(max_width_inches))
            
            # Minimal spacing
            para.paragraph_format.space_before = Pt(6)
            para.paragraph_format.space_after = Pt(6)
            
            self._image_count += 1
            
        except Exception as e:
            self._errors.append(f"Image error: {e}")
    
    def _add_image_to_cell(self, cell, img: ImageInfo, max_width: float):
        """Add an image to a table cell."""
        if not img.data:
            return
        
        try:
            image_stream = io.BytesIO(img.data)
            
            # Calculate size for cell
            img_width_ratio = img.width / max_width if max_width > 0 else 0.5
            width_inches = min(img_width_ratio * 3.0, 2.8)
            width_inches = max(width_inches, 0.5)
            
            para = cell.add_paragraph()
            run = para.add_run()
            run.add_picture(image_stream, width=Inches(width_inches))
            
            self._image_count += 1
            
        except Exception as e:
            self._errors.append(f"Cell image error: {e}")
    
    def _add_table(self, table_info: TableInfo, page_width: float):
        """Add a proper table to the document."""
        if not table_info.cells or table_info.num_rows == 0:
            return
        
        try:
            # Create table
            table = self.doc.add_table(rows=table_info.num_rows, cols=table_info.num_cols)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Calculate width
            table_width_ratio = (table_info.bbox[2] - table_info.bbox[0]) / page_width
            table_width_inches = min(table_width_ratio * 7.5, 7.0)
            col_width = table_width_inches / table_info.num_cols
            
            # Fill cells
            for row_idx, row_cells in enumerate(table_info.cells):
                for col_idx, cell_info in enumerate(row_cells):
                    if row_idx < len(table.rows) and col_idx < len(table.rows[row_idx].cells):
                        cell = table.rows[row_idx].cells[col_idx]
                        cell.width = Inches(col_width)
                        
                        # Set cell text
                        para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                        
                        # Clear any default text
                        if para.runs:
                            para.clear()
                        
                        run = para.add_run(cell_info.text)
                        
                        # Formatting
                        run.font.size = Pt(cell_info.font_size)
                        run.font.bold = cell_info.is_bold
                        
                        # Cell vertical alignment
                        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            
            # Apply table borders
            self._style_table_borders(table)
            
            # Spacing around table
            # Add empty paragraph before if needed
            
            self._table_count += 1
            
        except Exception as e:
            self._errors.append(f"Table error: {e}")
    
    def _style_table_borders(self, table):
        """Add borders to table."""
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
        """Remove borders from table (for column layout)."""
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
