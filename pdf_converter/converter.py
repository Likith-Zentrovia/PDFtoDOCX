"""
PDF to DOCX Converter

Main converter that orchestrates:
1. Layout hints (optional, from Claude Vision - only 1 API call)
2. Content extraction (Python/PyMuPDF - all text/images/tables)
3. DOCX generation (python-docx - accurate output)

The conversion works WITHOUT AI - Vision is only used for optional hints.
"""

import os
from pathlib import Path
from typing import Optional, List
from dataclasses import dataclass

from .pdf_extractor import PDFExtractor
from .docx_generator import DOCXGenerator, GenerationResult
from .layout_analyzer import get_layout_hints, LayoutHints


@dataclass
class ConversionResult:
    """Final conversion result."""
    success: bool
    output_path: str
    pages: int
    text_blocks: int
    images: int
    tables: int
    columns_detected: int
    layout_type: str
    errors: List[str]


class PDFtoDOCXConverter:
    """
    Main PDF to DOCX converter.
    
    Workflow:
    1. (Optional) Get layout hints from Vision if API key provided
    2. Extract all content with PyMuPDF (fast, accurate, handles layout)
    3. Generate DOCX using extracted content
    
    NO AI REQUIRED for conversion to work!
    """
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize converter.
        
        Args:
            api_key: Optional Anthropic API key for layout hints (NOT required)
        """
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        self.extractor = PDFExtractor()
        self.generator = DOCXGenerator()
    
    def convert(
        self,
        pdf_path: str,
        output_path: Optional[str] = None
    ) -> ConversionResult:
        """
        Convert PDF to DOCX.
        
        Args:
            pdf_path: Path to input PDF
            output_path: Path for output DOCX (auto-generated if not provided)
        
        Returns:
            ConversionResult with details
        """
        pdf_path = str(Path(pdf_path).resolve())
        
        if not os.path.exists(pdf_path):
            return ConversionResult(
                success=False,
                output_path="",
                pages=0,
                text_blocks=0,
                images=0,
                tables=0,
                columns_detected=1,
                layout_type="unknown",
                errors=[f"File not found: {pdf_path}"]
            )
        
        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix('.docx'))
        
        errors = []
        
        print(f"\n{'='*60}")
        print(f"PDF TO DOCX CONVERTER v4.0")
        print(f"{'='*60}")
        print(f"Input:  {pdf_path}")
        print(f"Output: {output_path}")
        
        # Step 1: Optional layout hints (only if API key provided)
        layout_hints = None
        if self.api_key:
            print(f"\n[1/3] Getting layout hints (optional)...")
            layout_hints = get_layout_hints(pdf_path, self.api_key)
            if layout_hints:
                print(f"       Hint: {layout_hints.num_columns}-column layout")
            else:
                print(f"       No hints available, using auto-detection")
        else:
            print(f"\n[1/3] Layout hints: Skipped (no API key)")
            print(f"       Using automatic column detection")
        
        # Step 2: Content Extraction (Python - no AI)
        print(f"\n[2/3] Extracting content...")
        
        try:
            pages = self.extractor.extract_document(pdf_path, [])
        except Exception as e:
            return ConversionResult(
                success=False,
                output_path=output_path,
                pages=0,
                text_blocks=0,
                images=0,
                tables=0,
                columns_detected=1,
                layout_type="unknown",
                errors=[f"Extraction failed: {e}"]
            )
        
        # Calculate statistics
        total_blocks = sum(len(p.text_blocks) for p in pages)
        total_images = sum(len(p.images) for p in pages)
        total_tables = sum(len(p.tables) for p in pages)
        total_elements = sum(len(p.elements) for p in pages)
        
        # Determine dominant column count
        column_counts = [p.column_info.num_columns for p in pages]
        dominant_columns = max(set(column_counts), key=column_counts.count) if column_counts else 1
        
        print(f"       Pages: {len(pages)}")
        print(f"       Text blocks: {total_blocks}")
        print(f"       Images: {total_images}")
        print(f"       Tables: {total_tables}")
        print(f"       Columns detected: {dominant_columns}")
        print(f"       Total elements: {total_elements}")
        
        # Step 3: DOCX Generation
        print(f"\n[3/3] Generating DOCX...")
        
        result = self.generator.generate(pages, None, output_path)
        
        if result.errors:
            errors.extend(result.errors)
        
        # Summary
        print(f"\n{'='*60}")
        if result.success:
            print(f"SUCCESS!")
            print(f"{'='*60}")
            print(f"Output: {result.output_path}")
            print(f"Pages:  {result.pages_processed}")
            print(f"Blocks: {result.text_blocks_written}")
            print(f"Images: {result.images_added}")
            print(f"Tables: {result.tables_added}")
        else:
            print(f"FAILED")
            print(f"{'='*60}")
            for err in errors:
                print(f"  Error: {err}")
        
        return ConversionResult(
            success=result.success,
            output_path=result.output_path,
            pages=result.pages_processed,
            text_blocks=result.text_blocks_written,
            images=result.images_added,
            tables=result.tables_added,
            columns_detected=dominant_columns,
            layout_type=f"{dominant_columns}-column",
            errors=errors
        )


def convert(
    pdf_path: str,
    output_path: Optional[str] = None,
    api_key: Optional[str] = None
) -> ConversionResult:
    """
    Convert PDF to DOCX - main entry point.
    
    Args:
        pdf_path: Path to PDF file
        output_path: Path for DOCX output (optional)
        api_key: Anthropic API key for optional layout hints (optional)
    
    Returns:
        ConversionResult
    """
    converter = PDFtoDOCXConverter(api_key)
    return converter.convert(pdf_path, output_path)
