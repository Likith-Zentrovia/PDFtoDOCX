"""
PDF to DOCX Converter

Main converter that orchestrates:
1. Layout analysis (Claude Vision - quick, layout only)
2. Content extraction (Python/PyMuPDF - all text/images)
3. DOCX generation (python-docx - accurate output)
"""

import os
from pathlib import Path
from typing import Optional, List
from dataclasses import dataclass

from .layout_analyzer import LayoutAnalyzer, DocumentLayout
from .pdf_extractor import PDFExtractor, PageContent
from .docx_generator import DOCXGenerator, GenerationResult


@dataclass
class ConversionResult:
    """Final conversion result."""
    success: bool
    output_path: str
    pages: int
    text_blocks: int
    images: int
    tables: int
    layout_type: str
    errors: List[str]


class PDFtoDOCXConverter:
    """
    Main PDF to DOCX converter.

    Workflow:
    1. Analyze layout with Claude Vision (samples 3 pages for speed)
    2. Extract all content with PyMuPDF (fast, accurate)
    3. Generate DOCX using layout info (preserves structure)
    """

    def __init__(self, api_key: Optional[str] = None):
        """
        Initialize converter.

        Args:
            api_key: Anthropic API key for layout analysis
        """
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        self.layout_analyzer = None
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
                layout_type="unknown",
                errors=[f"File not found: {pdf_path}"]
            )

        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix('.docx'))

        errors = []

        print(f"\n{'='*60}")
        print(f"PDF TO DOCX CONVERTER")
        print(f"{'='*60}")
        print(f"Input:  {pdf_path}")
        print(f"Output: {output_path}")

        # Step 1: Layout Analysis (Vision - quick)
        print(f"\n[1/3] Analyzing layout...")

        layout = self._analyze_layout(pdf_path)
        if layout:
            print(f"       Detected: {layout.dominant_columns}-column layout")
            print(f"       Pages: {layout.page_count}")
        else:
            print(f"       Using default single-column layout")

        # Step 2: Content Extraction (Python - fast)
        print(f"\n[2/3] Extracting content...")

        pages = self.extractor.extract_document(
            pdf_path,
            layout.pages if layout else []
        )

        total_blocks = sum(len(p.text_blocks) for p in pages)
        total_images = sum(len(p.images) for p in pages)
        total_tables = sum(len(p.tables) for p in pages)
        total_elements = sum(len(p.elements) for p in pages)
        print(f"       Text blocks: {total_blocks}")
        print(f"       Images: {total_images}")
        print(f"       Tables: {total_tables}")
        print(f"       Total elements (ordered): {total_elements}")

        # Step 3: DOCX Generation
        print(f"\n[3/3] Generating DOCX...")

        result = self.generator.generate(pages, layout, output_path)

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
            layout_type=f"{layout.dominant_columns}-column" if layout else "single-column",
            errors=errors
        )

    def _analyze_layout(self, pdf_path: str) -> Optional[DocumentLayout]:
        """Analyze PDF layout using Vision (if API key available)."""
        if not self.api_key:
            print(f"       (No API key - using default layout)")
            return None

        try:
            self.layout_analyzer = LayoutAnalyzer(self.api_key)
            return self.layout_analyzer.analyze_document(pdf_path, sample_pages=3)
        except Exception as e:
            print(f"       Warning: Layout analysis failed ({e})")
            print(f"       Using default single-column layout")
            return None


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
        api_key: Anthropic API key for layout analysis (optional)

    Returns:
        ConversionResult
    """
    converter = PDFtoDOCXConverter(api_key)
    return converter.convert(pdf_path, output_path)
