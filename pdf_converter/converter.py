"""
Core PDF to DOCX Converter Module

This module provides high-fidelity PDF to DOCX conversion with preservation of:
- Text formatting (font family, size, weight, style, color)
- Images with accurate positioning
- Tables with proper structure and styling
- Page layout and margins
"""

import os
import sys
from pathlib import Path
from typing import Optional, List, Tuple, Callable
from dataclasses import dataclass

from pdf2docx import Converter
import fitz  # PyMuPDF


@dataclass
class ConversionResult:
    """Result of a PDF to DOCX conversion."""
    success: bool
    input_path: str
    output_path: str
    pages_converted: int
    error_message: Optional[str] = None
    warnings: Optional[List[str]] = None


class PDFConverter:
    """
    High-fidelity PDF to DOCX converter.

    This converter uses pdf2docx library which provides excellent preservation of:
    - Text: font family, size, bold, italic, color, underline
    - Images: embedded images with positioning
    - Tables: structure, cell content, borders
    - Layout: margins, columns, spacing

    Usage:
        converter = PDFConverter()
        result = converter.convert("input.pdf", "output.docx")

        # Or convert specific pages
        result = converter.convert("input.pdf", "output.docx", pages=[0, 1, 2])

        # Batch conversion
        results = converter.batch_convert(["file1.pdf", "file2.pdf"])
    """

    def __init__(self, verbose: bool = True):
        """
        Initialize the converter.

        Args:
            verbose: If True, print progress information during conversion.
        """
        self.verbose = verbose
        self._progress_callback: Optional[Callable[[int, int], None]] = None

    def set_progress_callback(self, callback: Callable[[int, int], None]) -> None:
        """
        Set a callback function for progress updates.

        Args:
            callback: Function that takes (current_page, total_pages) as arguments.
        """
        self._progress_callback = callback

    def _log(self, message: str) -> None:
        """Print a message if verbose mode is enabled."""
        if self.verbose:
            print(message)

    def get_pdf_info(self, pdf_path: str) -> dict:
        """
        Get information about a PDF file.

        Args:
            pdf_path: Path to the PDF file.

        Returns:
            Dictionary containing PDF metadata and statistics.
        """
        pdf_path = str(Path(pdf_path).resolve())

        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        doc = fitz.open(pdf_path)

        info = {
            "path": pdf_path,
            "page_count": len(doc),
            "metadata": doc.metadata,
            "pages": []
        }

        for page_num, page in enumerate(doc):
            page_info = {
                "number": page_num + 1,
                "width": page.rect.width,
                "height": page.rect.height,
                "rotation": page.rotation,
                "has_images": len(page.get_images()) > 0,
                "image_count": len(page.get_images()),
                "has_text": len(page.get_text().strip()) > 0
            }
            info["pages"].append(page_info)

        doc.close()
        return info

    def convert(
        self,
        pdf_path: str,
        docx_path: Optional[str] = None,
        pages: Optional[List[int]] = None,
        start_page: int = 0,
        end_page: Optional[int] = None
    ) -> ConversionResult:
        """
        Convert a PDF file to DOCX format.

        Args:
            pdf_path: Path to the input PDF file.
            docx_path: Path for the output DOCX file. If not provided,
                      uses the same name as the PDF with .docx extension.
            pages: List of specific page numbers to convert (0-indexed).
                  If provided, start_page and end_page are ignored.
            start_page: Starting page number (0-indexed). Default is 0.
            end_page: Ending page number (0-indexed, exclusive).
                     Default is None (convert to the last page).

        Returns:
            ConversionResult object with conversion status and details.
        """
        pdf_path = str(Path(pdf_path).resolve())
        warnings = []

        # Validate input file
        if not os.path.exists(pdf_path):
            return ConversionResult(
                success=False,
                input_path=pdf_path,
                output_path="",
                pages_converted=0,
                error_message=f"PDF file not found: {pdf_path}"
            )

        if not pdf_path.lower().endswith('.pdf'):
            warnings.append("Input file does not have .pdf extension")

        # Set output path
        if docx_path is None:
            docx_path = str(Path(pdf_path).with_suffix('.docx'))
        else:
            docx_path = str(Path(docx_path).resolve())

        # Ensure output directory exists
        output_dir = os.path.dirname(docx_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        self._log(f"Converting: {pdf_path}")
        self._log(f"Output: {docx_path}")

        try:
            # Create converter instance
            cv = Converter(pdf_path)

            # Get total pages for progress tracking
            total_pages = len(cv.pages)
            self._log(f"Total pages in PDF: {total_pages}")

            # Determine pages to convert
            if pages is not None:
                # Convert specific pages
                self._log(f"Converting specific pages: {pages}")
                for i, page_num in enumerate(pages):
                    cv.convert(docx_path, pages=[page_num])
                    if self._progress_callback:
                        self._progress_callback(i + 1, len(pages))
                pages_converted = len(pages)
            else:
                # Convert page range
                if end_page is None:
                    end_page = total_pages

                self._log(f"Converting pages {start_page} to {end_page - 1}")

                # Convert with progress tracking
                cv.convert(docx_path, start=start_page, end=end_page)
                pages_converted = end_page - start_page

            cv.close()

            self._log(f"Conversion complete! {pages_converted} pages converted.")

            return ConversionResult(
                success=True,
                input_path=pdf_path,
                output_path=docx_path,
                pages_converted=pages_converted,
                warnings=warnings if warnings else None
            )

        except Exception as e:
            error_msg = str(e)
            self._log(f"Error during conversion: {error_msg}")

            return ConversionResult(
                success=False,
                input_path=pdf_path,
                output_path=docx_path,
                pages_converted=0,
                error_message=error_msg,
                warnings=warnings if warnings else None
            )

    def batch_convert(
        self,
        pdf_paths: List[str],
        output_dir: Optional[str] = None,
        continue_on_error: bool = True
    ) -> List[ConversionResult]:
        """
        Convert multiple PDF files to DOCX format.

        Args:
            pdf_paths: List of paths to PDF files.
            output_dir: Directory for output files. If not provided,
                       each DOCX is saved alongside its source PDF.
            continue_on_error: If True, continue converting remaining files
                              even if one fails. Default is True.

        Returns:
            List of ConversionResult objects, one for each input file.
        """
        results = []
        total_files = len(pdf_paths)

        self._log(f"Batch conversion: {total_files} files")

        for i, pdf_path in enumerate(pdf_paths):
            self._log(f"\nProcessing file {i + 1}/{total_files}")

            # Determine output path
            if output_dir:
                output_dir_path = Path(output_dir)
                output_dir_path.mkdir(parents=True, exist_ok=True)
                docx_path = str(output_dir_path / Path(pdf_path).with_suffix('.docx').name)
            else:
                docx_path = None

            result = self.convert(pdf_path, docx_path)
            results.append(result)

            if not result.success and not continue_on_error:
                self._log("Stopping batch conversion due to error")
                break

        # Summary
        successful = sum(1 for r in results if r.success)
        self._log(f"\nBatch conversion complete: {successful}/{len(results)} successful")

        return results

    def convert_directory(
        self,
        input_dir: str,
        output_dir: Optional[str] = None,
        recursive: bool = False,
        continue_on_error: bool = True
    ) -> List[ConversionResult]:
        """
        Convert all PDF files in a directory.

        Args:
            input_dir: Path to directory containing PDF files.
            output_dir: Directory for output files. If not provided,
                       DOCX files are saved alongside source PDFs.
            recursive: If True, also process subdirectories.
            continue_on_error: If True, continue even if some files fail.

        Returns:
            List of ConversionResult objects.
        """
        input_path = Path(input_dir)

        if not input_path.exists():
            raise FileNotFoundError(f"Directory not found: {input_dir}")

        if not input_path.is_dir():
            raise NotADirectoryError(f"Not a directory: {input_dir}")

        # Find all PDF files
        if recursive:
            pdf_files = list(input_path.rglob("*.pdf"))
        else:
            pdf_files = list(input_path.glob("*.pdf"))

        pdf_paths = [str(f) for f in pdf_files]

        self._log(f"Found {len(pdf_paths)} PDF files in {input_dir}")

        if not pdf_paths:
            return []

        return self.batch_convert(pdf_paths, output_dir, continue_on_error)


def convert_pdf_to_docx(
    pdf_path: str,
    docx_path: Optional[str] = None,
    pages: Optional[List[int]] = None,
    verbose: bool = True
) -> ConversionResult:
    """
    Convenience function for quick PDF to DOCX conversion.

    Args:
        pdf_path: Path to the input PDF file.
        docx_path: Path for the output DOCX file (optional).
        pages: Specific pages to convert (optional, 0-indexed).
        verbose: Print progress information.

    Returns:
        ConversionResult object.

    Example:
        result = convert_pdf_to_docx("document.pdf")
        if result.success:
            print(f"Converted to: {result.output_path}")
    """
    converter = PDFConverter(verbose=verbose)
    return converter.convert(pdf_path, docx_path, pages)
