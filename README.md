# PDF to DOCX Converter v2.0

A high-fidelity PDF to DOCX converter with **multi-column layout preservation**, **structure analysis**, and **content validation**. Runs locally via command line with zero content loss.

## Key Features

- **Multi-Column Layout Preservation** - Accurately detects and recreates 2-column, 3-column, and complex layouts
- **Structure Analysis** - Scans PDF structure page-by-page to understand layout before conversion
- **Content Validation** - Compares output with original to ensure no content loss
- **Post-Processing Cleanup** - Removes blank pages, fixes spacing issues automatically
- **Text Formatting** - Preserves font family, size, bold, italic, color, underline
- **Image Handling** - Extracts and positions images accurately
- **Table Detection** - Recreates table structures with proper formatting

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/PDFtoDOCX.git
cd PDFtoDOCX

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Quick Start

### Basic Conversion (Simple PDFs)
```bash
python convert.py document.pdf
```

### Advanced Conversion (Multi-Column PDFs) - RECOMMENDED
```bash
python convert.py document.pdf --advanced
```

### With Validation and Cleanup
```bash
python convert.py document.pdf --advanced --validate --cleanup
```

## Usage Modes

### 1. Standard Mode
Fast conversion using pdf2docx library. Best for simple, single-column documents.

```bash
python convert.py document.pdf
python convert.py document.pdf -o output.docx
```

### 2. Advanced Mode (Recommended for Complex PDFs)
Uses custom structure analyzer for multi-column layout preservation.

```bash
python convert.py document.pdf --advanced
python convert.py document.pdf --advanced --quality high
```

Quality levels:
- `draft` - Fast conversion, basic formatting
- `standard` - Balanced quality and speed
- `high` - Maximum fidelity (default)

### 3. Analyze Mode
Analyze PDF structure without converting. Useful for understanding document layout.

```bash
python convert.py document.pdf --analyze
python convert.py document.pdf --analyze --detailed
```

Example output:
```
Analyzing: document.pdf
------------------------------------------------------------
Total Pages: 10
Dominant Layout: two_column
Consistent Layout: No (mixed layouts)

Layout Distribution:
  two_column: 7 pages (70%)
  single_column: 2 pages (20%)
  three_column: 1 pages (10%)

Summary:
  Total Text Blocks: 245
  Total Images: 12
  Multi-Column Pages: 8

Recommendations:
  - Use --advanced mode for best multi-column preservation
  - Images will be extracted and positioned in output
  - Mixed layouts detected - consider page-by-page review
```

### 4. Validation Mode
Validate conversion by comparing content with original PDF.

```bash
# Convert and validate
python convert.py document.pdf --validate

# Compare existing files
python convert.py document.pdf --compare existing.docx
```

### 5. Cleanup Mode
Remove blank pages and fix formatting issues.

```bash
# Convert with cleanup
python convert.py document.pdf --cleanup

# Cleanup existing DOCX
python convert.py --cleanup-only existing.docx
```

## Command Line Options

| Option | Description |
|--------|-------------|
| `--advanced` | Use advanced converter with layout preservation |
| `--analyze` | Analyze PDF structure only (no conversion) |
| `--detailed` | Show detailed per-page analysis |
| `--validate` | Validate conversion against original |
| `--cleanup` | Run post-processing cleanup |
| `--compare DOCX` | Compare PDF with existing DOCX |
| `--cleanup-only DOCX` | Cleanup existing DOCX (no conversion) |
| `--quality LEVEL` | Conversion quality (draft/standard/high) |
| `-o, --output PATH` | Output file path |
| `-p, --pages PAGES` | Specific pages (e.g., "0,1,2" or "0-5") |
| `-q, --quiet` | Suppress progress output |

## Python API

### Standard Conversion
```python
from pdf_converter import PDFConverter

converter = PDFConverter()
result = converter.convert("input.pdf", "output.docx")

if result.success:
    print(f"Converted {result.pages_converted} pages")
```

### Advanced Conversion with Layout Preservation
```python
from pdf_converter import AdvancedPDFConverter
from pdf_converter.advanced_converter import ConversionQuality

converter = AdvancedPDFConverter(quality=ConversionQuality.HIGH)
output_path, stats, validation = converter.convert(
    "input.pdf",
    "output.docx",
    validate=True
)

print(f"Text blocks converted: {stats.text_blocks_converted}")
print(f"Columns preserved: {stats.columns_preserved}")
print(f"Validation passed: {validation.is_valid}")
```

### PDF Structure Analysis
```python
from pdf_converter import PDFAnalyzer

with PDFAnalyzer("document.pdf") as analyzer:
    structure = analyzer.analyze()

    print(f"Pages: {structure.page_count}")
    print(f"Layout: {structure.dominant_layout.value}")

    for page in structure.pages:
        print(f"Page {page.page_num + 1}: {page.layout_type.value}")
        print(f"  Columns: {len(page.columns)}")
        print(f"  Text blocks: {len(page.text_blocks)}")
```

### Post-Processing
```python
from pdf_converter import PostProcessor, ConversionValidator

# Cleanup document
processor = PostProcessor()
result = processor.cleanup_document("output.docx")
print(f"Blank paragraphs removed: {result.blank_paragraphs_removed}")

# Validate conversion
validator = ConversionValidator()
report = validator.full_validation("input.pdf", "output.docx")
print(f"Content match: {report['content_match']:.1f}%")
```

## How It Works

### 1. Structure Analysis
The converter first analyzes the PDF structure:
- Detects column layouts (single, double, triple, mixed)
- Identifies text blocks and their positions
- Maps reading order for multi-column content
- Locates headers, footers, and images

### 2. Layout-Aware Conversion
Based on the analysis:
- Single-column pages are converted directly
- Multi-column pages use table-based layout preservation
- Images are extracted and positioned
- Text formatting is mapped to DOCX equivalents

### 3. Post-Processing Validation
After conversion:
- Compares word content between PDF and DOCX
- Identifies missing or misplaced content
- Removes unnecessary blank elements
- Reports issues and suggestions

## Handling Complex Documents

### Multi-Column Layouts
For documents with 2 or 3 columns (like academic papers, newsletters):
```bash
python convert.py paper.pdf --advanced --validate
```

### Mixed Layouts
For documents with varying layouts per page:
```bash
python convert.py document.pdf --analyze --detailed  # First, understand structure
python convert.py document.pdf --advanced --cleanup
```

### Large Documents
For documents with many pages:
```bash
# Convert in batches
python convert.py large.pdf --pages 0-50 -o part1.docx
python convert.py large.pdf --pages 50-100 -o part2.docx
```

## Troubleshooting

### "Content match below 85%"
- Check if PDF contains scanned images (requires OCR)
- Try `--advanced` mode for complex layouts
- Review specific pages with `--analyze --detailed`

### Blank pages in output
```bash
python convert.py document.pdf --cleanup
# Or fix existing file
python convert.py --cleanup-only output.docx
```

### Columns not preserved
```bash
python convert.py document.pdf --advanced --quality high
```

### Missing images
- Ensure PDF contains extractable images (not scanned)
- Check for image extraction errors in output

## Limitations

- **Scanned PDFs**: Image-based PDFs require OCR (not included)
- **Complex Graphics**: Vector graphics may not convert perfectly
- **Custom Fonts**: Substituted with similar system fonts if unavailable
- **Form Fields**: Converted to static content
- **Very Complex Layouts**: May need manual adjustment

## Project Structure

```
PDFtoDOCX/
├── convert.py                 # CLI entry point
├── pdf_converter/
│   ├── __init__.py           # Package exports
│   ├── converter.py          # Standard converter (pdf2docx)
│   ├── analyzer.py           # PDF structure analyzer
│   ├── advanced_converter.py # Advanced layout-preserving converter
│   └── postprocessor.py      # Validation and cleanup
├── requirements.txt          # Dependencies
├── test_installation.py      # Installation verification
└── README.md                 # Documentation
```

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## License

MIT License - See LICENSE file for details.
