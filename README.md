# PDF to DOCX Converter

Convert PDF to DOCX with accurate layout preservation - creating a true replica of the original document.

## How It Works

```
PDF → [Vision: Layout Analysis] → [Python: Content Extraction] → DOCX
```

1. **Claude Vision** analyzes page structure (columns, headers, tables, images, element order)
2. **PyMuPDF** extracts all text, images, and tables with formatting and positions
3. **python-docx** generates DOCX with elements in exact Y-order for proper layout

Vision is used to understand layout complexity and element relationships. All actual content extraction is done by Python.

## Key Features

- **Accurate Element Positioning**: Text, images, and tables are interleaved by Y-position (not images at end)
- **Table Detection & Extraction**: Tables are detected and rendered with proper structure
- **Multi-Column Layout**: Columns are preserved with proper element flow
- **Spacing Preservation**: Vertical gaps between elements are maintained
- **Font & Formatting**: Size, bold, italic, color preserved
- **Reading Order**: Correct flow across columns and pages

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Basic conversion
python convert.py document.pdf

# Specify output
python convert.py document.pdf -o output.docx

# With API key for better layout detection
export ANTHROPIC_API_KEY='your-key'
python convert.py document.pdf
```

## With vs Without API Key

| Feature | Without API Key | With API Key |
|---------|-----------------|--------------|
| Single-column PDFs | Works | Works |
| Multi-column PDFs | Basic (may miss columns) | Accurate column detection |
| Tables | Auto-detected by PyMuPDF | Enhanced with Vision hints |
| Images | Positioned by Y-order | Better placement context |
| Speed | Fast | Slightly slower (3 API calls) |

The API key enables Claude Vision to analyze layout structure including:
- Number of columns and their boundaries
- Table regions with row/column structure
- Image positions and their relationship to text
- Element order for accurate reconstruction

## What Gets Preserved

- **Text** - All text content with formatting
- **Fonts** - Size, bold, italic, color
- **Columns** - Multi-column layouts
- **Images** - Extracted and positioned in correct location (not at end of page)
- **Tables** - Detected and rendered with proper structure
- **Reading Order** - Correct flow across columns
- **Spacing** - Vertical gaps between elements

## Example Output

```
============================================================
PDF TO DOCX CONVERTER
============================================================
Input:  document.pdf
Output: document.docx

[1/3] Analyzing layout...
       Detected: 2-column layout
       Pages: 10

[2/3] Extracting content...
       Text blocks: 245
       Images: 12
       Tables: 3
       Total elements (ordered): 260

[3/3] Generating DOCX...

============================================================
SUCCESS!
============================================================
Output: document.docx
Pages:  10
Blocks: 245
Images: 12
Tables: 3
```

## Python API

```python
from pdf_converter import convert

result = convert("document.pdf")

if result.success:
    print(f"Output: {result.output_path}")
    print(f"Layout: {result.layout_type}")
    print(f"Tables: {result.tables}")
    print(f"Images: {result.images}")
```

## Architecture

The converter uses a unified element list approach:

1. **Layout Analysis** (Claude Vision) - Detects:
   - Column count and boundaries
   - Header/footer regions  
   - Table bounding boxes
   - Image regions and positions
   - Element ordering hints

2. **Content Extraction** (PyMuPDF) - Extracts:
   - Text blocks with font info
   - Tables using built-in table finder
   - Images with positions
   - Creates unified element list sorted by Y-position

3. **DOCX Generation** (python-docx) - Generates:
   - Elements in exact Y-order (interleaved text, images, tables)
   - Tables with borders and formatting
   - Images at correct positions
   - Multi-column layouts using Word tables

## Requirements

- Python 3.8+
- PyMuPDF
- python-docx
- Pillow
- anthropic (optional, for layout analysis)

## Project Structure

```
PDFtoDOCX/
├── convert.py              # CLI
├── pdf_converter/
│   ├── __init__.py         # Package exports
│   ├── converter.py        # Main converter orchestration
│   ├── layout_analyzer.py  # Vision-based layout analysis
│   ├── pdf_extractor.py    # PyMuPDF content extraction
│   └── docx_generator.py   # DOCX generation with Y-order
└── requirements.txt
```

## Changelog

### v4.0.0
- **Major**: Elements now interleaved by Y-position (images/tables no longer at end)
- Added table detection and extraction using PyMuPDF
- Enhanced Vision prompt to detect tables and image regions
- Created unified PageElement list for accurate layout
- Improved spacing preservation between elements
- Added element_order hints from Vision analysis

### v3.0.0
- Initial multi-column support
- Basic text and image extraction

## License

MIT
