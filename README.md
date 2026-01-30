# PDF to DOCX Converter

Convert PDF to DOCX with accurate layout preservation - **NO AI REQUIRED**.

## Key Features

- **Works WITHOUT API key** - Pure Python extraction and conversion
- **Automatic column detection** - Analyzes text positions to detect multi-column layouts
- **Proper table extraction** - Tables rendered with structure and borders
- **Images at correct positions** - Placed inline at their actual Y-positions
- **Font formatting preserved** - Size, bold, italic, color maintained
- **Reading order preserved** - Correct flow across columns

## How It Works

```
PDF → [PyMuPDF: Extract Everything] → [Auto Column Detection] → DOCX
```

1. **Text Extraction**: Line-by-line with font formatting
2. **Table Detection**: Using PyMuPDF's built-in table finder
3. **Image Extraction**: With exact bounding boxes
4. **Column Detection**: Automatic from text X-positions (no AI)
5. **Reading Order**: Elements sorted by Y-position, respecting columns
6. **DOCX Generation**: Elements placed in exact visual order

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Basic conversion (NO API KEY NEEDED)
python convert.py document.pdf

# Specify output
python convert.py document.pdf -o output.docx

# Optional: With API key for layout hints (1 API call max)
export ANTHROPIC_API_KEY='your-key'
python convert.py document.pdf
```

## With vs Without API Key

| Feature | Without API Key | With API Key |
|---------|-----------------|--------------|
| Single-column | Works | Works |
| Multi-column | Auto-detected | Hint-assisted |
| Tables | Full support | Full support |
| Images | Full support | Full support |
| API Calls | 0 | 1 (optional) |
| Cost | Free | ~$0.002 |

**The API key is completely optional.** It only provides hints for complex layouts.

## Example Output

```
============================================================
PDF TO DOCX CONVERTER v4.0
============================================================
Input:  document.pdf
Output: document.docx

[1/3] Layout hints: Skipped (no API key)
       Using automatic column detection

[2/3] Extracting content...
       Pages: 10
       Text blocks: 245
       Images: 12
       Tables: 3
       Columns detected: 2
       Total elements: 260

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

# Without API key (recommended for most documents)
result = convert("document.pdf")

# With API key (for complex layouts)
result = convert("document.pdf", api_key="your-key")

if result.success:
    print(f"Output: {result.output_path}")
    print(f"Layout: {result.layout_type}")
    print(f"Columns: {result.columns_detected}")
    print(f"Tables: {result.tables}")
    print(f"Images: {result.images}")
```

## Architecture

### PDF Extractor (`pdf_extractor.py`)
- Extracts text as individual lines (not blocks) for accuracy
- Detects columns by analyzing horizontal text distribution
- Extracts tables using PyMuPDF's `find_tables()`
- Groups lines into logical paragraphs
- Creates unified element list sorted by Y-position

### DOCX Generator (`docx_generator.py`)
- Processes elements in exact reading order
- Renders multi-column content using Word tables
- Places images inline at correct positions
- Renders tables with borders and formatting
- Preserves font formatting

### Layout Analyzer (`layout_analyzer.py`)
- **OPTIONAL** - only used if API key provided
- Single API call to get column hints
- Does NOT block conversion if unavailable

## Requirements

- Python 3.8+
- PyMuPDF (for PDF extraction)
- python-docx (for DOCX creation)
- Pillow (for image handling)
- anthropic (optional, for layout hints)

## Project Structure

```
PDFtoDOCX/
├── convert.py              # CLI
├── pdf_converter/
│   ├── __init__.py         # Package exports
│   ├── converter.py        # Main converter
│   ├── pdf_extractor.py    # Pure Python extraction
│   ├── docx_generator.py   # DOCX generation
│   └── layout_analyzer.py  # Optional Vision hints
└── requirements.txt
```

## Changelog

### v4.1.0
- **Major rewrite**: Works WITHOUT AI completely
- Automatic column detection from text positions
- Line-by-line text extraction for accuracy
- Improved table detection and rendering
- Images placed at correct Y-positions
- Optional layout hints (1 API call max)

### v4.0.0
- Added table detection
- Element interleaving by Y-position

### v3.0.0
- Initial multi-column support

## License

MIT
