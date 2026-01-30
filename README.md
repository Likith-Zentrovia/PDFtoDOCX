# PDF to DOCX Converter

Convert PDF to DOCX with accurate layout preservation.

## How It Works

```
PDF → [Vision: Layout Analysis] → [Python: Content Extraction] → DOCX
```

1. **Claude Vision** analyzes page structure (columns, headers) - NOT content
2. **PyMuPDF** extracts all text, images, and formatting
3. **python-docx** generates DOCX matching original structure

Vision is used **only** to understand layout complexity. All actual content extraction is done by Python.

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

# With API key for better multi-column detection
export ANTHROPIC_API_KEY='your-key'
python convert.py document.pdf
```

## With vs Without API Key

| Feature | Without API Key | With API Key |
|---------|-----------------|--------------|
| Single-column PDFs | Works | Works |
| Multi-column PDFs | Basic (may miss columns) | Accurate column detection |
| Speed | Fast | Slightly slower (3 API calls) |

The API key enables Claude Vision to analyze layout structure. Without it, the converter uses basic heuristics.

## What Gets Preserved

- **Text** - All text content with formatting
- **Fonts** - Size, bold, italic, color
- **Columns** - Multi-column layouts
- **Images** - Extracted and positioned
- **Reading Order** - Correct flow across columns

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

[3/3] Generating DOCX...

============================================================
SUCCESS!
============================================================
Output: document.docx
Pages:  10
Blocks: 245
Images: 12
```

## Python API

```python
from pdf_converter import convert

result = convert("document.pdf")

if result.success:
    print(f"Output: {result.output_path}")
    print(f"Layout: {result.layout_type}")
```

## Requirements

- Python 3.8+
- PyMuPDF
- python-docx
- anthropic (optional, for layout analysis)

## Project Structure

```
PDFtoDOCX/
├── convert.py              # CLI
├── pdf_converter/
│   ├── converter.py        # Main converter
│   ├── layout_analyzer.py  # Vision-based layout analysis
│   ├── pdf_extractor.py    # PyMuPDF content extraction
│   └── docx_generator.py   # DOCX generation
└── requirements.txt
```

## License

MIT
