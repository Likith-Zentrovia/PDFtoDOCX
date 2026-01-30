# PDF to DOCX Converter v3.0

**Intelligent PDF to DOCX conversion powered by Claude Vision AI**

This converter uses Claude's vision capabilities to analyze each page visually - understanding exact layout, columns, text formatting, and structure - then recreates the document accurately in DOCX format.

## How It Works

```
PDF Page → Claude Vision Analysis → Intelligent DOCX Generation
```

1. **Visual Analysis**: Each page is rendered as an image and sent to Claude Vision
2. **Layout Understanding**: Claude identifies columns, headers, footers, tables, images
3. **Content Extraction**: Text blocks are identified with their exact formatting
4. **Smart Reconstruction**: DOCX is generated preserving the original layout
5. **Auto-Validation**: Output is validated against original for quality assurance

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Your API Key

```bash
export ANTHROPIC_API_KEY='your-anthropic-api-key'
```

Get your API key from [console.anthropic.com](https://console.anthropic.com/)

### 3. Convert

```bash
python convert.py document.pdf
```

That's it. One command. The converter handles everything automatically.

## Usage

```bash
# Basic conversion (output: document.docx)
python convert.py document.pdf

# Specify output file
python convert.py document.pdf -o output.docx

# Convert specific pages (0-indexed)
python convert.py document.pdf --pages 0-5

# Pass API key directly
python convert.py document.pdf -k your-api-key

# Quiet mode (minimal output)
python convert.py document.pdf -q
```

## What Gets Preserved

| Feature | Description |
|---------|-------------|
| **Multi-Column Layouts** | 2-column, 3-column, and complex layouts |
| **Text Formatting** | Font size, bold, italic, colors |
| **Reading Order** | Correct order across columns |
| **Headers/Footers** | Identified and positioned correctly |
| **Images** | Extracted and placed accurately |
| **Tables** | Structure and content preserved |

## Example Output

```
╔══════════════════════════════════════════════════════════════════╗
║        PDF to DOCX Converter v3.0 - Claude Vision Powered        ║
╚══════════════════════════════════════════════════════════════════╝

============================================================
INTELLIGENT PDF TO DOCX CONVERTER
============================================================
Input:  document.pdf
Output: document.docx
Pages:  10

[1/4] Analyzing document with Claude Vision...
       Layout: two_column
       Type: article

[2/4] Creating document structure...

[3/4] Converting pages...
       Page 1: two_column (2 col, 24 blocks)
       Page 2: two_column (2 col, 31 blocks)
       ...

[4/4] Saving document...

============================================================
CONVERSION COMPLETE
============================================================
Text elements: 245
Images:        12
Tables:        3
Quality score: 95%

[SUCCESS] Conversion complete!
Output saved to: document.docx
```

## Python API

```python
from pdf_converter import convert_pdf

# Simple conversion
result = convert_pdf("document.pdf")
print(f"Saved to: {result.output_path}")
print(f"Quality: {result.validation_score:.0%}")

# With options
result = convert_pdf(
    "document.pdf",
    output_path="output.docx",
    api_key="your-key"  # Optional if env var is set
)

# Check result
if result.success:
    print(f"Converted {result.pages_converted} pages")
    print(f"Text elements: {result.text_elements_processed}")
    print(f"Images: {result.images_extracted}")
```

## Requirements

- Python 3.8+
- Anthropic API key (for Claude Vision)
- Dependencies: `anthropic`, `PyMuPDF`, `python-docx`, `Pillow`

## Installation

```bash
# Clone repository
git clone https://github.com/yourusername/PDFtoDOCX.git
cd PDFtoDOCX

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Set API key
export ANTHROPIC_API_KEY='your-key'

# Test
python convert.py your-document.pdf
```

## API Key Usage

The Anthropic API key is used **only** for visual analysis of PDF pages:
- Pages are converted to images locally
- Images are sent to Claude Vision for layout analysis
- No data is stored or extracted beyond the analysis
- Analysis results guide the DOCX generation

Typical usage: ~1 API call per page

## Troubleshooting

### "ANTHROPIC_API_KEY not set"
```bash
export ANTHROPIC_API_KEY='your-key-here'
# Or pass directly:
python convert.py document.pdf -k your-key
```

### "Module not found"
```bash
pip install -r requirements.txt
```

### Poor conversion quality
- Ensure the PDF is not scanned (requires OCR)
- Check if PDF has unusual layouts
- Verify API key has sufficient credits

## Project Structure

```
PDFtoDOCX/
├── convert.py                    # CLI entry point
├── pdf_converter/
│   ├── __init__.py              # Package exports
│   ├── vision_analyzer.py       # Claude Vision analysis
│   └── intelligent_converter.py # DOCX generation
├── requirements.txt             # Dependencies
└── README.md                    # Documentation
```

## License

MIT License
