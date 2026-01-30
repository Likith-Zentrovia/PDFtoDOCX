# PDF to DOCX Converter

A high-fidelity PDF to DOCX converter that runs locally via command line. Preserves text formatting, images, tables, and layout structure with maximum accuracy.

## Features

- **Text Preservation**: Maintains font family, size, weight (bold), style (italic), color, and underline
- **Image Handling**: Extracts and positions images accurately within the document
- **Table Detection**: Automatically detects and recreates table structures with proper cell formatting
- **Layout Retention**: Preserves margins, columns, spacing, and page structure
- **Batch Processing**: Convert multiple PDFs at once or entire directories
- **Page Selection**: Convert specific pages or page ranges
- **Zero Content Loss**: Designed for 100% content preservation

## Installation

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. Clone the repository:
```bash
git clone https://github.com/yourusername/PDFtoDOCX.git
cd PDFtoDOCX
```

2. Create a virtual environment (recommended):
```bash
python -m venv venv

# On Linux/Mac:
source venv/bin/activate

# On Windows:
venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Conversion

Convert a single PDF file (output will be saved as `document.docx` in the same location):

```bash
python convert.py document.pdf
```

### Specify Output Path

```bash
python convert.py document.pdf -o /path/to/output.docx
```

### Convert Specific Pages

Pages are 0-indexed (first page is 0):

```bash
# Convert pages 1, 2, and 3 (0-indexed)
python convert.py document.pdf --pages 0,1,2

# Convert a range of pages
python convert.py document.pdf --pages 0-5

# Mix of specific pages and ranges
python convert.py document.pdf --pages 0,2-4,7,9-12
```

### Convert Page Range

```bash
# Convert from page 2 to page 10 (0-indexed, end is exclusive)
python convert.py document.pdf --start 2 --end 10
```

### Batch Conversion

Convert multiple files:

```bash
python convert.py file1.pdf file2.pdf file3.pdf

# With output directory
python convert.py file1.pdf file2.pdf -o /output/directory/
```

### Directory Conversion

Convert all PDFs in a directory:

```bash
# Convert all PDFs in a directory
python convert.py --dir /path/to/pdfs/

# Convert recursively (including subdirectories)
python convert.py --dir /path/to/pdfs/ -r

# Specify output directory
python convert.py --dir /path/to/pdfs/ -o /output/directory/
```

### Show PDF Information

Display PDF metadata before converting:

```bash
python convert.py document.pdf --info
```

### Quiet Mode

Suppress progress output:

```bash
python convert.py document.pdf -q
```

## Python API Usage

You can also use the converter programmatically in your Python code:

```python
from pdf_converter import PDFConverter

# Create converter instance
converter = PDFConverter(verbose=True)

# Convert a single file
result = converter.convert("input.pdf", "output.docx")

if result.success:
    print(f"Converted {result.pages_converted} pages to {result.output_path}")
else:
    print(f"Error: {result.error_message}")

# Convert specific pages
result = converter.convert("input.pdf", pages=[0, 1, 2])

# Batch convert multiple files
results = converter.batch_convert(["file1.pdf", "file2.pdf"])

# Convert all PDFs in a directory
results = converter.convert_directory("/path/to/pdfs", recursive=True)

# Get PDF information
info = converter.get_pdf_info("document.pdf")
print(f"Pages: {info['page_count']}")
print(f"Images: {sum(p['image_count'] for p in info['pages'])}")
```

### Quick Conversion Function

For simple one-off conversions:

```python
from pdf_converter.converter import convert_pdf_to_docx

result = convert_pdf_to_docx("document.pdf")
print(f"Output: {result.output_path}")
```

## Command Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--output` | `-o` | Output file path or directory |
| `--pages` | `-p` | Specific pages to convert (e.g., "0,1,2" or "0-5") |
| `--start` | | Start page (0-indexed) |
| `--end` | | End page (0-indexed, exclusive) |
| `--dir` | `-d` | Directory containing PDFs to convert |
| `--recursive` | `-r` | Process subdirectories recursively |
| `--info` | | Show PDF information before converting |
| `--quiet` | `-q` | Suppress progress output |
| `--version` | `-v` | Show version number |

## How It Works

The converter uses the `pdf2docx` library which leverages:

1. **PyMuPDF (fitz)**: For parsing PDF structure, extracting text, images, and detecting layout
2. **python-docx**: For generating the DOCX output with proper formatting

### Conversion Process

1. **Parse PDF**: Extract text blocks, images, tables, and layout information
2. **Analyze Structure**: Detect paragraphs, headings, lists, and table structures
3. **Map Formatting**: Convert PDF styling (fonts, colors, sizes) to DOCX equivalents
4. **Position Elements**: Place images and tables with accurate positioning
5. **Generate DOCX**: Create the output document with all preserved formatting

## Limitations

While the converter achieves high fidelity, some edge cases may have limitations:

- **Complex layouts**: Very complex multi-column layouts may need manual adjustment
- **Custom fonts**: If the original PDF uses custom fonts not available on your system, a similar font will be substituted
- **Scanned PDFs**: PDFs that are scanned images (not searchable) require OCR (not included)
- **Form fields**: Interactive PDF form fields are converted to static content
- **Annotations**: PDF annotations may not be fully preserved

## Troubleshooting

### "No module named 'pdf2docx'"
Make sure you've installed the dependencies:
```bash
pip install -r requirements.txt
```

### Conversion takes a long time
Large PDFs with many images take longer to process. Use `--pages` to convert specific pages for testing.

### Output formatting looks different
Try converting with the `--info` flag to check the PDF structure. Some PDFs have unusual formatting that may not convert perfectly.

### Memory errors on large PDFs
For very large PDFs, try converting in batches of pages:
```bash
python convert.py large.pdf --pages 0-50 -o part1.docx
python convert.py large.pdf --pages 50-100 -o part2.docx
```

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

## License

MIT License - See LICENSE file for details.
