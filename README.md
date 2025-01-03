# doc2pdf

A Python tool to convert MHTML (.doc) files exported from Confluence to PDF format. This script handles MIME structure, QuotedPrintable encoding, and Microsoft Office-specific elements to produce clean PDF output.

## Features

- Converts MHTML/DOC files to PDF
- Handles QuotedPrintable encoding
- Cleans Microsoft Office-specific elements
- Maintains document structure and formatting
- Supports both command line and Python API usage

## Usage

```bash
python doc2pdf.py <input_file> <output_file>
# or run in docker
docker run -it --rm -v $(pwd):/app soulmachine/doc2pdf <input_file> <output_file>
```

Or

```bash
python doc2pdf.py input_dir output_dir
# or run in docker
docker run -it --rm -v doc_dir:/input_dir pdf_dir:/output_dir soulmachine/doc2pdf /input_dir /output_dir
```
