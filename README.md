# PDF Utility Suite - Splitter, Merger, Redaction, Compressor

This Python application provides a comprehensive suite of PDF utilities:

## Features

### Splitter
- Search for specific **keywords** in PDF files
- Extract **alphanumeric IDs** located next to ID-related keywords (e.g. "fileNo: 12345", "case no: ABC123")
- Save matched pages to new PDF files with extracted IDs as filenames
- **NEW**: Generate Excel/CSV reports with columns:
  - CaseNo/FileNo (extracted ID)
  - Current Datestamp (when processing occurred)
  - PDF Modified Date (original file modification date)
  - Source Path (original PDF file path)

### Merger
- Remove permissions from PDFs
- Merge multiple PDFs into a single file

### Redaction
- PDF redaction capabilities (coming soon)

### Compressor
- Compress PDF files to reduce file size

## Technical Features

- Supports `.pdf` files
- Uses **easyOCR** and **OCR (Tesseract)** for image-based text extraction
- Converts PDFs to images when necessary
- Logs metadata in `.txt` files
- Generates Excel (.xlsx) and CSV reports for splitter operations
- Simple **Graphic User Interface**

## Requirements

Install dependencies using:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl pytesseract pillow PyPDF2 pdf2image easyocr numpy torch
```
