# 📄 PDF/Image Keyword & ID Extractor

This Python script allows you to:

- 🔍 Search for specific **keywords** in PDF or image files.
- 🆔 Extract **alphanumeric IDs** located next to an ID-related keyword (e.g. "fileNo: 12345").
- 📁 Save matched pages to a new PDF file.
- 🗃️ Log all activity (filename, keyword, timestamp, extracted ID) into a SQLite database.

## 📦 Features

- Supports `.pdf`, `.jpg`, `.jpeg`, `.png` files.
- Uses **OCR (Tesseract)** for image-based text extraction.
- Converts PDFs to images when necessary.
- Logs metadata in `database.db`.
- Simple **CLI interface**.

## 🛠️ Requirements

Install dependencies using:

```bash
pip install pytesseract pillow PyPDF2 pdf2image
```
