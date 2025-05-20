# PDF Keyword Search and OCR Extraction Tool

This project allows users to upload PDF files or images (JPEG, PNG) to a web application. The tool searches for a specified keyword within the document or image and extracts pages that contain the keyword. If the document is a scanned PDF (image-based), Optical Character Recognition (OCR) is used to detect text from the images.

## Features

- Search for a keyword in PDF files and image files (JPEG, PNG).
- Support for both text-based PDFs and image-based (scanned) PDFs using OCR.
- Extract matched pages from PDFs or save the image with detected text.
- Log all searches with metadata (PDF name, keyword, timestamp) in an SQLite database.
- Flask web interface for file upload and keyword search.

## Technologies Used

- **Flask**: Web framework for building the application.
- **SQLite**: Lightweight database for storing logs.
- **PyPDF2**: Library for handling PDF files.
- **pdf2image**: Converts PDF pages to images.
- **Pillow (PIL)**: Python Imaging Library for handling images.
- **pytesseract**: OCR tool for extracting text from images.

## Prerequisites

Before running the application, you need to install the required Python dependencies. Make sure you have Python 3.7+ installed.

### Dependencies
You can install the required libraries using `pip`:

```bash
pip install flask pypdf2 pdf2image pillow pytesseract
