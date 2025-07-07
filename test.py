
import os
from pdf2image import convert_from_path
import easyocr
from PIL import Image

def extract_text_with_easyocr(pdf_path):
    # Convert PDF pages to images
    pages = convert_from_path(pdf_path, dpi=300)
    reader = easyocr.Reader(['en'], gpu=False)

    full_text = ""
    for i, page in enumerate(pages):
        print(f"Processing page {i + 1}/{len(pages)}...")
        # Convert PIL image to RGB (required by EasyOCR)
        page = page.convert("RGB")
        # OCR using EasyOCR
        result = reader.readtext(np.array(page), detail=0)
        page_text = "\n".join(result)
        full_text += f"\n\n=== Page {i + 1} ===\n{page_text}"

    return full_text

# Update this path (use forward slashes!)
pdf_path = "C:\\Users\\Owner\\Downloads\\New folder (2)\\04252025_Lien_Nova1.pdf"  # Replace with your PDF file

if not os.path.exists(pdf_path):
    print("PDF not found!")
else:
    import numpy as np  # Required by EasyOCR
    text = extract_text_with_easyocr(pdf_path)
    print(text)

# Path to your PDF


