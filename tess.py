from pdf2image import convert_from_path
import pytesseract

def ocr_with_tesseract(pdf_path):
    pages = convert_from_path(pdf_path, dpi=300)
    for i, image in enumerate(pages):
        text = pytesseract.image_to_string(image)
        print(f"\n--- Page {i + 1} ---")
        print(text.strip())

# Example usage
ocr_with_tesseract("C:\\Users\\Owner\\Downloads\\New folder (2)\\04252025_Lien_Nova1.pdf")
