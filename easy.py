from pdf2image import convert_from_path
import easyocr
import numpy as np

reader = easyocr.Reader(['en'])  # Use GPU: easyocr.Reader(['en'], gpu=True)

def ocr_with_easyocr(pdf_path):
    pages = convert_from_path(pdf_path, dpi=300)
    for i, image in enumerate(pages):
        np_image = np.array(image)
        results = reader.readtext(np_image, detail=0)
        print(f"\n--- Page {i + 1} ---")
        print("\n".join(results))

# Example usage
ocr_with_easyocr("C:\Users\Owner\Downloads\New folder (2)\04252025_Lien_Nova1.pdf")
