from pdf2image import convert_from_path
import easyocr
import numpy as np
from PIL import Image, ImageTk, ImageEnhance, ImageOps


reader = easyocr.Reader(['en'])  
def ocr_with_easyocr(pdf_path):
    target_pages = [14, 56, 104]  # 1-based page numbers
    pages = convert_from_path(pdf_path, dpi=300, first_page=min(target_pages), last_page=max(target_pages))

    for i, page_num in enumerate(range(min(target_pages), max(target_pages) + 1)):
        if page_num in target_pages:
            image = pages[i]
            image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
            np_image = np.array(image.convert("RGB"))

            results = reader.readtext(np_image, detail=0)
            print(f"\n--- Page {page_num} ---")
            print("\n".join(results))


ocr_with_easyocr("C:\\Users\\Owner\\Downloads\\New folder\\04252025_Dismissal_Nova1.pdf")
