from pdf2image import convert_from_path
import pytesseract



def ocr_with_tesseract(pdf_path):
    target_pages = [14, 56, 104]  # Page numbers (1-based index)
    pages = convert_from_path(pdf_path, dpi=300, first_page=min(target_pages), last_page=max(target_pages))

    for i, page_num in enumerate(range(min(target_pages), max(target_pages) + 1)):
        if page_num in target_pages:
            text = pytesseract.image_to_string(pages[i])
            print(f"\n--- Page {page_num} ---")
            print(text.strip())


ocr_with_tesseract("C:\\Users\\Owner\\Downloads\\New folder\\04252025_Dismissal_Nova1.pdf")
