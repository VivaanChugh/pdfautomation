import os
import datetime
import sqlite3
import re
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import pytesseract

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
DB_PATH = 'database.db'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# DB setup
def init_db():
    with sqlite3.connect(DB_PATH) as co:
        co.execute('''CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pdf_name TEXT,
            keyword TEXT,
            timestamp TEXT,
            extracted_id TEXT
        )''')
init_db()

# Extract alphanumeric ID next to a keyword
def extract_id_next_to_keyword(image, id_keyword):
    text = pytesseract.image_to_string(image)
    lines = text.split('\n')

    for line in lines:
        if id_keyword.upper() in line.upper():
            words = line.strip().replace("=", " ").replace(":", " ").split()
            for i, word in enumerate(words):
                if id_keyword.upper() in word.upper():
                    if i + 1 < len(words):
                        next_word = words[i + 1]
                        return next_word if next_word else "N/A"
    return "N/A"

# Extract text from full image
def extract_text_from_image(image):
    return pytesseract.image_to_string(image)

# Main logic to process image or PDF with keyword search and ID extraction
def search_and_split(file_path, keyword, id_keyword):
    ext = os.path.splitext(file_path)[-1].lower()
    matched_pages = []
    extracted_id = None
    output_name = None

    if ext in ['.jpg', '.jpeg', '.png']:
        image = Image.open(file_path)
        text = extract_text_from_image(image)

        if keyword.lower() in text.lower():
            output_path = os.path.join(OUTPUT_FOLDER, os.path.basename(file_path))
            image.save(output_path)
            output_name = os.path.basename(file_path)

        extracted_id = extract_id_next_to_keyword(image, id_keyword)

    elif ext == '.pdf':
        reader = PdfReader(file_path)
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            if keyword.lower() in text.lower():
                matched_pages.append(i)

        images = convert_from_path(file_path)

        for i, image in enumerate(images):
            if extracted_id is None:
                extracted_id = extract_id_next_to_keyword(image, id_keyword)

            if i in matched_pages:
                continue
            text = extract_text_from_image(image)
            if keyword.lower() in text.lower():
                matched_pages.append(i)

        if matched_pages:
            writer = PdfWriter()
            for i in matched_pages:
                writer.add_page(reader.pages[i])
            output_name = os.path.basename(file_path).replace(".pdf", f"_{keyword}.pdf")
            output_path = os.path.join(OUTPUT_FOLDER, output_name)
            with open(output_path, "wb") as f:
                writer.write(f)

    # Log metadata
    with sqlite3.connect(DB_PATH) as co:
        co.execute("INSERT INTO logs (pdf_name, keyword, timestamp, extracted_id) VALUES (?, ?, ?, ?)",
                   (os.path.basename(file_path), keyword, datetime.datetime.now().isoformat(), extracted_id))

    return output_name, extracted_id

# -------------------------------
# CLI-Based Main Function
# -------------------------------
def main():
    print("📄 Keyword & ID Extractor")
    file_path = input("Enter path to PDF/Image file (without quotations): ").strip()
    if not os.path.isfile(file_path):
        print("File not found.")
        return

    keyword = input("Enter search keyword (e.g. Invoice): ").strip()
    id_keyword = input("Enter ID keyword (e.g. fileNo): ").strip()

    upload_path = os.path.join(UPLOAD_FOLDER, os.path.basename(file_path))
    if file_path != upload_path:
        with open(file_path, "rb") as src, open(upload_path, "wb") as dst:
            dst.write(src.read())

    result_file, extracted_id = search_and_split(upload_path, keyword, id_keyword)

    
    if result_file:
        print(f"Matched pages saved as: {os.path.join(OUTPUT_FOLDER, result_file)}")
    print(f"Extracted ID: {extracted_id}")

if __name__ == '__main__':
    main()
