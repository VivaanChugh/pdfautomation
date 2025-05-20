import os
import datetime
import sqlite3
from flask import Flask, request, render_template, redirect, url_for
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import pytesseract

app = Flask(__name__)
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
            timestamp TEXT
        )''')
init_db()

# Text extraction from image using OCR
def extract_text_from_image(image):
    return pytesseract.image_to_string(image)

# Main logic to process PDF or image
def search_and_split(file_path, keyword):
    ext = os.path.splitext(file_path)[-1].lower()

    matched_pages = []

    if ext in ['.jpg', '.jpeg', '.png']:
        # Single image file
        text = extract_text_from_image(Image.open(file_path))
        if keyword.lower() in text.lower():
            output_path = os.path.join(OUTPUT_FOLDER, os.path.basename(file_path))
            Image.open(file_path).save(output_path)

            with sqlite3.connect(DB_PATH) as co:
                co.execute("INSERT INTO logs (pdf_name, keyword, timestamp) VALUES (?, ?, ?)",
                           (os.path.basename(file_path), keyword, datetime.datetime.now().isoformat()))
            return os.path.basename(file_path)

    elif ext == '.pdf':
        # Try regular PDF reading first
        reader = PdfReader(file_path)
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            if keyword.lower() in text.lower():
                matched_pages.append(i)

        # If no matches, try OCR on scanned PDF
        if not matched_pages:
            images = convert_from_path(file_path)
            for i, image in enumerate(images):
                text = extract_text_from_image(image)
                if keyword.lower() in text.lower():
                    matched_pages.append(i)

        # Save matched pages
        if matched_pages:
            writer = PdfWriter()
            reader = PdfReader(file_path)
            for i in matched_pages:
                writer.add_page(reader.pages[i])
            output_name = os.path.basename(file_path).replace(".pdf", f"_{keyword}.pdf")
            output_path = os.path.join(OUTPUT_FOLDER, output_name)
            with open(output_path, "wb") as f:
                writer.write(f)

            with sqlite3.connect(DB_PATH) as co:
                co.execute("INSERT INTO logs (pdf_name, keyword, timestamp) VALUES (?, ?, ?)",
                           (os.path.basename(file_path), keyword, datetime.datetime.now().isoformat()))
            return output_name

    return None

# Routes
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        keyword = request.form['keyword']
        file = request.files['file']
        if file and keyword:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            result_file = search_and_split(file_path, keyword)
            return render_template('index.html', result=result_file)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
