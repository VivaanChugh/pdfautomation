import os
import datetime
import sqlite3
from flask import Flask, request, render_template
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
            timestamp TEXT,
            extracted_id TEXT
        )''')
init_db()

# Extract alphanumeric ID from custom cropped region
def extract_unique_id_from_fixed_region(image, left, top, right, bottom):
    cropped = image.crop((left, top, right, bottom))
    raw_text = pytesseract.image_to_string(cropped)
    alphanumeric_id = ''.join(filter(str.isalnum, raw_text))
    return alphanumeric_id or "N/A"

# Extract text from full image
def extract_text_from_image(image):
    return pytesseract.image_to_string(image)

# Main logic to process image or PDF with keyword search and crop bounds
def search_and_split(file_path, keyword, crop_bounds):
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

        extracted_id = extract_unique_id_from_fixed_region(image, *crop_bounds)

    elif ext == '.pdf':
        reader = PdfReader(file_path)
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            if keyword.lower() in text.lower():
                matched_pages.append(i)

        if not matched_pages:
            images = convert_from_path(file_path)
            for i, image in enumerate(images):
                text = extract_text_from_image(image)
                if keyword.lower() in text.lower():
                    matched_pages.append(i)
                if extracted_id is None:
                    extracted_id = extract_unique_id_from_fixed_region(image, *crop_bounds)
        else:
            images = convert_from_path(file_path, first_page=1, last_page=1)
            extracted_id = extract_unique_id_from_fixed_region(images[0], *crop_bounds)

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

# Main route
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        keyword = request.form['keyword']
        file = request.files['file']

        # Get crop bounds from form
        try:
            left = int(request.form['left'])
            top = int(request.form['top'])
            right = int(request.form['right'])
            bottom = int(request.form['bottom'])
            crop_bounds = (left, top, right, bottom)
        except Exception as e:
            return f"❌ Invalid crop values: {e}"

        if file and keyword:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            result_file, extracted_id = search_and_split(file_path, keyword, crop_bounds)
            return render_template('index.html', result=result_file, extracted_id=extracted_id)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
