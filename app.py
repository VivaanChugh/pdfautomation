import os
import re
import gc
import torch
import pytesseract
import easyocr
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
from PIL import ImageEnhance, ImageOps

# Init OCR
ocr_reader = easyocr.Reader(['en'], gpu=torch.cuda.is_available())

# Generate timestamped log file in AppData
appdata_dir = os.getenv('APPDATA') or os.path.expanduser("~")
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_path = os.path.join(appdata_dir, f"log_{timestamp}.txt")

def log_text(pdf_name, page_number, extracted_id):
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{pdf_name} - Page {page_number}]\n")
        f.write(f"Extracted ID: {extracted_id or 'None'}\n\n")

def extract_id(image, id_keyword, notice_type):
    try:
        # OCR: EasyOCR for dismissal, pytesseract for lien
        if notice_type.lower() == "dismissal":
            image_resized = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
            np_image = np.array(image_resized.convert("RGB"))
            text = " ".join(ocr_reader.readtext(np_image, detail=0)).replace("\n", " ")
        else:
            gray = image.convert("L")
            text = pytesseract.image_to_string(gray)
        print(text)
        text_lower = text.lower()
        
        # Define keyword variations
        if id_keyword.lower() == "caseno":
            keyword_variations = ["case no", "caseno", "case number", "case #"]
        elif id_keyword.lower() == "fileno":
            keyword_variations = ["file no", "fileno", "file number", "file #"]
        else:
            keyword_variations = [id_keyword.lower()]

        for kw in keyword_variations:
            if kw in text_lower:
                idx = text_lower.find(kw)
                after_kw_original = text[idx + len(kw):]

                # Extract the first ID-looking token (starts with digit or contains both letters and numbers)
                tokens = re.findall(r"\b[\w\-]+\b", after_kw_original)
                for token in tokens:
                    if len(token) >= 4 and (any(c.isdigit() for c in token)):
                        return token[:50]
        return None
    except Exception as e:
        print(f"Error in extract_id: {e}")
        return None



def get_unique_filename(base_path, base_name, extension=".pdf"):
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(base_path, filename)):
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    return os.path.join(base_path, filename)

def process_pdf(pdf_path, output_base, id_keyword, notice_type, progress_callback, index, total_files):
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_dir = os.path.join(output_base, pdf_name)
    os.makedirs(output_dir, exist_ok=True)

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    for i, page in enumerate(reader.pages):
        try:
            writer = PdfWriter()
            writer.add_page(page)

            temp_path = os.path.join(output_dir, f"__temp_page_{i+1}.pdf")
            with open(temp_path, 'wb') as f:
                writer.write(f)

            image = convert_from_path(temp_path, dpi=225)[0]
            image = ImageOps.autocontrast(image)
            image = ImageEnhance.Sharpness(image).enhance(2.0)
            os.remove(temp_path)

            extracted_id = extract_id(image, id_keyword, notice_type)
            log_text(pdf_name, i + 1, extracted_id)

            if extracted_id:
                safe_id = re.sub(r'[^\w\-]', '', extracted_id)[:50]
                if not safe_id:
                    safe_id = "UnknownID"
                base_filename = f"{safe_id}_Notice Of {notice_type}"
                final_path = get_unique_filename(output_dir, base_filename)
                with open(final_path, 'wb') as out_f:
                    writer.write(out_f)

        except Exception as e:
            print(f"Error processing page {i+1} of {pdf_name}: {e}")

        gc.collect()
        if torch.cuda.is_available():
            torch.cuda.empty_cache()

        progress = ((index + (i + 1) / total_pages) / total_files) * 100
        progress_callback(progress)

class SplitPDFApp:
    def __init__(self, root):
        self.root = root
        root.title("Split Dismissal & Lien PDFs (Extract ID)")
        root.geometry("900x320")

        self.dismissal_folder = tk.StringVar()
        self.lien_folder = tk.StringVar()
        self.processing = False

        frame = tk.Frame(root)
        frame.pack(pady=10)

        left = tk.Frame(frame)
        left.grid(row=0, column=0, padx=30)
        tk.Label(left, text="Dismissal PDFs (FileNo)").pack()
        tk.Entry(left, textvariable=self.dismissal_folder, width=40).pack()
        self.dismissal_btn = tk.Button(left, text="Browse", command=self.browse_dismissal)
        self.dismissal_btn.pack(pady=2)
        self.progress_dismissal = ttk.Progressbar(left, length=300, mode="determinate")
        self.progress_dismissal.pack(pady=10)

        right = tk.Frame(frame)
        right.grid(row=0, column=1, padx=30)
        tk.Label(right, text="Lien PDFs (CaseNo)").pack()
        tk.Entry(right, textvariable=self.lien_folder, width=40).pack()
        self.lien_btn = tk.Button(right, text="Browse", command=self.browse_lien)
        self.lien_btn.pack(pady=2)
        self.progress_lien = ttk.Progressbar(right, length=300, mode="determinate")
        self.progress_lien.pack(pady=10)

    def disable_buttons(self, state):
        self.dismissal_btn["state"] = state
        self.lien_btn["state"] = state

    def browse_dismissal(self):
        if self.processing:
            return
        path = filedialog.askdirectory()
        if path:
            self.dismissal_folder.set(path)
            self.run_type(path, "dismissal", "FileNo", "Dismissal", self.progress_dismissal)

    def browse_lien(self):
        if self.processing:
            return
        path = filedialog.askdirectory()
        if path:
            self.lien_folder.set(path)
            self.run_type(path, "lien", "CaseNo", "Lien", self.progress_lien)

    def run_type(self, folder, keyword_match, id_keyword, notice_type, progressbar):
        def worker():
            if not os.path.isdir(folder):
                messagebox.showerror("Error", "Invalid folder path.")
                return

            pdfs = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith('.pdf') and keyword_match in f.lower()
            ]

            if not pdfs:
                messagebox.showerror("Error", f"No '{keyword_match}' PDFs found.")
                return

            self.processing = True
            self.disable_buttons("disabled")

            progressbar["value"] = 0
            total_files = len(pdfs)

            def update_progress(val):
                progressbar["value"] = val
                self.root.update_idletasks()

            try:
                for idx, path in enumerate(pdfs):
                    process_pdf(path, folder, id_keyword, notice_type, update_progress, idx, total_files)
                messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).")
            except Exception as e:
                messagebox.showerror("Error", str(e))
            finally:
                self.processing = False
                self.disable_buttons("normal")

        threading.Thread(target=worker).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = SplitPDFApp(root)
    root.mainloop()
