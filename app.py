import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image, ImageTk, ImageEnhance, ImageOps
import pytesseract
import easyocr
import numpy as np
import torch
import gc
from datetime import datetime, timedelta
import sys


CURRENT_PROCESSING = {
    "pdf": None,
    "page": None,
    "total_pages": None
}


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

APP_LOG_DIR = os.path.join(os.getenv("APPDATA"), "PDFSplitter", "logs")
os.makedirs(APP_LOG_DIR, exist_ok=True)
log_file_path = os.path.join(APP_LOG_DIR, f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")

def clean_old_logs():
    for filename in os.listdir(APP_LOG_DIR):
        full_path = os.path.join(APP_LOG_DIR, filename)
        if os.path.isfile(full_path):
            created_time = datetime.fromtimestamp(os.path.getctime(full_path))
            if datetime.now() - created_time > timedelta(days=30):
                os.remove(full_path)

def log_text(pdf_name, page_number, extracted_id, final_path=None):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] [{pdf_name} - Page {page_number}]\n")
        if extracted_id:
            f.write(f"Extracted ID found: {extracted_id}\n")
            if final_path:
                f.write(f"Renamed and saved as: {final_path}\n")
        else:
            f.write("No ID extracted on this page.\n")
        f.write("\n")

def log_exception(context, error):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] ERROR in {context}:\n{error}\n\n")


easyocr_reader = easyocr.Reader(['en'], gpu=torch.cuda.is_available())

def preprocess_image(image):
    image = image.convert("L")
    image = ImageOps.autocontrast(image)
    image = ImageEnhance.Sharpness(image).enhance(2.0)
    return image

def extract_id_dismissal(image):
    try:
        image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail=0)
        text = "\n".join(results)
        matches = re.findall(r'(?:File\s*No[:.]?\s*)([A-Za-z0-9\-]+)', text, re.IGNORECASE)
        return matches[0] if matches else None
    except Exception as e:
        log_exception("extract_id_dismissal", e)
        return None


def extract_id_lien(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()

        for line in lines:
            line_lower = line.lower()
            if "case no" in line_lower:
                idx = line_lower.find("case no")
                after = line[idx + len("case no"):].strip(" .:_-")
                parts = after.split()
                if parts:
                    cleaned = parts[0].strip(" .:_-")
                    return cleaned
        return None
    except Exception as e:
        log_exception("extract_id_lien", e)
        return None




def get_unique_filename(base_path, base_name, extension=".pdf"):
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(base_path, filename)):
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    return os.path.join(base_path, filename)

def process_pdf(pdf_path, output_base, id_keyword, progress_callback, index, total_files):
    try:
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.join(output_base, pdf_name)
        os.makedirs(output_dir, exist_ok=True)

        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)

        for i, page in enumerate(reader.pages):
            CURRENT_PROCESSING["pdf"] = pdf_name
            CURRENT_PROCESSING["page"] = i + 1
            CURRENT_PROCESSING["total_pages"] = total_pages

            try:
                writer = PdfWriter()
                writer.add_page(page)
                temp_path = os.path.join(output_dir, f"__temp_page_{i+1}.pdf")
                with open(temp_path, 'wb') as f:
                    writer.write(f)

                image = convert_from_path(temp_path, dpi=300, poppler_path=resource_path("poppler-bin"))[0]

                os.remove(temp_path)

                if "fileno" in id_keyword.lower():
                    extracted_id = extract_id_dismissal(image)
                    notice_label = "Notice Of Dismissal"
                else:
                    extracted_id = extract_id_lien(image)
                    notice_label = "Notice Of Lien"

                if extracted_id:
                    base_filename = f"{extracted_id}_{notice_label}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, extracted_id, final_path)
                else:
                    log_text(pdf_name, i + 1, None)

            except Exception as e:
                log_exception("process_pdf", f"file-level error in {pdf_name}:\n{e}")

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_pdf", e)


class SplitPDFApp:
    def __init__(self, root):
        self.root = root
        root.title("Split Dismissal & Lien PDFs (Extract ID)")
        root.geometry("900x400")
        self.processing = False
        self.dismissal_folder = tk.StringVar()
        self.lien_folder = tk.StringVar()

        logo_path = resource_path("logo.png")
        if os.path.exists(logo_path):
            logo_image = Image.open(logo_path)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(root, image=logo_photo)
            logo_label.image = logo_photo
            logo_label.pack(pady=(10, 5))

        frame = tk.Frame(root)
        frame.pack(pady=10)

       
        left = tk.Frame(frame)
        left.grid(row=0, column=0, padx=30)
        tk.Label(left, text="Dismissal PDFs (FileNo)").pack()
        tk.Entry(left, textvariable=self.dismissal_folder, width=40).pack()
        tk.Button(left, text="Browse", command=self.browse_dismissal).pack(pady=2)
        self.progress_dismissal = ttk.Progressbar(left, length=300, mode="determinate")
        self.progress_dismissal.pack(pady=10)

   
        right = tk.Frame(frame)
        right.grid(row=0, column=1, padx=30)
        tk.Label(right, text="Lien PDFs (CaseNo)").pack()
        tk.Entry(right, textvariable=self.lien_folder, width=40).pack()
        tk.Button(right, text="Browse", command=self.browse_lien).pack(pady=2)
        self.progress_lien = ttk.Progressbar(right, length=300, mode="determinate")
        self.progress_lien.pack(pady=10)

    def browse_dismissal(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.dismissal_folder.set(path)
            self.run_type(path, "dismissal", "FileNo", self.progress_dismissal)

    def browse_lien(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.lien_folder.set(path)
            self.run_type(path, "lien", "CaseNo", self.progress_lien)

    def run_type(self, folder, keyword_match, id_keyword, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf') and keyword_match in f.lower()]
                if not pdfs:
                    messagebox.showerror("Error", f"No '{keyword_match}' PDFs found.")
                    return

                progressbar["value"] = 0
                total_files = len(pdfs)
                for idx, path in enumerate(pdfs):
                    process_pdf(path, folder, id_keyword, update_progress, idx, total_files)
                messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).")
                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type", e)
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()

    def on_closing(self):
        if CURRENT_PROCESSING["pdf"]:
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] WARNING: Program closed while processing "
                        f"{CURRENT_PROCESSING['pdf']} at page {CURRENT_PROCESSING['page']} of "
                        f"{CURRENT_PROCESSING['total_pages']}.\n")
        else:
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Program closed normally.\n")
        self.root.destroy()


if __name__ == "__main__":
    clean_old_logs()
    root = tk.Tk()
    app = SplitPDFApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
