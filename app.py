import os
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import numpy as np
import easyocr
import torch
import gc


ocr_reader = easyocr.Reader(['en'], gpu=torch.cuda.is_available())

def log_text(pdf_name, page_number, extracted_id, output_base):
    log_path = os.path.join(output_base, "log.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{pdf_name} - Page {page_number}]\n")
        f.write(f"Extracted ID: {extracted_id or 'None'}\n\n")

def extract_text(image):
    try:
        image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = ocr_reader.readtext(np_image, detail=0)
        return "\n".join(results)
    except RuntimeError as e:
        if "CUDA out of memory" in str(e):
            reader_cpu = easyocr.Reader(['en'], gpu=False)
            np_image = np.array(image.convert("RGB"))
            results = reader_cpu.readtext(np_image, detail=0)
            return "\n".join(results)
        else:
            raise e

def extract_id(image, id_keyword):
    try:
        text = extract_text(image)

        keyword_variations = []
        if id_keyword.lower() == "caseno":
            keyword_variations = ["case no", "caseno", "case number", "case #"]
        elif id_keyword.lower() == "fileno":
            keyword_variations = ["file no", "fileno", "file number", "file #"]
        else:
            keyword_variations = [id_keyword]

        for kw in keyword_variations:
            regex_kw = re.sub(r' ', r'\\s+', kw)
            pattern = rf'{regex_kw}[^A-Za-z0-9]*_?([A-Za-z0-9][A-Za-z0-9\-_]+)'
            matches = re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE)
            for match in matches:
                candidate = match.group(1)
                if len(candidate) >= 3:
                    return candidate
        return None
    except Exception as e:
        return None

def get_unique_filename(base_path, base_name, extension=".pdf"):
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(base_path, filename)):
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    return os.path.join(base_path, filename)

def process_pdf(pdf_path, output_base, id_keyword, progress_callback, index, total_files):
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

            image = convert_from_path(temp_path, dpi=150)[0]
            os.remove(temp_path)

            extracted_id = extract_id(image, id_keyword)

            log_text(pdf_name, i + 1, extracted_id, output_base)

            if extracted_id:
                base_filename = f"{extracted_id}_NoticeOfDismissal"
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
        path = filedialog.askdirectory()
        if path:
            self.dismissal_folder.set(path)
            self.run_type(path, "dismissal", "FileNo", self.progress_dismissal)

    def browse_lien(self):
        path = filedialog.askdirectory()
        if path:
            self.lien_folder.set(path)
            self.run_type(path, "lien", "CaseNo", self.progress_lien)

    def run_type(self, folder, keyword_match, id_keyword, progressbar):
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

            progressbar["value"] = 0
            total_files = len(pdfs)

            def update_progress(val):
                progressbar["value"] = val
                self.root.update_idletasks()

            try:
                for idx, path in enumerate(pdfs):
                    process_pdf(path, folder, id_keyword, update_progress, idx, total_files)
                messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        threading.Thread(target=worker).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = SplitPDFApp(root)
    root.mainloop()
