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
import io
import subprocess

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

pytesseract.pytesseract.tesseract_cmd = os.path.join(resource_path("Tesseract-OCR"), "tesseract.exe")

APP_LOG_DIR = os.path.join(os.getenv("APPDATA"), "PDFSplitter", "logs")
os.makedirs(APP_LOG_DIR, exist_ok=True)

def clean_old_logs():
    for filename in os.listdir(APP_LOG_DIR):
        full_path = os.path.join(APP_LOG_DIR, filename)
        if os.path.isfile(full_path):
            created_time = datetime.fromtimestamp(os.path.getctime(full_path))
            if datetime.now() - created_time > timedelta(days=30):
                os.remove(full_path)

def log_text(pdf_name, page_number, extracted_id, log_file_path, final_path=None):
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

def log_exception(context, error, log_file_path):
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
        matches = re.findall(r'(?:File\s*No[:.;]?\s*)([A-Za-z0-9.,\-]+)', text, re.IGNORECASE)

        if matches:
            clean_id = re.sub(r'[.,]', '', matches[0])
            return clean_id
        return None
    except Exception as e:
        log_exception("extract_id_dismissal", e, log_file_path=None)
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
        log_exception("extract_id_lien", e, log_file_path=None)
        return None

def get_unique_filename(base_path, base_name, extension=".pdf"):
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(base_path, filename)):
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    return os.path.join(base_path, filename)

def process_pdf(pdf_path, output_base, id_keyword, progress_callback, index, total_files, log_file_path):
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

                if extracted_id:
                    base_filename = f"{extracted_id}_{notice_label}" if "fileno" in id_keyword.lower() else f"{extracted_id}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, extracted_id, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)

            except Exception as e:
                log_exception("process_pdf", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_pdf", e, log_file_path)

class SplitPDFApp:
    def __init__(self, root):
        self.root = root
        root.title("PDF Utility Suite - Splitter, Merger, Redaction, Compressor")
        root.geometry("950x500")
        self.processing = False
        self.dismissal_folder = tk.StringVar()
        self.lien_folder = tk.StringVar()
        self.latest_log_file = None

        # --- Top Logo ---
        logo_path = resource_path("logo.png")
        self.logo_frame = tk.Frame(root)
        self.logo_frame.pack(pady=(10, 5))
        if os.path.exists(logo_path):
            logo_image = Image.open(logo_path)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(self.logo_frame, image=logo_photo)
            logo_label.image = logo_photo
            logo_label.pack()

        # --- Feature Button Bar ---
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(pady=(0, 10))
        self.feature_buttons = {}
        features = [
            ("Splitter", self.show_splitter),
            ("Merger", self.show_merger),
            ("Redaction", self.show_redaction),
            ("Compressor", self.show_compressor)
        ]
        for i, (name, cmd) in enumerate(features):
            btn = tk.Button(self.button_frame, text=name, font=("Arial", 13, "bold"), width=18, height=2, command=cmd)
            btn.grid(row=0, column=i, padx=10)
            self.feature_buttons[name] = btn

        # --- Main Content Frames ---
        self.content_frame = tk.Frame(root)
        self.content_frame.pack(fill='both', expand=True)

        self.splitter_tab = tk.Frame(self.content_frame)
        self.merger_tab = tk.Frame(self.content_frame)
        self.redaction_tab = tk.Frame(self.content_frame)
        self.compressor_tab = tk.Frame(self.content_frame)

        self.init_splitter_tab()
        self.init_merger_tab()
        self.init_redaction_tab()
        self.init_compressor_tab()

        # Show Splitter by default
        self.show_splitter()

    def show_splitter(self):
        self._raise_tab(self.splitter_tab)
        self._highlight_button("Splitter")

    def show_merger(self):
        self._raise_tab(self.merger_tab)
        self._highlight_button("Merger")

    def show_redaction(self):
        self._raise_tab(self.redaction_tab)
        self._highlight_button("Redaction")

    def show_compressor(self):
        self._raise_tab(self.compressor_tab)
        self._highlight_button("Compressor")

    def _raise_tab(self, tab):
        for frame in [self.splitter_tab, self.merger_tab, self.redaction_tab, self.compressor_tab]:
            frame.pack_forget()
        tab.pack(fill='both', expand=True)

    def _highlight_button(self, name):
        for btn_name, btn in self.feature_buttons.items():
            if btn_name == name:
                btn.config(bg="#1976d2", fg="white")
            else:
                btn.config(bg="SystemButtonFace", fg="black")

    def init_splitter_tab(self):
        

        frame = tk.Frame(self.splitter_tab)
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

    def init_merger_tab(self):
        label = tk.Label(self.merger_tab, text="PDF Merger", font=("Arial", 14))
        label.pack(pady=10)
        
        self.merger_folder = tk.StringVar()
        self.merger_subfolders = []
        self.merger_files_var = tk.StringVar(value=[])
        
        select_btn = tk.Button(self.merger_tab, text="Select Folder to Merge PDFs by Subfolder", command=self.select_merger_folder)
        select_btn.pack(pady=5)
        
        self.folder_label = tk.Label(self.merger_tab, textvariable=self.merger_folder, fg="gray")
        self.folder_label.pack(pady=2)
        
        self.files_listbox = tk.Listbox(self.merger_tab, listvariable=self.merger_files_var, width=70, height=8)
        self.files_listbox.pack(pady=5)
        
        merge_btn = tk.Button(self.merger_tab, text="Merge PDFs in Each Subfolder", command=self.merge_pdfs_by_subfolder)
        merge_btn.pack(pady=10)

    def select_merger_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.merger_folder.set(folder)
            # List subfolders
            subfolders = [os.path.join(folder, d) for d in os.listdir(folder) if os.path.isdir(os.path.join(folder, d))]
            self.merger_subfolders = subfolders
            # List PDFs directly in the folder
            direct_pdfs = [f for f in os.listdir(folder) if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(folder, f))]
            display = []
            if direct_pdfs:
                display.append(f"[This Folder]: {len(direct_pdfs)} PDFs")
            display += [f"{os.path.basename(sf)}: {len([f for f in os.listdir(sf) if f.lower().endswith('.pdf')])} PDFs" for sf in subfolders]
            self.merger_files_var.set(display)
            self.merger_direct_pdfs = direct_pdfs

    def merge_pdfs_by_subfolder(self):
        folder = self.merger_folder.get()
        if not folder:
            messagebox.showerror("No Folder Selected", "Please select a folder with PDFs or subfolders containing PDFs.")
            return
        log_file_path = os.path.join(APP_LOG_DIR, f"merger_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        self.latest_log_file = log_file_path
        try:
            # Merge PDFs directly in the selected folder (if any)
            if hasattr(self, 'merger_direct_pdfs') and self.merger_direct_pdfs:
                merger = PdfWriter()
                pdf_files = [os.path.join(folder, f) for f in self.merger_direct_pdfs]
                for pdf_file in pdf_files:
                    try:
                        reader = PdfReader(pdf_file)
                        for page in reader.pages:
                            merger.add_page(page)
                    except Exception as e:
                        log_exception("merge_pdfs_by_subfolder", f"Failed to read {pdf_file}: {e}", log_file_path)
                        continue
                output_path = os.path.join(folder, f"{os.path.basename(folder)}.pdf")
                with open(output_path, "wb") as f_out:
                    merger.write(f_out)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] Merged PDF files in [This Folder]:\n")
                    for file in pdf_files:
                        f.write(f"  - {file}\n")
                    f.write(f"Saved merged PDF as: {output_path}\n\n")
            # Merge PDFs in each subfolder
            for subfolder in self.merger_subfolders:
                pdf_files = [os.path.join(subfolder, f) for f in os.listdir(subfolder) if f.lower().endswith('.pdf')]
                if not pdf_files:
                    continue
                merger = PdfWriter()
                for pdf_file in pdf_files:
                    try:
                        reader = PdfReader(pdf_file)
                        for page in reader.pages:
                            merger.add_page(page)
                    except Exception as e:
                        log_exception("merge_pdfs_by_subfolder", f"Failed to read {pdf_file}: {e}", log_file_path)
                        continue
                output_path = os.path.join(folder, f"{os.path.basename(subfolder)}.pdf")
                with open(output_path, "wb") as f_out:
                    merger.write(f_out)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] Merged PDF files in {subfolder}:\n")
                    for file in pdf_files:
                        f.write(f"  - {file}\n")
                    f.write(f"Saved merged PDF as: {output_path}\n\n")
            messagebox.showinfo("Success", f"Merged PDFs in folder and all subfolders. See log for details.")
        except Exception as e:
            log_exception("merge_pdfs_by_subfolder", e, log_file_path)
            messagebox.showerror("Error", f"Failed to merge PDFs:\n{e}")

    def init_redaction_tab(self):
        label = tk.Label(self.redaction_tab, text="PDF Redaction", font=("Arial", 14))
        label.pack(pady=20)
        # Placeholder for redaction UI

    def init_compressor_tab(self):
        label = tk.Label(self.compressor_tab, text="PDF Compressor", font=("Arial", 14))
        label.pack(pady=10)
        
        self.compress_input_file = None
        self.compress_input_folder = None
        self.compress_original_size_var = tk.StringVar(value="Original Size: N/A")
        self.compress_compressed_size_var = tk.StringVar(value="Compressed Size: N/A")
        
        select_file_btn = tk.Button(self.compressor_tab, text="Select PDF File to Compress", command=self.select_compress_file)
        select_file_btn.pack(pady=5)
        
        select_folder_btn = tk.Button(self.compressor_tab, text="Select Folder to Compress All PDFs", command=self.select_compress_folder)
        select_folder_btn.pack(pady=5)
        
        self.compress_file_label = tk.Label(self.compressor_tab, text="No file or folder selected", fg="gray")
        self.compress_file_label.pack(pady=2)
        
        tk.Label(self.compressor_tab, textvariable=self.compress_original_size_var).pack(pady=2)
        tk.Label(self.compressor_tab, textvariable=self.compress_compressed_size_var).pack(pady=2)
        
        compress_btn = tk.Button(self.compressor_tab, text="Compress PDF(s)", command=self.compress_pdf)
        compress_btn.pack(pady=10)

    def select_compress_file(self):
        file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file:
            self.compress_input_file = file
            self.compress_input_folder = None
            self.compress_file_label.config(text=os.path.basename(file), fg="black")
            size = os.path.getsize(file)
            self.compress_original_size_var.set(f"Original Size: {self.format_size(size)}")
            self.compress_compressed_size_var.set("Compressed Size: N/A")

    def select_compress_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.compress_input_folder = folder
            self.compress_input_file = None
            self.compress_file_label.config(text=f"Folder: {os.path.basename(folder)}", fg="black")
            self.compress_original_size_var.set("Original Size: N/A")
            self.compress_compressed_size_var.set("Compressed Size: N/A")

    def compress_pdf(self):
        if self.compress_input_file:
            self._compress_single_pdf(self.compress_input_file)
        elif self.compress_input_folder:
            self._compress_folder_pdfs(self.compress_input_folder)
        else:
            messagebox.showerror("No PDF or Folder Selected", "Please select a PDF file or a folder to compress.")

    def _compress_single_pdf(self, input_file):
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save Compressed PDF As")
        if not output_path:
            return
        log_file_path = os.path.join(APP_LOG_DIR, f"compressor_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        self.latest_log_file = log_file_path
        try:
            gs_exe = os.path.join(resource_path("ghostscript-bin"), "gswin64c.exe")
            gs_command = [
                gs_exe,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                "-dPDFSETTINGS=/ebook",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                input_file
            ]
            result = subprocess.run(gs_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode != 0:
                raise RuntimeError(f"Ghostscript error: {result.stderr.decode('utf-8')}")
            compressed_size = os.path.getsize(output_path)
            self.compress_compressed_size_var.set(f"Compressed Size: {self.format_size(compressed_size)}")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] Compressed PDF file: {input_file}\n")
                f.write(f"Saved compressed PDF as: {output_path}\n")
                f.write(f"Original size: {self.compress_original_size_var.get()}\n")
                f.write(f"Compressed size: {self.compress_compressed_size_var.get()}\n\n")
            messagebox.showinfo("Success", f"Compressed PDF saved to:\n{output_path}")
        except FileNotFoundError:
            log_exception("compress_pdf", "Ghostscript Not Found", log_file_path)
            messagebox.showerror("Ghostscript Not Found", "Ghostscript is not bundled or not found. Please ensure ghostscript-bin/gswin64c.exe is present.")
        except Exception as e:
            log_exception("compress_pdf", e, log_file_path)
            messagebox.showerror("Error", f"Failed to compress PDF:\n{e}")

    def _compress_folder_pdfs(self, input_folder):
        output_folder = input_folder.rstrip("/\\") + "_compressed"
        os.makedirs(output_folder, exist_ok=True)
        log_file_path = os.path.join(APP_LOG_DIR, f"compressor_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        self.latest_log_file = log_file_path
        pdfs_to_compress = []
        # PDFs directly in the folder
        for f in os.listdir(input_folder):
            if f.lower().endswith('.pdf') and os.path.isfile(os.path.join(input_folder, f)):
                pdfs_to_compress.append((os.path.join(input_folder, f), os.path.join(output_folder, f)))
        # PDFs in subfolders
        for root, dirs, files in os.walk(input_folder):
            if root == input_folder:
                continue  # already handled
            rel = os.path.relpath(root, input_folder)
            out_subfolder = os.path.join(output_folder, rel)
            os.makedirs(out_subfolder, exist_ok=True)
            for f in files:
                if f.lower().endswith('.pdf'):
                    pdfs_to_compress.append((os.path.join(root, f), os.path.join(out_subfolder, f)))
        count = 0
        for in_path, out_path in pdfs_to_compress:
            try:
                gs_exe = os.path.join(resource_path("ghostscript-bin"), "gswin64c.exe")
                gs_command = [
                    gs_exe,
                    "-sDEVICE=pdfwrite",
                    "-dCompatibilityLevel=1.4",
                    "-dPDFSETTINGS=/ebook",
                    "-dNOPAUSE",
                    "-dQUIET",
                    "-dBATCH",
                    f"-sOutputFile={out_path}",
                    in_path
                ]
                result = subprocess.run(gs_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                if result.returncode != 0:
                    raise RuntimeError(f"Ghostscript error: {result.stderr.decode('utf-8')}")
                count += 1
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] Compressed PDF file: {in_path}\n")
                    f.write(f"Saved compressed PDF as: {out_path}\n\n")
            except Exception as e:
                log_exception("compress_pdf", e, log_file_path)
        messagebox.showinfo("Done", f"Compressed {count} PDF(s). Output folder: {output_folder}")

    def format_size(self, size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024*1024:
            return f"{size_bytes/1024:.1f} KB"
        else:
            return f"{size_bytes/1024/1024:.2f} MB"

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

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                for idx, path in enumerate(pdfs):
                    process_pdf(path, folder, id_keyword, update_progress, idx, total_files, log_file_path)
                messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).")
                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()

    def on_closing(self):
        log_file_path = self.latest_log_file
        if log_file_path:
            with open(log_file_path, "a", encoding="utf-8") as f:
                if CURRENT_PROCESSING["pdf"]:
                    f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] WARNING: Program closed while processing "
                            f"{CURRENT_PROCESSING['pdf']} at page {CURRENT_PROCESSING['page']} of "
                            f"{CURRENT_PROCESSING['total_pages']}.\n")
                else:
                    f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Program closed normally.\n")
        self.root.destroy()

if __name__ == "__main__":
    clean_old_logs()
    root = tk.Tk()
    app = SplitPDFApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
