import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import pytesseract

# --- Configuration --- #
SPLIT_FOLDER = 'split_output'
PROCESSED_FOLDER = 'processed_output'

os.makedirs(SPLIT_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def extract_text(image):
    return pytesseract.image_to_string(image)

def extract_id(image, id_keyword):
    try:
        text = extract_text(image)
        lines = text.split('\n')
        normalized_keyword = re.sub(r'[:=]+$', '', id_keyword.strip(), flags=re.IGNORECASE)

        for line in lines:
            line = line.strip()

         
            match = re.search(
                fr'{re.escape(normalized_keyword)}\s*[:=]?\s*([\$a-zA-Z0-9\-]+)',
                line,
                re.IGNORECASE
            )
            if match:
                id_raw = match.group(1)
                return id_raw.replace('$', 'S')



        return "N/A"
    except Exception as e:
        return f"Error: {e}"
import os



def split_pdfs(pdf_folder):
    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, filename)
            reader = PdfReader(pdf_path)
            base_name = os.path.splitext(filename)[0]
            output_dir = os.path.join(SPLIT_FOLDER, base_name)
            os.makedirs(output_dir, exist_ok=True)

            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                single_page_path = os.path.join(output_dir, f"page_{i+1}.pdf")
                with open(single_page_path, 'wb') as f:
                    writer.write(f)

def process_pages(operation, keyword, id_keyword):
    for root, dirs, files in os.walk(SPLIT_FOLDER):
        for file in files:
            if file.lower().endswith(".pdf"):
                page_path = os.path.join(root, file)
                try:
                    images = convert_from_path(page_path)
                except Exception as e:
                    print(f"Error converting {page_path}: {e}")
                    continue
                image = images[0]  # Single-page

                result_dir = os.path.join(PROCESSED_FOLDER, os.path.basename(os.path.dirname(page_path)))
                os.makedirs(result_dir, exist_ok=True)

                if operation == "keyword_search":
                    text = extract_text(image)
                    if keyword.lower() in text.lower():
                        shutil.copy(page_path, os.path.join(result_dir, file))

                elif operation == "id_extraction":
                    extracted_id = extract_id(image, id_keyword)
                    with open(os.path.join(result_dir, f"{file}_id.txt"), 'w') as f:
                        f.write(f"Extracted ID: {extracted_id}\n")

                elif operation == "full_text_ocr":
                    text = extract_text(image)
                    with open(os.path.join(result_dir, f"{file}_ocr.txt"), 'w', encoding='utf-8') as f:
                        f.write(text)



# --- GUI --- #
class PDFProcessorApp:
    def __init__(self, root):
        self.root = root
        root.title("Batch PDF Processor")
        root.geometry("800x500")       # Set initial size: width x height
        root.minsize(700, 400)   
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure root grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        
        # Create main container with padding
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Batch PDF Processor", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # Input section frame
        input_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="15")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        input_frame.columnconfigure(1, weight=1)
        
        # Folder selection
        self.folder_path = tk.StringVar()
        ttk.Label(input_frame, text="PDF Folder:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='e', padx=(0, 10), pady=(0, 10))
        
        folder_entry = ttk.Entry(input_frame, textvariable=self.folder_path, width=50, 
                                font=('Arial', 10))
        folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=(0, 10))
        
        browse_btn = ttk.Button(input_frame, text="Browse", command=self.browse_folder)
        browse_btn.grid(row=0, column=2, padx=(10, 0), pady=(0, 10))
        
        # Keyword
        ttk.Label(input_frame, text="Keyword:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='e', padx=(0, 10), pady=(0, 10))
        self.keyword_entry = ttk.Entry(input_frame, width=50, font=('Arial', 10))
        self.keyword_entry.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # ID Keyword
        ttk.Label(input_frame, text="ID Keyword:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, sticky='e', padx=(0, 10), pady=(0, 10))
        self.id_keyword_entry = ttk.Entry(input_frame, width=50, font=('Arial', 10))
        self.id_keyword_entry.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Operation dropdown
        ttk.Label(input_frame, text="Operation:", font=('Arial', 10, 'bold')).grid(
            row=3, column=0, sticky='e', padx=(0, 10), pady=(0, 10))
        self.operation = tk.StringVar(value="keyword_search")
        operation_combo = ttk.Combobox(input_frame, textvariable=self.operation, 
                                      values=["keyword_search", "id_extraction", "full_text_ocr"], 
                                      width=47, font=('Arial', 10), state="readonly")
        operation_combo.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Action section frame
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=2, column=0, columnspan=3, pady=(10, 20))
        
        # Run button with custom styling
        run_btn = tk.Button(action_frame, text="▶ Run Process", command=self.run_process, 
                           bg="#4CAF50", fg="white", font=('Arial', 12, 'bold'),
                           relief=tk.RAISED, bd=2, padx=30, pady=10,
                           activebackground="#45a049", activeforeground="white",
                           cursor="hand2")
        run_btn.pack()
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="15")
        status_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process PDFs...")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                font=('Arial', 10), foreground="#666666")
        status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Progress bar (initially hidden)
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Instructions text
        instructions_text = tk.Text(status_frame, height=6, wrap=tk.WORD, 
                                   font=('Arial', 9), bg="#f8f9fa", 
                                   relief=tk.FLAT, bd=1)
        instructions_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        instructions = """Instructions:
1. Select a folder containing PDF files using the Browse button
2. Enter a keyword to search for (optional, depending on operation)
3. Enter an ID keyword for extraction (optional, for id_extraction operation)
4. Choose the operation type from the dropdown menu
5. Click 'Run Process' to start processing

Results will be saved in the 'processed_output' folder."""
        
        instructions_text.insert(tk.END, instructions)
        instructions_text.config(state=tk.DISABLED)
        
        # Add scrollbar to instructions
        scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=instructions_text.yview)
        instructions_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
        
        # Bind hover effects to run button
        run_btn.bind("<Enter>", lambda e: run_btn.config(bg="#45a049"))
        run_btn.bind("<Leave>", lambda e: run_btn.config(bg="#4CAF50"))
        
        # Add tooltips
        self.create_tooltip(browse_btn, "Select folder containing PDF files")
        self.create_tooltip(operation_combo, "Choose the type of processing operation")
        self.create_tooltip(run_btn, "Start processing the selected PDFs")

    def create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = tk.Label(tooltip, text=text, background="#ffffe0", 
                           relief=tk.SOLID, borderwidth=1, font=('Arial', 8))
            label.pack()
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select PDF Folder")
        if folder:
            self.folder_path.set(folder)
            self.status_var.set(f"Selected folder: {os.path.basename(folder)}")

    def run_process(self):
        folder = self.folder_path.get().strip()
        keyword = self.keyword_entry.get().strip()
        id_keyword = self.id_keyword_entry.get().strip()
        operation = self.operation.get()

        if not os.path.isdir(folder):
            messagebox.showerror("Error", "Please select a valid folder containing PDFs.")
            return
        
        # Update status and show progress
        self.status_var.set("Processing PDFs... Please wait.")
        self.progress.start(10)
        self.root.update()
        
        try:
            # Note: These functions need to be defined elsewhere in your code
            split_pdfs(folder)
            process_pages(operation, keyword, id_keyword)
            
            # Stop progress and update status
            self.progress.stop()
            self.status_var.set("Processing completed successfully!")
            
            messagebox.showinfo("Process Complete", 
                              f"Operation '{operation}' completed successfully!\n\n"
                              f"Results saved in 'processed_output' folder.\n"
                              f"Folder: {folder}\n"
                              f"Operation: {operation}")
            
        except Exception as e:
            self.progress.stop()
            self.status_var.set("Error occurred during processing.")
            messagebox.showerror("Processing Error", 
                               f"An error occurred during processing:\n\n{str(e)}")

# --- Run the App --- #
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFProcessorApp(root)
    root.mainloop()
