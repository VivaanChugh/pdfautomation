import os
import re
import threading
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException, UploadFile, File, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageOps
import pytesseract
import easyocr
import numpy as np
import torch
import gc
from datetime import datetime, timedelta
import sys
import io
import subprocess
import pandas as pd
import csv
from pathlib import Path
import uvicorn
import shutil
import tempfile
import zipfile
from fastapi import Form


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


def create_splitter_report(data_records, output_folder, keyword_match):
    """Create only an Excel report for the splitter function"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    
    # Create DataFrame
    df = pd.DataFrame(data_records, columns=[
        'CaseNo/FileNo', 
        'Current Datestamp', 
        'PDF Modified Date', 
        'Source Path'
    ])
    
    # Save as Excel only
    excel_path = os.path.join(output_folder, f"{keyword_match}_splitter_report_{timestamp}.xlsx")
    try:
        df.to_excel(excel_path, index=False, engine='openpyxl')
    except Exception as e:
        raise RuntimeError(f"Failed to create Excel report: {e}")
    
    return excel_path


def create_general_report(data_records, output_folder, keyword_match):
    """Create Excel report for general extraction with case number and date"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    
    # Create DataFrame
    df = pd.DataFrame(data_records, columns=[
        'Case Number', 
        'Date Found',
        'Current Datestamp', 
        'PDF Modified Date', 
        'Source Path'
    ])
    
    # Save as Excel only
    excel_path = os.path.join(output_folder, f"{keyword_match}_general_report_{timestamp}.xlsx")
    try:
        df.to_excel(excel_path, index=False, engine='openpyxl')
    except Exception as e:
        raise RuntimeError(f"Failed to create Excel report: {e}")
    
    return excel_path


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
            # Remove commas and periods, but keep the ID structure
            clean_id = re.sub(r'[.,]', '', matches[0])
            return clean_id
        return None
    except Exception as e:
        log_exception("extract_id_dismissal", e, log_file_path=None)
        return None


def get_unique_filename(base_path, base_name, extension=".pdf"):
    filename = f"{base_name}{extension}"
    counter = 1
    while os.path.exists(os.path.join(base_path, filename)):
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    return os.path.join(base_path, filename)


def process_pdf(pdf_path, output_base, id_keyword, index, total_files, log_file_path, process_start_time):
    data_records = []
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

                image = convert_from_path(temp_path, dpi=350, poppler_path=resource_path("poppler-bin"))[0]
                os.remove(temp_path)

                if "fileno" in id_keyword.lower():
                    extracted_id = extract_id_dismissal(image)
                    notice_label = "Notice Of Dismissal"
                

                if extracted_id:
                    if "fileno" in id_keyword.lower():
                        base_filename = f"{extracted_id}_{notice_label}"
                    elif "case number" in id_keyword.lower():
                        base_filename = f"{extracted_id}_{notice_label}"
                    else:
                        base_filename = f"{extracted_id}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, extracted_id, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if extracted_id:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank ID if none found)
                data_records.append([
                    extracted_id if extracted_id else "",  # Blank if no ID found
                    process_start_time,
                    pdf_modified_date,
                    pdf_path
                ])

            except Exception as e:
                log_exception("process_pdf", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_pdf", e, log_file_path)
    
    return data_records


# Pydantic models for API requests/responses
class ProcessingStatus(BaseModel):
    status: str
    message: str
    progress: Optional[float] = None
    current_pdf: Optional[str] = None
    current_page: Optional[int] = None
    total_pages: Optional[int] = None


class SplitterRequest(BaseModel):
    folder_path: str
    document_type: str
    id_keyword: str


class MergerRequest(BaseModel):
    folder_path: str


class CompressorRequest(BaseModel):
    folder_path: str


class ProcessingResult(BaseModel):
    success: bool
    message: str
    output_folder: Optional[str] = None
    report_path: Optional[str] = None
    processed_files: Optional[int] = None


# Initialize FastAPI app
app = FastAPI(
    title="PDF Utility Suite API",
    description="API for PDF processing operations including splitting, merging, redaction, and compression",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global processing state
processing_state = {
    "is_processing": False,
    "current_operation": None,
    "progress": 0.0,
    "current_pdf": None,
    "current_page": None,
    "total_pages": None
}


@app.on_event("startup")
async def startup_event():
    """Initialize the application on startup"""
    clean_old_logs()


@app.get("/")
async def root():
    """Root endpoint with API information"""
    return {
        "message": "PDF Utility Suite API",
        "version": "1.0.0",
        "endpoints": {
            "splitter": "/splitter",
            "merger": "/merger", 
            "compressor": "/compressor",
            "status": "/status"
        }
    }


@app.get("/status")
async def get_status():
    """Get current processing status"""
    return ProcessingStatus(
        status="processing" if processing_state["is_processing"] else "idle",
        message=processing_state.get("message", "No operation in progress"),
        progress=processing_state.get("progress", 0.0),
        current_pdf=processing_state.get("current_pdf"),
        current_page=processing_state.get("current_page"),
        total_pages=processing_state.get("total_pages")
    )


@app.post("/splitter", response_model=ProcessingResult)
async def process_pdfs_splitter(request: SplitterRequest, background_tasks: BackgroundTasks):
    """Process PDFs for splitting based on document type"""
    if processing_state["is_processing"]:
        raise HTTPException(status_code=400, detail="Another operation is already in progress")
    
    if not os.path.isdir(request.folder_path):
        raise HTTPException(status_code=400, detail="Invalid folder path")
    
    # Add background task for processing
    background_tasks.add_task(
        process_pdfs_background,
        request.folder_path,
        request.document_type,
        request.id_keyword
    )
    
    return ProcessingResult(
        success=True,
        message=f"Started processing {request.document_type} PDFs from {request.folder_path}",
        output_folder=request.folder_path
    )


@app.post("/merger", response_model=ProcessingResult)
async def merge_pdfs(request: MergerRequest, background_tasks: BackgroundTasks):
    """Merge PDFs from a folder"""
    if processing_state["is_processing"]:
        raise HTTPException(status_code=400, detail="Another operation is already in progress")
    
    if not os.path.isdir(request.folder_path):
        raise HTTPException(status_code=400, detail="Invalid folder path")
    
    # Add background task for merging
    background_tasks.add_task(merge_pdfs_background, request.folder_path)
    
    return ProcessingResult(
        success=True,
        message=f"Started merging PDFs from {request.folder_path}",
        output_folder=request.folder_path
    )


@app.post("/compressor", response_model=ProcessingResult)
async def compress_pdfs(request: CompressorRequest, background_tasks: BackgroundTasks):
    """Compress PDFs from a folder"""
    if processing_state["is_processing"]:
        raise HTTPException(status_code=400, detail="Another operation is already in progress")
    
    if not os.path.isdir(request.folder_path):
        raise HTTPException(status_code=400, detail="Invalid folder path")
    
    # Add background task for compression
    background_tasks.add_task(compress_pdfs_background, request.folder_path)
    
    return ProcessingResult(
        success=True,
        message=f"Started compressing PDFs from {request.folder_path}",
        output_folder=request.folder_path
    )


@app.post("/upload-pdfs")
async def upload_pdfs(files: List[UploadFile] = File(...)):
    """Upload PDF files for processing"""
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    
    # Create temporary directory for uploaded files
    temp_dir = tempfile.mkdtemp()
    uploaded_files = []
    
    try:
        for file in files:
            if not file.filename.lower().endswith('.pdf'):
                continue
            
            file_path = os.path.join(temp_dir, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            uploaded_files.append(file_path)
        
        if not uploaded_files:
            raise HTTPException(status_code=400, detail="No valid PDF files uploaded")
        
        return {
            "success": True,
            "message": f"Uploaded {len(uploaded_files)} PDF files",
            "temp_directory": temp_dir,
            "files": [os.path.basename(f) for f in uploaded_files]
        }
    
    except Exception as e:
        # Clean up on error
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")


@app.get("/download-report/{report_name}")
async def download_report(report_name: str):
    """Download a generated report"""
    report_path = os.path.join(APP_LOG_DIR, report_name)
    
    if not os.path.exists(report_path):
        raise HTTPException(status_code=404, detail="Report not found")
    
    return FileResponse(
        path=report_path,
        filename=report_name,
        media_type="application/octet-stream"
    )


@app.get("/download-output/{folder_name}")
async def download_output(folder_name: str):
    """Download processed output as a ZIP file"""
    output_path = os.path.join(APP_LOG_DIR, folder_name)
    
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="Output folder not found")
    
    # Create ZIP file
    zip_path = f"{output_path}.zip"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, output_path)
                zipf.write(file_path, arcname)
    
    return FileResponse(
        path=zip_path,
        filename=f"{folder_name}.zip",
        media_type="application/zip"
    )


# Background processing functions
def process_pdfs_background(folder_path: str, document_type: str, id_keyword: str):
    """Background task for processing PDFs"""
    global processing_state
    
    try:
        processing_state["is_processing"] = True
        processing_state["current_operation"] = f"Processing {document_type} PDFs"
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None
        
        # Find PDFs in folder
        pdfs = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                if f.lower().endswith('.pdf') and document_type.lower() in f.lower()]
        
        if not pdfs:
            processing_state["is_processing"] = False
            processing_state["message"] = f"No {document_type} PDFs found in folder"
            return
        
        log_file_path = os.path.join(APP_LOG_DIR, f"{document_type}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        total_files = len(pdfs)
        all_data_records = []
        
        for idx, pdf_path in enumerate(pdfs):
            processing_state["current_pdf"] = os.path.basename(pdf_path)
            processing_state["progress"] = (idx / total_files) * 100
            processing_state["message"] = f"Processing {os.path.basename(pdf_path)} ({idx + 1}/{total_files})"
            
            data_records = process_pdf(
                pdf_path, folder_path, id_keyword, idx, total_files, 
                log_file_path, process_start_time
            )
            
            if data_records:
                all_data_records.extend(data_records)
        
        # Create report
        try:
            excel_path = create_splitter_report(all_data_records, APP_LOG_DIR, document_type)
            processing_state["message"] = f"Processed {total_files} PDFs. Report: {os.path.basename(excel_path)}"
        except Exception as e:
            processing_state["message"] = f"Processed {total_files} PDFs. Error creating report: {str(e)}"
    
    except Exception as e:
        processing_state["message"] = f"Processing failed: {str(e)}"
    finally:
        processing_state["is_processing"] = False
        processing_state["current_operation"] = None
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None


def merge_pdfs_background(folder_path: str):
    """Background task for merging PDFs"""
    global processing_state
    
    try:
        processing_state["is_processing"] = True
        processing_state["current_operation"] = "Merging PDFs"
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None
        
        # Find all PDFs in folder and subfolders
        pdf_files = []
        for root, dirs, files in os.walk(folder_path):
            for f in files:
                if f.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, f))
        
        if not pdf_files:
            processing_state["is_processing"] = False
            processing_state["message"] = "No PDFs found in folder"
            return
        
        log_file_path = os.path.join(APP_LOG_DIR, f"merger_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        
        # Create output folder
        output_folder = folder_path.rstrip("/\\") + "_merged"
        os.makedirs(output_folder, exist_ok=True)
        
        # Merge PDFs
        merger = PdfWriter()
        for i, pdf_file in enumerate(pdf_files):
            try:
                processing_state["current_pdf"] = os.path.basename(pdf_file)
                processing_state["progress"] = ((i + 1) / len(pdf_files)) * 100
                processing_state["message"] = f"Merging {os.path.basename(pdf_file)} ({i + 1}/{len(pdf_files)})"
                
                reader = PdfReader(pdf_file)
                for page in reader.pages:
                    merger.add_page(page)
                
            except Exception as e:
                log_exception("merge_pdfs_background", f"Failed to read {pdf_file}: {e}", log_file_path)
                continue
        
        # Save merged PDF
        output_path = os.path.join(output_folder, f"{os.path.basename(folder_path)}_merged.pdf")
        with open(output_path, "wb") as f_out:
            merger.write(f_out)
        
        # Log results
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] Merged {len(pdf_files)} PDF files:\n")
            for file in pdf_files:
                f.write(f"  - {file}\n")
            f.write(f"Saved merged PDF as: {output_path}\n\n")
        
        processing_state["message"] = f"Successfully merged {len(pdf_files)} PDFs into {output_path}"
    
    except Exception as e:
        processing_state["message"] = f"Merging failed: {str(e)}"
    finally:
        processing_state["is_processing"] = False
        processing_state["current_operation"] = None
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None


def compress_pdfs_background(folder_path: str):
    """Background task for compressing PDFs"""
    global processing_state
    
    try:
        processing_state["is_processing"] = True
        processing_state["current_operation"] = "Compressing PDFs"
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None
        
        # Create output folder
        output_folder = folder_path.rstrip("/\\") + "_compressed"
        os.makedirs(output_folder, exist_ok=True)
        
        log_file_path = os.path.join(APP_LOG_DIR, f"compressor_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        
        # Find all PDFs
        pdfs_to_compress = []
        for root, dirs, files in os.walk(folder_path):
            rel = os.path.relpath(root, folder_path)
            out_subfolder = os.path.join(output_folder, rel) if rel != '.' else output_folder
            os.makedirs(out_subfolder, exist_ok=True)
            
            for f in files:
                if f.lower().endswith('.pdf'):
                    in_path = os.path.join(root, f)
                    base_name = os.path.splitext(f)[0]
                    out_path = get_unique_filename(out_subfolder, base_name)
                    pdfs_to_compress.append((in_path, out_path))
        
        # Compress PDFs
        count = 0
        for i, (in_path, out_path) in enumerate(pdfs_to_compress):
            try:
                processing_state["current_pdf"] = os.path.basename(in_path)
                processing_state["progress"] = ((i + 1) / len(pdfs_to_compress)) * 100
                processing_state["message"] = f"Compressing {os.path.basename(in_path)} ({i + 1}/{len(pdfs_to_compress)})"
                
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
                
                # Log compression
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(log_file_path, "a", encoding="utf-8") as f:
                    f.write(f"[{timestamp}] Compressed PDF file: {in_path}\n")
                    f.write(f"Saved compressed PDF as: {out_path}\n\n")
                
            except Exception as e:
                log_exception("compress_pdfs_background", e, log_file_path)
        
        processing_state["message"] = f"Successfully compressed {count} PDF(s). Output folder: {output_folder}"
    
    except Exception as e:
        processing_state["message"] = f"Compression failed: {str(e)}"
    finally:
        processing_state["is_processing"] = False
        processing_state["current_operation"] = None
        processing_state["progress"] = 0.0
        processing_state["current_pdf"] = None
        processing_state["current_page"] = None
        processing_state["total_pages"] = None


if __name__ == "__main__":
    clean_old_logs()
    uvicorn.run(app, host="0.0.0.0", port=8000)



