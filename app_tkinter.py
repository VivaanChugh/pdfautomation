# ============================================================================
# PDF AUTOMATION SYSTEM - COMPREHENSIVE EXPLANATION
# ============================================================================
# This system automates the extraction of case numbers and file numbers from legal PDF documents.
# It uses two different OCR engines: EasyOCR for complex documents and Tesseract for simpler ones.
# The system processes PDFs page by page, extracts text using OCR, finds specific patterns,
# renames files based on extracted information, and generates Excel reports.
# ============================================================================

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
import pandas as pd
import csv
from pathlib import Path

# ============================================================================
# GLOBAL VARIABLES AND CONFIGURATION
# ============================================================================
# This dictionary tracks the current processing status across all functions.
# It's used by the GUI to show progress and by error handling to identify issues.
# The values are updated in real-time as PDFs are processed.
CURRENT_PROCESSING = {
    "pdf": None,           # Name of the PDF currently being processed
    "page": None,          # Current page number being processed
    "total_pages": None    # Total number of pages in the current PDF
}

# ============================================================================
# RESOURCE PATH HANDLING
# ============================================================================
# This function handles file paths differently depending on whether the code is running
# in development mode or has been compiled into an executable with PyInstaller.
# When compiled, PyInstaller creates a temporary directory (_MEIPASS) where all
# the required files are stored. This function finds the correct path in both scenarios.
def resource_path(relative_path):
    try:
        # If running as compiled executable, use PyInstaller's temporary directory
        base_path = sys._MEIPASS
    except AttributeError:
        # If running in development, use the current directory
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ============================================================================
# TESSERACT OCR CONFIGURATION
# ============================================================================
# Tesseract is a traditional OCR engine that works well for simple, clear text.
# We set the path to the Tesseract executable so the system knows where to find it.
# This is essential for the pytesseract library to work properly.
pytesseract.pytesseract.tesseract_cmd = os.path.join(resource_path("Tesseract-OCR"), "tesseract.exe")

# ============================================================================
# LOGGING DIRECTORY SETUP
# ============================================================================
# Create a dedicated folder for log files in the user's AppData directory.
# This ensures logs are stored in a standard location and persist between sessions.
# The logs help track processing progress and troubleshoot any issues that arise.
APP_LOG_DIR = os.path.join(os.getenv("APPDATA"), "PDFSplitter", "logs")
os.makedirs(APP_LOG_DIR, exist_ok=True)

# ============================================================================
# LOG MAINTENANCE FUNCTION
# ============================================================================
# This function prevents log files from accumulating indefinitely by removing
# files older than 30 days. This keeps the system running efficiently and
# prevents disk space issues from old log files.
def clean_old_logs():
    for filename in os.listdir(APP_LOG_DIR):
        full_path = os.path.join(APP_LOG_DIR, filename)
        if os.path.isfile(full_path):
            created_time = datetime.fromtimestamp(os.path.getctime(full_path))
            if datetime.now() - created_time > timedelta(days=30):
                os.remove(full_path)

# ============================================================================
# SUCCESS LOGGING FUNCTION
# ============================================================================
# This function records successful operations in the log file, including:
# - When each page was processed
# - What ID was extracted (if any)
# - What the new filename was (if a file was created)
# This creates a complete audit trail of all processing activities.
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

# ============================================================================
# ERROR LOGGING FUNCTION
# ============================================================================
# This function records all errors that occur during processing, including:
# - What function was running when the error occurred
# - The specific error message
# - When the error happened
# This information is crucial for debugging and improving the system.
def log_exception(context, error, log_file_path):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] ERROR in {context}:\n{error}\n\n")

# ============================================================================
# EXCEL REPORT GENERATION FOR SPLITTER FUNCTION
# ============================================================================
# This function creates Excel reports specifically for the basic PDF splitting function.
# It takes the extracted data and formats it into a professional Excel spreadsheet
# with columns for case numbers, timestamps, and file paths.
# The timestamp in the filename ensures each report is unique.
d

# ============================================================================
# EXCEL REPORT GENERATION FOR GENERAL EXTRACTION
# ============================================================================
# This function creates Excel reports for the advanced extraction functions that
# extract both case numbers and dates. It creates a more comprehensive report
# that includes all the extracted information in an organized format.
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

# ============================================================================
# EASYOCR INITIALIZATION
# ============================================================================
# EasyOCR is a deep learning-based OCR engine that provides superior accuracy
# for complex documents with varying fonts, layouts, and image quality.
# We initialize it with English language support and enable GPU acceleration
# if available. GPU acceleration significantly speeds up processing.
easyocr_reader = easyocr.Reader(['en'], gpu=torch.cuda.is_available())

# ============================================================================
# IMAGE PREPROCESSING FUNCTION
# ============================================================================
# This function enhances image quality before OCR processing to improve accuracy.
# It converts images to grayscale (which often improves OCR results),
# automatically adjusts contrast to make text more readable,
# and increases sharpness to make text edges clearer.
# These enhancements are particularly important for Tesseract OCR.
def preprocess_image(image):
    image = image.convert("L")  # Convert to grayscale
    image = ImageOps.autocontrast(image)  # Auto-adjust contrast
    image = ImageEnhance.Sharpness(image).enhance(2.0)  # Increase sharpness
    return image

# ============================================================================
# FILE NUMBER EXTRACTION FUNCTION (USING EASYOCR)
# ============================================================================
# This function extracts file numbers from dismissal notices using EasyOCR.
# It's designed for documents that have clear "File No:" labels.
# The function includes critical image resizing to prevent memory issues.
# 
# WHY RESIZING IS CRITICAL:
# - EasyOCR uses deep learning models that require significant memory
# - Large images (350 DPI) can cause GPU memory overflow
# - Resizing to 50% reduces memory usage by approximately 75%
# - Without resizing, EasyOCR fails and returns None, causing downstream errors
def extract_id_dismissal(image):
    try:
        # CRITICAL: Resize image to prevent memory issues with EasyOCR
        # This line is essential for system stability and reliability
        image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
        
        # Convert PIL image to numpy array format required by EasyOCR
        np_image = np.array(image.convert("RGB"))
        
        # Use EasyOCR to extract text from the image
        # detail=0 means we only want the text, not bounding boxes
        results = easyocr_reader.readtext(np_image, detail=0)
        
        # Combine all extracted text lines into a single string for pattern matching
        text = "\n".join(results)
        
        # Use regular expression to find file numbers
        # Pattern looks for "File No:", "File No.", "File No;" etc.
        # followed by alphanumeric characters, commas, periods, and hyphens
        matches = re.findall(r'(?:File\s*No[:.;]?\s*)([A-Za-z0-9.,\-]+)', text, re.IGNORECASE)

        if matches:
            # Clean the extracted ID by removing commas and periods
            # This preserves the ID structure while removing formatting artifacts
            clean_id = re.sub(r'[.,]', '', matches[0])
            return clean_id
        return None
    except Exception as e:
        # Log any errors that occur during extraction
        log_exception("extract_id_dismissal", e, log_file_path=None)
        return None

# ============================================================================
# CASE NUMBER EXTRACTION FUNCTION (USING TESSERACT)
# ============================================================================
# This function extracts case numbers from lien documents using Tesseract OCR.
# It's designed for documents that have "case no" or "caseno" labels.
# Tesseract works well for simpler documents and doesn't require image resizing.
# 
# WHY TESSERACT HERE:
# - Lien documents typically have simpler, clearer text
# - Tesseract is faster and uses less memory than EasyOCR
# - It's more reliable for consistent document formats
def extract_id_lien(image):
    try:
        # Apply image preprocessing to improve OCR accuracy
        image = preprocess_image(image)
        
        # Use Tesseract OCR to extract text from the image
        text = pytesseract.image_to_string(image)
        
        # Split the extracted text into individual lines for processing
        lines = text.splitlines()

        for line in lines:
            line_lower = line.lower()  # Convert to lowercase for case-insensitive matching
            
            # Check for "case no" pattern (with space between words)
            if "case no" in line_lower:
                idx = line_lower.find("case no")  # Find the position of "case no"
                after = line[idx + len("case no"):].strip(" .:_-")  # Get text after "case no"
                
                # Use regex to extract alphanumeric characters and spaces
                # This captures the complete case number even if it contains spaces
                match = re.match(r'^([A-Za-z0-9\s]+)', after)
                if match:
                    # Remove all spaces from the matched ID to create a clean identifier
                    cleaned = re.sub(r'\s+', '', match.group(1))
                    if cleaned:  # Only return if we have a valid, non-empty ID
                        return cleaned
            
            # Check for "caseno" pattern (without space) as a fallback
            # Some documents might use this format instead
            elif "caseno" in line_lower:
                idx = line_lower.find("caseno")  # Find the position of "caseno"
                after = line[idx + len("caseno"):].strip(" .:_-")  # Get text after "caseno"
                
                # Same regex pattern as above for consistency
                match = re.match(r'^([A-Za-z0-9\s]+)', after)
                if match:
                    cleaned = re.sub(r'\s+', '', match.group(1))
                    if cleaned:
                        return cleaned
        
        return None  # Return None if no case number was found
        
    except Exception as e:
        # Log any errors that occur during extraction
        log_exception("extract_id_lien", e, log_file_path=None)
        return None

# ============================================================================
# CASE NUMBER EXTRACTION FUNCTION FOR JUDGMENTS (USING TESSERACT)
# ============================================================================
# This function extracts case numbers from judgment documents using Tesseract OCR.
# It's designed for documents that have "case number" labels.
# The function processes each line to find the specific pattern.
def extract_id_judgement(image):
    try:
        # Apply image preprocessing to improve OCR accuracy
        image = preprocess_image(image)
        
        # Use Tesseract OCR to extract text from the image
        text = pytesseract.image_to_string(image)
        
        # Split the extracted text into individual lines for processing
        lines = text.splitlines()

        for line in lines:
            line_lower = line.lower()  # Convert to lowercase for case-insensitive matching
            
            # Check for "case number" pattern
            if "case number" in line_lower:
                idx = line_lower.find("case number")  # Find the position of "case number"
                after = line[idx + len("case number"):].strip(" .:_-")  # Get text after "case number"
                
                # Remove all spaces from the matched text to create a clean identifier
                match = after.replace(" ","")
                return match
            
            # Note: This function could be enhanced with additional fallback patterns
        
        return None  # Return None if no case number was found
        
    except Exception as e:
        # Log any errors that occur during extraction
        log_exception("extract_id_judgement", e, log_file_path=None)
        return None

# ============================================================================
# EXTRACT MD JUDGEMENTS CAVA
# ============================================================================
# This function extracts case number and date for MD Judgements CAVA
def extract_md_judgements_cava(image):
    """Extract case number and date for MD Judgements CAVA"""
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        
        case_number = None
        date_found = None
        
        for line in lines:
            line_lower = line.lower()
            
            if case_number is None:
                if "case number" in line_lower:
                    idx = line_lower.find("case number")
                    after = line[idx + len("case number"):].strip(" .:_-")
                    case_number = after.replace(" ","")
                elif "case no" in line_lower:
                    idx = line_lower.find("case no")
                    after = line[idx + len("case no"):].strip(" .:_-")
                    case_number = after.replace(" ","")
                            
            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None and "on" in line_lower:
                idx = line_lower.find("on")
                after = line[idx + len("on"):].strip(" .:-")

                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_md_judgements_cava", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT VA JUDGEMENTS LVNV
# ============================================================================
# This function extracts case number and date for VA Judgements LVNV
def extract_va_judgements_lvnv(image):
    """Extract case number and date for VA Judgements LVNV"""
    try:
        image = image.resize((image.width // 2, image.height //2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail = 0)
        text = "\n".join(results)
        lines = text.splitlines()
        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number in None:
    
                # Check for "case number" pattern
                if ("case" in line_lower) and ("further case" not in line_lower) and ("case warrant" not in line_lower) and ("case information" not in line_lower) and ("case details" not in line_lower) and ("case number" not in line_lower):
                    idx = line_lower.find("case")
                    after = line[idx + len("case"):].strip(" .:_-")
                    match = re.match(r'^([A-Za-z0-9\s]+)', after)
                    if match:
                        case_number = re.sub(r'\s+', '', match.group(1))

            
            
            
            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None:
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        if case_number is None:
            id = lines.index("Case")

            case_number = lines[id+1]
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_va_judgements_lvnv", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT VA JUDGEMENTS CAVA
# ============================================================================
# This function extracts case number and date for VA Judgements CAVA
def extract_va_judgements_cava(image):
    try:
        image = image.resize((image.width // 2, image.height //2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail = 0)
        text = "\n".join(results)
        lines = text.splitlines()
        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number in None:
    
                # Check for "case number" pattern
                if ("case" in line_lower) and ("further case" not in line_lower) and ("case warrant" not in line_lower) and ("case information" not in line_lower) and ("case details" not in line_lower) and ("case number" not in line_lower):
                    idx = line_lower.find("case")
                    after = line[idx + len("case"):].strip(" .:_-")
                    match = re.match(r'^([A-Za-z0-9\s]+)', after)
                    if match:
                        case_number = re.sub(r'\s+', '', match.group(1))

            
            

            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None:
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        if case_number is None:
            id = lines.index("Case")

            case_number = lines[id+1]
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_va_judgements_cava", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT JUDGEMENTS MCM
# ============================================================================
# This function extracts case number and date for Judgements MCM
def extract_judgements_mcm(image):
    try:
        image = image.resize((image.width // 2, image.height //2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail = 0)
        text = "\n".join(results)
        lines = text.splitlines()
        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number in None:
    
                # Check for "case number" pattern
                if ("case" in line_lower) and ("further case" not in line_lower) and ("case warrant" not in line_lower) and ("case information" not in line_lower) and ("case details" not in line_lower) and ("case number" not in line_lower):
                    idx = line_lower.find("case")
                    after = line[idx + len("case"):].strip(" .:_-")
                    match = re.match(r'^([A-Za-z0-9\s]+)', after)
                    if match:
                        case_number = re.sub(r'\s+', '', match.group(1))

            
            
            
            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None:
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        if case_number is None:
            id = lines.index("Case")

            case_number = lines[id+1]
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_judgements_mcm", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT ORDER OF SATISFACTION
# ============================================================================
# This function extracts FileNo for Order of Satisfaction
def extract_order_satisfaction(image):
    """Extract FileNo for Order of Satisfaction"""
    try:
        image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail=0)
        text = "\n".join(results)
        matches = re.findall(r'(?:File\s*No[:.;]?\s*)([A-Za-z0-9.,\-]+)', text, re.IGNORECASE)

        if matches:
            # Remove commas, periods, and all spaces from the entire ID
            clean_id = re.sub(r'[.,\s]', '', matches[0])
            return clean_id
        return None
    except Exception as e:
        log_exception("extract_order_satisfaction", e, log_file_path=None)
        return None

# ============================================================================
# EXTRACT UPDATE DISMISSAL RESURGENT CAVALRY
# ============================================================================
# This function extracts case number and date for Update Dismissal Resurgent Cavalry
def extract_update_dismissal_resurgent_cavalry(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        

        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number is None:
                # Check for "case number" pattern
                
                if "number" in line_lower:
                    idx = line_lower.find("case number")
                    after = line[idx + len("number"):].strip(" .:_")
                    case_number = after
                
            
            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None and "on" in line_lower:
                idx = line_lower.find("on")
                after = line[idx + len("on"):].strip(" .:-")
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break

        if case_number is None:
            idx = line_lower.find("number:")
            after = line[idx + len("number:"):].strip(" .:_")
            case_number = after
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_update_dismissal_resurgent_cavalry", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT UPDATE LIEN CAC/CAVALRY
# ============================================================================
# This function extracts case number and date for Update Lien CAC/Cavalry
def extract_update_lien_cac_cavalry(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        

        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number is None:

                # Check for "case number" pattern
                if "number" in line_lower:
                    idx = line_lower.find("number")
                    after = line[idx + len("number"):].strip(" .:_")
                    case_number = after

            
            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None and "on" in line_lower:
                idx = line_lower.find("on")
                after = line[idx + len("on"):].strip(" .:-")
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        
        if case_number is None:
            idx = line_lower.find("number:")
            after = line[idx + len("number:"):].strip(" .:_")
            case_number = after
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_update_lien_cac_cavalry", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT UPDATE SERVICE MD GARNS
# ============================================================================
# This function extracts case number and date for Update Service MD Garns
def extract_update_service_md_garns(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        
        
        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number is None:

                # Check for "case number" pattern
                if "number" in line_lower:
                    idx = line_lower.find("number:")
                    after = line[idx + len("number:"):].strip(" .:_-")
                    case_number = after


            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None and "on" in line_lower:
                idx = line_lower.find("on")
                after = line[idx + len("on"):].strip(" .:-")
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        
        if case_number is None:
            idx = line_lower.find("number:")
            after = line[idx + len("number:"):].strip(" .:_")
            case_number = after
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_update_service_md_garns", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT MD LVNV
# ============================================================================
# This function extracts case number and date for MD LVNV
def extract_md_lvnv(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        
        
        case_number = None
        date_found = None
        for line in lines:
            line_lower = line.lower()
            if case_number is None:

                # Check for "case number" pattern
                if "number" in line_lower:
                    idx = line_lower.find("number:")
                    after = line[idx + len("number:"):].strip(" .:_-")
                    case_number = after


            # Look for date patterns (dd/mm/yyyy or dd-mm-yyyy)
            if date_found is None and "on" in line_lower:
                idx = line_lower.find("on")
                after = line[idx + len("on"):].strip(" .:-")
                # Look for dd/mm/yyyy or dd-mm-yyyy patterns
                date_patterns = [
                    r'\b(\d{1,2}/\d{1,2}/\d{4})\b',  # dd/mm/yyyy
                    r'\b(\d{1,2}-\d{1,2}-\d{4})\b',  # dd-mm-yyyy
                    r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b',  # dd.mm.yyyy
                    r'\b(\d{4}-\d{1,2}-\d{1,2})\b',  # yyyy-mm-dd
                    r'\b(\d{4}/\d{1,2}/\d{1,2})\b',  # yyyy/mm/dd
                ]
                for pattern in date_patterns:
                    match = re.search(pattern, line)
                    if match:
                        date_found = match.group(1)
                        break
        
        if case_number is None:
            idx = line_lower.find("number:")
            after = line[idx + len("number:"):].strip(" .:_")
            case_number = after
        return case_number, date_found
        
    except Exception as e:
        log_exception("extract_md_lvnv", e, log_file_path=None)
        return None, None

# ============================================================================
# EXTRACT LIEN REQ
# ============================================================================
# This function extracts case number for Lien Req
def extract_lien_req(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()
        case_number = None
        pattern = re.compile(r'\bC\d{7}\b')
        for line in lines:
            line_lower = line.lower()
            if case_number is None:

                
                match = pattern.search(line)
                if match:
                    case_number = match.group()
                    break

        return case_number

    except Exception as e:
        log_exception("extract_lien_req", e , log_file_path=None)
        return None

# ============================================================================
# EXTRACT BUS REC
# ============================================================================
# This function extracts case number for Business Records
def extract_bus_rec(image):
    try:
        image = image.resize((image.width // 2, image.height // 2), Image.Resampling.LANCZOS)
        np_image = np.array(image.convert("RGB"))
        results = easyocr_reader.readtext(np_image, detail = 0)
        text = "\n".join(results)
        lines = text.splitlines()
        case_number = None
        pattern = re.compile(r'\b[CR].{7}\b', re.IGNORECASE)
        for line in lines:
            line_lower = line.lower()
            if case_number is None:


                match = pattern.search(line)
                if match:
                    case_number = match.group()
                    if case_number.lower() == "court of" or case_number.lower() == "records ":
                        case_number = None
                        continue
                    else:
                        break

        return case_number

    except Exception as e:
        log_exception("extract_bus_rec", e, log_file_path=None)
        return None

# ============================================================================
# EXTRACT EFILE STIP FOLDER
# ============================================================================
# This function extracts case number and notice for Efile Stipulations
def extract_efile_stip_folder(image):
    try:
        image = preprocess_image(image)
        text = pytesseract.image_to_string(image)
        lines = text.splitlines()

        case_number = None
        notice = None
        for line in lines:
            line_lower = line.lower()
            if "file no." in line_lower:

                idx = line_lower.find("file no.")
                after = line[idx + len("file no."):].strip(" .:-_")

                parts = after.split()
                if parts:
                    case_number = parts[0]
                    break
        for line in lines:
            line_lower = line.lower()
            if "stipulation" in line_lower:
                notice = "Stipulation"
                break
            elif "judgment" in line_lower:
                notice = "Judgment By Consent"
                break

        
        return case_number, notice
    
    except Exception as e:
        log_exception("extract_efile_strip_folder", e, log_file_path=None)
        return None, None
    
# ============================================================================
# UNIQUE FILENAME GENERATOR
# ============================================================================
# This function ensures that no two files have the same name in the output directory.
# It automatically adds "_copy1", "_copy2", etc. to filenames if duplicates exist.
# This prevents data loss from file overwriting and maintains data integrity.
# 
# WHY THIS IS IMPORTANT:
# - Multiple PDFs might have the same case number
# - Different pages from the same PDF might extract the same ID
# - Without this, files would overwrite each other, losing data
# - Legal documents require complete preservation of all information
def get_unique_filename(base_path, base_name, extension=".pdf"):
    # Start with the original filename
    filename = f"{base_name}{extension}"
    counter = 1
    
    # Keep checking if the filename exists, and if so, add a counter
    while os.path.exists(os.path.join(base_path, filename)):
        # Add "_copy1", "_copy2", etc. to make the filename unique
        filename = f"{base_name}_copy{counter}{extension}"
        counter += 1
    
    # Return the full path to the unique filename
    return os.path.join(base_path, filename)

# ============================================================================
# MAIN PDF PROCESSING FUNCTION - CORE OF THE SYSTEM
# ============================================================================
# This is the heart of the PDF automation system. It processes individual PDF files
# and extracts case numbers or file numbers based on the selected document type.
# 
# HOW IT WORKS STEP BY STEP:
# 1. Opens the PDF and processes each page individually for maximum flexibility
# 2. Converts each page to a high-resolution image (350 DPI) for optimal OCR accuracy
# 3. Uses the appropriate OCR engine based on document complexity:
#    - EasyOCR for complex dismissal notices (better accuracy, requires resizing)
#    - Tesseract for simple lien documents (faster, more reliable)
# 4. Applies pattern matching to find case numbers or file numbers in extracted text
# 5. Renames and saves individual pages with extracted information
# 6. Creates comprehensive logs for audit trails and troubleshooting
# 7. Tracks progress for real-time GUI updates
# 8. Returns organized data for Excel report generation
# 
# WHY THIS APPROACH IS SUPERIOR:
# - Page-by-page processing allows individual file naming and organization
# - High-resolution images (350 DPI) ensure OCR accuracy even with poor quality documents
# - Different OCR engines are optimized for different document types
# - Comprehensive logging creates audit trails for legal compliance
# - Memory management prevents crashes during large batch processing
# - Progress tracking provides user feedback during long operations
def process_pdf(pdf_path, output_base, id_keyword, progress_callback, index, total_files, log_file_path, process_start_time):
    data_records = []  # Master list to store all extracted data for Excel reporting
    
    try:
        # STEP 1: SETUP AND ORGANIZATION
        # Extract the PDF filename without extension for clean folder naming
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        # Create a dedicated output directory for this specific PDF
        # This keeps all extracted pages organized by source document
        # Example: If processing "Case123.pdf", creates folder "Case123/"
        output_dir = os.path.join(output_base, pdf_name)
        os.makedirs(output_dir, exist_ok=True)

        # STEP 2: PDF ANALYSIS
        # Open and read the PDF file to determine total page count
        # This information is used for progress tracking and user feedback
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)

        # STEP 3: PAGE-BY-PAGE PROCESSING
        # Process each page individually for maximum flexibility and error isolation
        for i, page in enumerate(reader.pages):
            # Update global processing status for real-time progress tracking
            # This information is displayed in the GUI to show current activity
            CURRENT_PROCESSING["pdf"] = pdf_name
            CURRENT_PROCESSING["page"] = i + 1
            CURRENT_PROCESSING["total_pages"] = total_pages

            try:
                # STEP 4: PAGE EXTRACTION
                # Create a new PDF writer for this single page
                # This allows us to save each page as a separate, named file
                writer = PdfWriter()
                writer.add_page(page)
                
                # Create a temporary file path for this individual page
                # The double underscore prefix indicates temporary files that will be deleted
                temp_path = os.path.join(output_dir, f"__temp_page_{i+1}.pdf")
                
                # Save the single page as a temporary PDF file
                # This temporary file is needed for the PDF-to-image conversion process
                with open(temp_path, 'wb') as f:
                    writer.write(f)

                # STEP 5: IMAGE CONVERSION
                # Convert the PDF page to a high-resolution image for OCR processing
                # 350 DPI provides excellent text clarity for accurate OCR results
                # Poppler is used for PDF-to-image conversion (more reliable than alternatives)
                # The [0] index gets the first (and only) page from the conversion result
                image = convert_from_path(temp_path, dpi=350, poppler_path=resource_path("poppler-bin"))[0]
                
                # Clean up: Remove the temporary PDF file to save disk space
                # We only need the image for OCR processing, not the temporary PDF
                os.remove(temp_path)

                # STEP 6: OCR ENGINE SELECTION
                # Choose the appropriate extraction method based on the document type
                # Different document types require different OCR approaches and patterns
                if "fileno" in id_keyword.lower():
                    # For dismissal notices, use EasyOCR (better for complex layouts)
                    # EasyOCR handles varying fonts, layouts, and image quality better
                    extracted_id = extract_id_dismissal(image)
                    notice_label = "Notice Of Dismissal"
                elif "case number" in id_keyword.lower():
                    # For judgment documents, use Tesseract (faster for simple text)
                    # Tesseract is more efficient for straightforward document formats
                    extracted_id = extract_id_judgement(image)
                    notice_label = ""
                else:
                    # For lien documents, use Tesseract (most reliable for this type)
                    # Lien documents typically have consistent, clear formatting
                    extracted_id = extract_id_lien(image)

                # STEP 7: FILE CREATION AND NAMING
                # Initialize final_path to prevent None value errors
                # This is crucial for preventing crashes when OCR extraction fails
                final_path = None
                
                if extracted_id:
                    # If an ID was successfully extracted, create a new filename
                    # The filename combines the extracted ID with a descriptive label
                    if "fileno" in id_keyword.lower():
                        base_filename = f"{extracted_id}_{notice_label}"
                    elif "case number" in id_keyword.lower():
                        base_filename = f"{extracted_id}_{notice_label}"
                    else:
                        base_filename = f"{extracted_id}"
                    
                    # Get a unique filename (adds _copy1, _copy2, etc. if duplicates exist)
                    # This prevents overwriting existing files and maintains data integrity
                    final_path = get_unique_filename(output_dir, base_filename)
                    
                    # Save the individual page with the new filename
                    # This creates a separate PDF file for each page with meaningful names
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    
                    # Log the successful extraction for audit purposes
                    # This creates a complete record of what was processed and when
                    log_text(pdf_name, i + 1, extracted_id, log_file_path, final_path)
                else:
                    # Log that no ID was found on this page
                    # This helps identify pages that need manual review or different processing
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # STEP 8: METADATA COLLECTION
                # Get the creation timestamp of the newly created file
                # This information is included in the Excel report for tracking purposes
                pdf_modified_date = ""
                if extracted_id and final_path:
                    # Only get the timestamp if both ID and file were successfully created
                    # This prevents errors when trying to access non-existent files
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # STEP 9: DATA RECORDING
                # Add this page's data to the master record list
                # This data will be used to generate the comprehensive Excel report
                data_records.append([
                    extracted_id if extracted_id else "",  # ID (blank if none found)
                    process_start_time,                    # When processing started
                    pdf_modified_date,                     # When new file was created
                    pdf_path                               # Original PDF path for reference
                ])

            except Exception as e:
                # STEP 10: ERROR HANDLING
                # Log any errors that occur while processing this specific page
                # This allows for page-level error handling without stopping the entire process
                # Users can see exactly which pages had issues and why
                log_exception("process_pdf", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            # STEP 11: MEMORY MANAGEMENT (CRITICAL FOR STABILITY)
            # Force garbage collection to free up memory after each page
            # This prevents memory buildup during large batch processing
            gc.collect()
            
            # If using GPU acceleration, clear the GPU memory cache
            # This prevents GPU memory overflow during batch processing
            # Without this, the system would crash after processing several large PDFs
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            # STEP 12: PROGRESS TRACKING
            # Calculate and update progress percentage for the GUI
            # Progress accounts for both current file and overall batch progress
            # This gives users accurate feedback on processing status
            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        # Clear the current processing status when finished
        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        # STEP 13: GLOBAL ERROR HANDLING
        # Log any errors that occur while processing the entire PDF
        # This catches errors that happen outside the page processing loop
        # Examples: PDF corruption, permission issues, disk space problems
        log_exception("process_pdf", e, log_file_path)
    
    return data_records  # Return all extracted data for Excel report generation

# ============================================================================
# PROCESS MD JUDGEMENTS CAVA
# ============================================================================
# This function processes MD Judgements CAVA PDFs and extracts data based on the selected keyword.
def process_md_judgements_cava(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_md_judgements_cava(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_md_judgements_cava", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_md_judgements_cava", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS VA JUDGEMENTS LVNV
# ============================================================================
# This function processes VA Judgements LVNV PDFs and extracts data based on the selected keyword.
def process_va_judgements_lvnv(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_va_judgements_lvnv(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])
            except Exception as e:
                error_msg = f"file-level error in {pdf_name} page {i+1}:\n{str(e)}"
                log_exception("process_va_judgements_lvnv", error_msg, log_file_path)
                
            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_va_judgements_lvnv", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS VA JUDGEMENTS CAVA
# ============================================================================
# This function processes VA Judgements CAVA PDFs and extracts data based on the selected keyword.
def process_va_judgements_cava(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_va_judgements_cava(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_va_judgements_cava", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_va_judgements_cava", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS JUDGEMENTS MCM
# ============================================================================
# This function processes Judgements MCM PDFs and extracts data based on the selected keyword.
def process_judgements_mcm(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_judgements_mcm(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_judgements_mcm", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_judgements_mcm", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS ORDER OF SATISFACTION
# ============================================================================
# This function extracts FileNo for Order of Satisfaction
def process_order_satisfaction(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract FileNo
                file_number = extract_order_satisfaction(image)

                if file_number:
                    base_filename = f"{file_number}_Order_of_Satisfaction"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, file_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if file_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    file_number if file_number else "",  # File Number
                    "",  # Date Found (not used for this type)
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_order_satisfaction", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_order_satisfaction", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS UPDATE DISMISSAL RESURGENT CAVALRY
# ============================================================================
# This function extracts case number and date for Update Dismissal Resurgent Cavalry
def process_update_dismissal_resurgent_cavalry(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                case_number, date_found = extract_update_dismissal_resurgent_cavalry(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_update_dismissal_resurgent_cavalry", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_update_dismissal_resurgent_cavalry", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS UPDATE LIEN CAC/CAVALRY
# ============================================================================
# This function extracts case number and date for Update Lien CAC/Cavalry
def process_update_lien_cac_cavalry(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_update_lien_cac_cavalry(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_update_lien_cac_cavalry", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_update_lien_cac_cavalry", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS UPDATE SERVICE MD GARNS
# ============================================================================
# This function extracts case number and date for Update Service MD Garns
def process_update_service_md_garns(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract both case number and date
                case_number, date_found = extract_update_service_md_garns(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # Case Number
                    date_found if date_found else "",    # Date Found
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_update_service_md_garns", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_update_service_md_garns", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS MD LVNV
# ============================================================================
# This function extracts case number and date for MD LVNV
def process_md_lvnv(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Extract FileNo
                case_number, date_found = extract_md_lvnv(image)

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank values if none found)
                data_records.append([
                    case_number if case_number else "",  # File Number
                    date_found if date_found else "",  # Date Found (not used for this type)
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_update_md_lvnv", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_md_lvnv", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS LIEN REQ
# ============================================================================
# This function extracts case number for Lien Req
def process_lien_req(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Use the dismissal extraction logic (FileNo extraction)
                case_number = extract_lien_req(image)
                date_found = None

                if case_number:
                    base_filename = f"{case_number}"
                    final_path = get_unique_filename(output_dir, base_filename)  # Default fallback
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank ID if none found)
                data_records.append([
                    case_number if case_number else "",  # File Number
                    date_found if date_found else "",  # Date Found (not used for this type)
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_lien_req", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_lien_req", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS BUS REC
# ============================================================================
# This function extracts case number for Business Records
def process_bus_rec(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Use the dismissal extraction logic (FileNo extraction)
                case_number = extract_bus_rec(image)
                date_found = None
                if case_number:
                    base_filename = f"{case_number}_Business Records"
                    final_path = get_unique_filename(output_dir, base_filename)  # Default fallback
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank ID if none found)
                data_records.append([
                    case_number if case_number else "",  # File Number
                    date_found if date_found else "",  # Date Found (not used for this type)
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_bus_rec", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_bus_rec", e, log_file_path)
    
    return data_records

# ============================================================================
# PROCESS EFILE STIP FOLDER
# ============================================================================
# This function extracts case number and notice for Efile Stipulations
def process_efile_stip_folder(pdf_path, output_base, progress_callback, index, total_files, log_file_path, process_start_time):
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

                # Use the dismissal extraction logic (FileNo extraction)
                case_number, notice = extract_efile_stip_folder(image)
                date_found = None

                if case_number:
                    base_filename = f"{case_number}_{notice}"
                    final_path = get_unique_filename(output_dir, base_filename)  # Default fallback
                    with open(final_path, 'wb') as out_f:
                        writer.write(out_f)
                    log_text(pdf_name, i + 1, case_number, log_file_path, final_path)
                else:
                    log_text(pdf_name, i + 1, None, log_file_path)
                
                # Get the creation date of the new page file (if it was created)
                pdf_modified_date = ""
                if case_number:
                    # Get the creation date of the newly created individual page file
                    pdf_modified_date = datetime.fromtimestamp(os.path.getctime(final_path)).strftime("%Y-%m-%d %H:%M:%S")
                
                # Add record to data_records (with blank ID if none found)
                data_records.append([
                    case_number if case_number else "",  # File Number
                    date_found if date_found else "",  # Date Found (not used for this type)
                    process_start_time,                  # Current Datestamp
                    pdf_modified_date,                   # PDF Modified Date
                    pdf_path                             # Source Path
                ])

            except Exception as e:
                log_exception("process_efile_stip_folder", f"file-level error in {pdf_name}:\n{e}", log_file_path)

            gc.collect()
            if torch.cuda.is_available():
                torch.cuda.empty_cache()

            progress = ((index + (i + 1) / total_pages) / total_files) * 100
            progress_callback(progress)

        CURRENT_PROCESSING["pdf"] = None

    except Exception as e:
        log_exception("process_efile_stip_folder", e, log_file_path)
    
    return data_records

# ============================================================================
# MAIN GUI APPLICATION CLASS - PDF UTILITY SUITE
# ============================================================================
# This is the main user interface for the PDF automation system. It provides
# a comprehensive suite of tools for processing legal documents including:
# 
# CORE FEATURES:
# 1. SPLITTER: Extracts individual pages from PDFs and renames them based on OCR results
# 2. MERGER: Combines multiple PDFs into single documents
# 3. REDACTION: Removes sensitive information from documents
# 4. COMPRESSOR: Reduces file sizes while maintaining quality
# 
# WHY THIS INTERFACE DESIGN:
# - Single application handles all PDF processing needs
# - Tabbed interface keeps different functions organized
# - Progress tracking shows real-time processing status
# - Comprehensive logging for audit trails and troubleshooting
# - Professional appearance suitable for legal office environments
# 
# TECHNICAL ARCHITECTURE:
# - Built with Tkinter for cross-platform compatibility
# - Multi-threaded processing prevents GUI freezing
# - Real-time progress updates during long operations
# - Error handling with user-friendly messages
# - Configuration persistence between sessions
class SplitPDFApp:
    def __init__(self, root):
        # ============================================================================
        # APPLICATION INITIALIZATION AND SETUP
        # ============================================================================
        # This method sets up the entire application interface, including:
        # - Window configuration and sizing
        # - Variable initialization for folder paths
        # - Logo and branding elements
        # - Main navigation buttons
        # - Content area setup for different functions
        # 
        # WHY THESE CHOICES:
        # - Large window (1800x900) accommodates complex legal document workflows
        # - StringVar() variables provide reactive GUI updates
        # - Professional logo establishes credibility for legal office use
        # - Tabbed interface keeps functions organized and accessible
        
        self.root = root
        root.title("PDF Utility Suite - Splitter, Merger, Redaction, Compressor")
        root.geometry("1800x900")  # Large window for complex workflows
        
        # Processing state flag to prevent multiple operations simultaneously
        self.processing = False
        
        # ============================================================================
        # FOLDER PATH VARIABLES FOR DIFFERENT DOCUMENT TYPES
        # ============================================================================
        # Each document type gets its own output folder to maintain organization
        # These variables store the user's folder selections and persist between sessions
        # StringVar() provides automatic GUI updates when values change
        
        # Basic document types
        self.dismissal_folder = tk.StringVar()      # Dismissal notices
        self.lien_folder = tk.StringVar()           # Lien documents
        self.judgement_folder = tk.StringVar()      # Judgment documents
        
        # Specialized document types for different jurisdictions
        self.md_judgements_cava_folder = tk.StringVar()      # MD judgments (CAVA)
        self.va_judgements_lvnv_folder = tk.StringVar()      # VA judgments (LVNV)
        self.va_judgements_cava_folder = tk.StringVar()      # VA judgments (CAVA)
        self.judgements_mcm_folder = tk.StringVar()          # MCM judgments
        
        # Order and update document types
        self.order_satisfaction_folder = tk.StringVar()      # Satisfaction orders
        self.update_dismissal_resurgent_cavalry_folder = tk.StringVar()  # Updated dismissals
        self.update_lien_cac_cavalry_folder = tk.StringVar()            # Updated liens
        self.update_service_md_garns_folder = tk.StringVar()            # Service updates
        
        # Additional document types
        self.lien_req_folder = tk.StringVar()       # Lien requests
        self.bus_rec_folder = tk.StringVar()        # Business records
        self.efile_stip_folder = tk.StringVar()     # E-filed stipulations
        self.upload_md_lvnv = tk.StringVar()        # MD LVNV uploads
        
        # Logging and tracking
        self.latest_log_file = None  # Tracks the most recent log file for error reporting


        # ============================================================================
        # USER INTERFACE COMPONENT SETUP
        # ============================================================================
        # This section creates the visual elements of the application:
        # - Professional logo for branding and credibility
        # - Navigation buttons for different functions
        # - Content areas for each major feature
        # - Default view selection
        
        # --- Top Logo Section ---
        # The logo establishes professional credibility for legal office use
        # It's loaded from the resource path to work in both development and compiled versions
        logo_path = resource_path("logo.png")
        self.logo_frame = tk.Frame(root)
        self.logo_frame.pack(pady=(10, 5))  # Add spacing above and below logo
        
        if os.path.exists(logo_path):
            # Load and display the logo image
            logo_image = Image.open(logo_path)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(self.logo_frame, image=logo_photo)
            logo_label.image = logo_photo  # Keep reference to prevent garbage collection
            logo_label.pack()

        # --- Main Navigation Button Bar ---
        # These buttons provide access to the four main functions of the application
        # Each button is sized and styled for easy use in professional environments
        self.button_frame = tk.Frame(root)
        self.button_frame.pack(pady=(0, 10))  # Add spacing below buttons
        
        self.feature_buttons = {}  # Dictionary to store button references for highlighting
        
        # Define the four main application features
        features = [
            ("Splitter", self.show_splitter),      # Main OCR-based page splitting
            ("Merger", self.show_merger),          # PDF combination tool
            ("Redaction", self.show_redaction),    # Sensitive data removal
            ("Compressor", self.show_compressor)   # File size reduction
        ]
        
        # Create and position each navigation button
        for i, (name, cmd) in enumerate(features):
            btn = tk.Button(
                self.button_frame, 
                text=name, 
                font=("Arial", 13, "bold"),  # Professional font styling
                width=18, height=2,          # Consistent button sizing
                command=cmd                  # Function to call when clicked
            )
            btn.grid(row=0, column=i, padx=10)  # Grid layout with spacing
            self.feature_buttons[name] = btn     # Store reference for highlighting

        # --- Main Content Area Setup ---
        # This frame contains all the functional content and expands to fill available space
        self.content_frame = tk.Frame(root)
        self.content_frame.pack(fill='both', expand=True)

        # Create individual content frames for each major function
        # Each frame will contain the specific interface elements for that function
        self.splitter_tab = tk.Frame(self.content_frame)      # OCR page splitting interface
        self.merger_tab = tk.Frame(self.content_frame)        # PDF merging interface
        self.redaction_tab = tk.Frame(self.content_frame)     # Redaction interface
        self.compressor_tab = tk.Frame(self.content_frame)    # Compression interface

        # Initialize each tab's interface components
        # This sets up all the buttons, fields, and controls for each function
        self.init_splitter_tab()      # Sets up OCR splitting interface
        self.init_merger_tab()        # Sets up PDF merging interface
        self.init_redaction_tab()     # Sets up redaction interface
        self.init_compressor_tab()    # Sets up compression interface

        # Set the default view to the Splitter function
        # This is the most commonly used feature for legal document processing
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
        # Main frame
        
        
        frame = tk.Frame(self.splitter_tab)
        frame.pack(pady=20)

        title_label = tk.Label(frame, text="PDF Document Processor", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))

        selection_frame = tk.Frame(frame)
        selection_frame.pack(pady=10)

        tk.Label(selection_frame, text="Select Document Type:", font=("Arial", 12)).pack(side=tk.LEFT, padx=(0, 10))
        
        # Create dropdown options with their corresponding browse functions
        self.document_options = [
            ("Dismissal PDFs", self.browse_dismissal, self.dismissal_folder, "progress_dismissal"),
            ("Lien PDFs", self.browse_lien, self.lien_folder, "progress_lien"),
            ("Judgement Satisfied PDFs", self.browse_judgement, self.judgement_folder, "progress_judgement"),
            ("Order of Satisfaction PDFs", self.browse_order_satisfaction, self.order_satisfaction_folder, "progress_order_satisfaction"),
            ("MD Judgements CAVA", self.browse_md_judgements_cava, self.md_judgements_cava_folder, "progress_md_judgements_cava"),
            ("VA Judgements LVNV", self.browse_va_judgements_lvnv, self.va_judgements_lvnv_folder, "progress_va_judgements_lvnv"),
            ("VA Judgements CAVA", self.browse_va_judgements_cava, self.va_judgements_cava_folder, "progress_va_judgements_cava"),
            ("Judgements MCM", self.browse_judgements_mcm, self.judgements_mcm_folder, "progress_judgements_mcm"),
            ("Update Dismissal Resurgent/Cavalry", self.browse_update_dismissal_resurgent_cavalry, self.update_dismissal_resurgent_cavalry_folder, "progress_update_dismissal_resurgent_cavalry"),
            ("Update Lien CAC/Cavalry", self.browse_update_lien_cac_cavalry, self.update_lien_cac_cavalry_folder, "progress_update_lien_cac_cavalry"),
            ("Update Service MD Garns", self.browse_update_service_md_garns, self.update_service_md_garns_folder, "progress_update_service_md_garns"),
            ("MD LVNV", self.browse_upload_md_lvnv, self.upload_md_lvnv, "progress_upload_md_lvnv"),
            ("Lien Req", self.browse_lien_req, self.lien_req_folder, "progress_lien_req"),
            ("Efile Stipulations", self.browse_efile_stip_folder, self.efile_stip_folder, "progress_efile_stip_folder"),
            ("Business Records", self.browse_bus_rec, self.bus_rec_folder, "progress_bus_rec")
        
        
        
        ]
        
        self.selected_document_type = tk.StringVar()
        self.document_dropdown = ttk.Combobox(selection_frame, textvariable=self.selected_document_type, 
                                             values=[doc[0] for doc in self.document_options], 
                                             state="readonly", width=40, font=("Arial", 11))
        self.document_dropdown.pack(side=tk.LEFT, padx=(0, 20))
        self.document_dropdown.set("Select a Document Type")

        # Browse button
        self.browse_button = tk.Button(selection_frame, text="Browse Folder", 
                                      command=self.browse_selected_document,
                                      font=("Arial", 12, "bold"), 
                                      bg="#1976d2", fg="white", 
                                      width=15, height=2)
        
        self.browse_button.pack(side=tk.LEFT)

        folder_frame = tk.Frame(frame)
        folder_frame.pack(pady=20)

        tk.Label(folder_frame, text="Selected Folder:", font=("Arial", 12)).pack(anchor=tk.W)
        
        self.folder_path_var = tk.StringVar(value="No folder selected")
        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path_var, width=60, 
                               font=("Arial", 10), state="readonly")
        
        folder_entry.pack(pady=(5, 10), fill=tk.X)



        self.progress_dismissal = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_lien = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_judgement = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_order_satisfaction = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_md_judgements_cava = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_va_judgements_lvnv = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_va_judgements_cava = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_judgements_mcm = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_update_dismissal_resurgent_cavalry = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_update_lien_cac_cavalry = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_update_service_md_garns = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_upload_md_lvnv = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_lien_req = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_efile_stip_folder = ttk.Progressbar(frame, length=180, mode="determinate")
        self.progress_bus_rec = ttk.Progressbar(frame, length=180, mode="determinate")
        
        

        self.hide_all_progress_bars()

    def hide_all_progress_bars(self):
        progress_bars = [self.progress_dismissal, self.progress_lien, self.progress_judgement, self.progress_order_satisfaction, self.progress_md_judgements_cava, self.progress_va_judgements_lvnv, self.progress_va_judgements_cava, self.progress_judgements_mcm, 
                         self.progress_update_dismissal_resurgent_cavalry, self.progress_update_lien_cac_cavalry,
                         self.progress_update_service_md_garns, self.progress_upload_md_lvnv, self.progress_lien_req, self.progress_efile_stip_folder, self.progress_bus_rec]
        for bar in progress_bars:
            bar.pack_forget()

    def show_progress_bar(self, progress_bar_name):
        self.hide_all_progress_bars()
        progress_bar = getattr(self,progress_bar_name)
        progress_bar.pack(pady=(0,10))

    def browse_selected_document(self):
        selected_text = self. selected_document_type.get()
        if selected_text == "Select a Document Type" or not selected_text:
            messagebox.showwarning("Selection Required", "Please select a document type first.")
            return
        
        selected_option = None
        for doc in self.document_options:
            if doc[0] == selected_text:
                selected_option = doc
                break
            
        if not selected_option:
            messagebox.showerror("Error", "Invalid document type selection.")
            return
        
        progress_bar_name = selected_option[3]
        self.show_progress_bar(progress_bar_name)

        browse_function = selected_option[1]
        browse_function()


        folder_var = selected_option[2]
        if folder_var.get():
            self.folder_path_var.set(folder_var.get())
        
    def init_merger_tab(self):
        label = tk.Label(self.merger_tab, text="PDF Merger", font=("Arial", 14))
        label.pack(pady=10)
        
        self.remove_permissions_folder = tk.StringVar()
        self.copies_output_folder = None  # Path to last _copies folder
        self.merger_files_var = tk.StringVar(value=[])
        self.copied_files_var = tk.StringVar(value=[])
        
        # --- Merge Section Layout: Horizontal Buttons ---
        merge_outer_frame = tk.Frame(self.merger_tab)
        merge_outer_frame.pack(pady=5)
        
        # Remove Permissions Section (Left)
        remove_frame = tk.Frame(merge_outer_frame)
        remove_frame.grid(row=0, column=0, padx=10, sticky="n")
        tk.Label(remove_frame, text="Step 1: Remove Permissions from PDFs").pack()
        tk.Button(remove_frame, text="Select Folder and Remove Permissions", command=self.remove_permissions_from_pdfs).pack(pady=2)
        self.remove_permissions_label = tk.Label(remove_frame, textvariable=self.remove_permissions_folder, fg="gray")
        self.remove_permissions_label.pack(pady=2)
        self.copied_files_listbox = tk.Listbox(remove_frame, listvariable=self.copied_files_var, width=70, height=4)
        self.copied_files_listbox.pack(pady=2)
        
        # Merge Section (Right)
        merge_frame = tk.Frame(merge_outer_frame)
        merge_frame.grid(row=0, column=1, padx=10, sticky="n")
        tk.Label(merge_frame, text="Step 2: Merge Cleaned PDFs").pack()
        self.merge_folder_label = tk.Label(merge_frame, text="No cleaned folder yet", fg="gray")
        self.merge_folder_label.pack(pady=2)
        self.files_listbox = tk.Listbox(merge_frame, listvariable=self.merger_files_var, width=70, height=4)
        self.files_listbox.pack(pady=2)
        self.merge_btn = tk.Button(merge_frame, text="Merge All PDFs in Cleaned Folder", command=self.merge_all_pdfs_in_folder, state="disabled")
        self.merge_btn.pack(pady=2)

    def remove_permissions_from_pdfs(self):
        folder = filedialog.askdirectory(title="Select Folder to Remove Permissions from PDFs")
        if not folder:
            return
        self.remove_permissions_folder.set(folder)
        output_folder = folder.rstrip("/\\") + "_copies"
        self.copies_output_folder = output_folder
        os.makedirs(output_folder, exist_ok=True)
        copied_files = []
        for root, dirs, files in os.walk(folder):
            rel = os.path.relpath(root, folder)
            out_subfolder = os.path.join(output_folder, rel) if rel != '.' else output_folder
            os.makedirs(out_subfolder, exist_ok=True)
            for f in files:
                if f.lower().endswith('.pdf'):
                    in_path = os.path.join(root, f)
                    out_path = os.path.join(out_subfolder, f)
                    try:
                        reader = PdfReader(in_path)
                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        with open(out_path, "wb") as out_f:
                            writer.write(out_f)
                        copied_files.append(out_path)
                    except Exception as e:
                        copied_files.append(f"ERROR: {in_path}")
        self.copied_files_var.set([f"Copied: {os.path.relpath(f, output_folder)}" if not f.startswith("ERROR") else f for f in copied_files])
        # Update merge section
        self.merge_folder_label.config(text=f"Will merge: {output_folder}", fg="black")
        self.merge_btn.config(state="normal")
        # List PDFs in output folder
        pdf_files = []
        for root, dirs, files in os.walk(output_folder):
            for f in files:
                if f.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, f))
        display = [f"{len(pdf_files)} PDFs found in cleaned folder"]
        self.merger_files_var.set(display)
        self.merger_pdf_files = pdf_files
        messagebox.showinfo("Done", f"Copied {len(copied_files)} PDFs to {output_folder}.")

    def merge_all_pdfs_in_folder(self):
        folder = self.copies_output_folder
        if not folder or not hasattr(self, 'merger_pdf_files') or not self.merger_pdf_files:
            messagebox.showerror("No PDFs Found", "Please run Step 1 to create cleaned PDFs before merging.")
            return
        log_file_path = os.path.join(APP_LOG_DIR, f"merger_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        self.latest_log_file = log_file_path
        try:
            merger = PdfWriter()
            for pdf_file in self.merger_pdf_files:
                try:
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        merger.add_page(page)
                except Exception as e:
                    log_exception("merge_all_pdfs_in_folder", f"Failed to read {pdf_file}: {e}", log_file_path)
                    continue
            output_path = os.path.join(folder, f"{os.path.basename(folder)}.pdf")
            with open(output_path, "wb") as f_out:
                merger.write(f_out)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] Merged PDF files in {folder} and all subfolders:\n")
                for file in self.merger_pdf_files:
                    f.write(f"  - {file}\n")
                f.write(f"Saved merged PDF as: {output_path}\n\n")
            messagebox.showinfo("Success", f"Merged {len(self.merger_pdf_files)} PDFs into {output_path}.")
        except Exception as e:
            log_exception("merge_all_pdfs_in_folder", e, log_file_path)
            messagebox.showerror("Error", f"Failed to merge PDFs:\n{e}")


    def init_redaction_tab(self):
        label = tk.Label(self.redaction_tab, text="PDF Redaction", font=("Arial", 14))
        label.pack(pady=20)


    def init_compressor_tab(self):
        label = tk.Label(self.compressor_tab, text="PDF Compressor", font=("Arial", 14))
        label.pack(pady=10)
       
        self.compress_input_folder = None
        self.compress_original_size_var = tk.StringVar(value="Original Size: N/A")
        self.compress_compressed_size_var = tk.StringVar(value="Compressed Size: N/A")
       
        select_folder_btn = tk.Button(self.compressor_tab, text="Select Folder to Compress All PDFs", command=self.select_compress_folder)
        select_folder_btn.pack(pady=5)
       
        self.compress_file_label = tk.Label(self.compressor_tab, text="No folder selected", fg="gray")
        self.compress_file_label.pack(pady=2)
       
        tk.Label(self.compressor_tab, textvariable=self.compress_original_size_var).pack(pady=2)
        tk.Label(self.compressor_tab, textvariable=self.compress_compressed_size_var).pack(pady=2)
       
        compress_btn = tk.Button(self.compressor_tab, text="Compress PDF(s)", command=self.compress_pdf)
        compress_btn.pack(pady=10)

    def select_compress_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.compress_input_folder = folder
            self.compress_file_label.config(text=f"Folder: {os.path.basename(folder)}", fg="black")
            self.compress_original_size_var.set("Original Size: N/A")
            self.compress_compressed_size_var.set("Compressed Size: N/A")

    def compress_pdf(self):
        if self.compress_input_folder:
            self._compress_folder_pdfs(self.compress_input_folder)
        else:
            messagebox.showerror("No Folder Selected", "Please select a folder to compress.")

    def _compress_folder_pdfs(self, input_folder):
        output_folder = input_folder.rstrip("/\\") + "_compressed"
        os.makedirs(output_folder, exist_ok=True)
        log_file_path = os.path.join(APP_LOG_DIR, f"compressor_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
        self.latest_log_file = log_file_path
        pdfs_to_compress = []
        # PDFs in all subfolders recursively
        for root, dirs, files in os.walk(input_folder):
            rel = os.path.relpath(root, input_folder)
            out_subfolder = os.path.join(output_folder, rel) if rel != '.' else output_folder
            os.makedirs(out_subfolder, exist_ok=True)
            for f in files:
                if f.lower().endswith('.pdf'):
                    in_path = os.path.join(root, f)
                    # Use get_unique_filename for output
                    base_name = os.path.splitext(f)[0]
                    out_path = get_unique_filename(out_subfolder, base_name)
                    pdfs_to_compress.append((in_path, out_path))
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


    def browse_judgement(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.judgement_folder.set(path)
            self.run_type(path, "judgement", "Case Number", self.progress_judgement)





    def browse_md_judgements_cava(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.md_judgements_cava_folder.set(path)
            self.run_type_md_judgements_cava(path, "md_judgements_cava", self.progress_md_judgements_cava)


    def browse_va_judgements_lvnv(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.va_judgements_lvnv_folder.set(path)
            self.run_type_va_judgements_lvnv(path, "va_judgements_lvnv", self.progress_va_judgements_lvnv)


    def browse_va_judgements_cava(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.va_judgements_cava_folder.set(path)
            self.run_type_va_judgements_cava(path, "va_judgements_cava", self.progress_va_judgements_cava)


    def browse_judgements_mcm(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.judgements_mcm_folder.set(path)
            self.run_type_judgements_mcm(path, "judgements_mcm", self.progress_judgements_mcm)


    def browse_order_satisfaction(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.order_satisfaction_folder.set(path)
            self.run_type_order_satisfaction(path, "order_satisfaction", self.progress_order_satisfaction)


    def browse_update_dismissal_resurgent_cavalry(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.update_dismissal_resurgent_cavalry_folder.set(path)
            self.run_type_update_dismissal_resurgent_cavalry(path, "update_dismissal_resurgent_cavalry", self.progress_update_dismissal_resurgent_cavalry)


    def browse_update_lien_cac_cavalry(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.update_lien_cac_cavalry_folder.set(path)
            self.run_type_update_lien_cac_cavalry(path, "update_lien_cac_cavalry", self.progress_update_lien_cac_cavalry)


    def browse_update_service_md_garns(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.update_service_md_garns_folder.set(path)
            self.run_type_update_service_md_garns(path, "update_service_md_garns", self.progress_update_service_md_garns)


    def browse_upload_md_lvnv(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.upload_md_lvnv.set(path)
            self.run_type_md_lvnv(path, "upload_md_lvnv", self.progress_upload_md_lvnv)

    def browse_lien_req(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.lien_req_folder.set(path)
            self.run_type_lien_req(path, "lien_req_folder", self.progress_lien_req)


    def browse_bus_rec(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.bus_rec_folder.set(path)
            self.run_type_bus_rec(path, "bus_rec_folder", self.progress_bus_rec)

    def browse_efile_stip_folder(self):
        if self.processing:
            messagebox.showwarning("Wait", "A process is already running.")
            return
        path = filedialog.askdirectory()
        if path:
            self.efile_stip_folder.set(path)
            self.run_type_efile_stip_folder(path, "efile_stip_folder", self.progress_efile_stip_folder)

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
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_pdf(path, folder, id_keyword, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} {keyword_match} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_md_judgements_cava(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_md_judgements_cava(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_md_judgements_cava", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_va_judgements_lvnv(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_va_judgements_lvnv(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_va_judgements_lvnv", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_va_judgements_cava(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_va_judgements_cava(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_va_judgements_cava", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_judgements_mcm(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_judgements_mcm(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_judgements_mcm", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_order_satisfaction(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_order_satisfaction(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_order_satisfaction", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_update_dismissal_resurgent_cavalry(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_update_dismissal_resurgent_cavalry(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_update_dismissal_resurgent_cavalry", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_update_lien_cac_cavalry(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_update_lien_cac_cavalry(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_update_lien_cac_cavalry", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_update_service_md_garns(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_update_service_md_garns(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_update_service_md_garns", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_upload_md_lvnv(self, folder, keyword_match, progressbar):

        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_md_lvnv(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_upload_md_lvnv", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()

    def run_type_lien_req(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_lien_req(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_lien_req", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()


    def run_type_bus_rec(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_bus_rec(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_bus_rec", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
            finally:
                self.processing = False

        threading.Thread(target=worker, daemon=True).start()

    def run_type_efile_stip_folder(self, folder, keyword_match, progressbar):
        def update_progress(val):
            if self.root.winfo_exists():
                self.root.after(0, lambda: progressbar.config(value=val))

        def worker():
            self.processing = True
            try:
                if not os.path.isdir(folder):
                    messagebox.showerror("Error", "Invalid folder path.")
                    return

                # No filename restrictions
                pdfs = [os.path.join(folder, f) for f in os.listdir(folder)
                        if f.lower().endswith('.pdf')]
                
                if not pdfs:
                    messagebox.showerror("Error", "No PDFs found in the selected folder.")
                    return

                log_file_path = os.path.join(APP_LOG_DIR, f"{keyword_match}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_log.txt")
                self.latest_log_file = log_file_path

                progressbar["value"] = 0
                total_files = len(pdfs)
                all_data_records = []

                process_start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                for idx, path in enumerate(pdfs):
                    data_records = process_efile_stip_folder(path, folder, update_progress, idx, total_files, log_file_path, process_start_time)
                    if data_records:
                        all_data_records.extend(data_records)

                try:
                    excel_path = create_general_report(all_data_records, APP_LOG_DIR, keyword_match)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nExcel report: {os.path.basename(excel_path)}")
                except Exception as e:
                    log_exception("create_general_report", e, log_file_path)
                    messagebox.showinfo("Done", f"Processed {total_files} PDF(s).\n\nError creating report: {str(e)}")

                progressbar["value"] = 0

            except Exception as e:
                log_exception("run_type_efile_stip_folder", e, self.latest_log_file or os.path.join(APP_LOG_DIR, "error_fallback.log"))
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



