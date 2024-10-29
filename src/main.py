import os
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import shutil  # For moving files
import io
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# Define keywords for classification
compliance_keywords = [
    'complain', 'grievance', 'mad', 'angry', 'upset', 'legal action', 'lawyer', 'frustration',
    'department of insurance', 'regulatory agency', 'lawsuit', 'deadline for resolution',
    'unauthorized', 'inappropriate', 'theft', 'forgery', 'fraud', 'media', 'better business bureau',
    'subpoena', 'attorney letterhead'
]
appeal_keywords = [
    'appeal', 'request for review', 'reconsider', 'decision', 'reopen case', 'case review', 'reassessment'
]

# Function to extract text from each page of a text-based PDF
def extract_text_from_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text() + "\n\n"
    doc.close()
    return text

# Function to perform OCR on rendered page images (for scanned PDFs)
def ocr_pdf_page(page):
    pix = page.get_pixmap()
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    text = pytesseract.image_to_string(img)
    return text

# Function to perform OCR on multi-page TIF files
def ocr_image_file(image_file):
    text = ""
    img = Image.open(image_file)

    for page in range(img.n_frames):
        img.seek(page)
        text += pytesseract.image_to_string(img) + "\n\n"
    return text

# Function to process PDF and TIF files
def process_file(file_path):
    file_ext = os.path.splitext(file_path)[1].lower()
    text = ""
    
    if file_ext == '.pdf':
        text = extract_text_from_pdf(file_path)
        if not text.strip():
            doc = fitz.open(file_path)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text += ocr_pdf_page(page) + "\n\n"
            doc.close()
    elif file_ext == '.tif':
        text = ocr_image_file(file_path)
    else:
        print(f"[ERROR] Unsupported file format: {file_ext}")
    
    return text

# Function to classify text based on keywords
def classify_text(text):
    text_lower = text.lower()
    
    if any(keyword in text_lower for keyword in compliance_keywords):
        return 'Compliance'
    elif any(keyword in text_lower for keyword in appeal_keywords):
        return 'Appeal'
    
    return 'Unable to Detect'

# Style and format the Excel file
def style_excel(filename):
    workbook = load_workbook(filename)
    sheet = workbook.active

    # Define styles
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    
    # Apply styles to headers
    for col in sheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
    
    # Apply border to all cells and adjust column widths
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Adjust column widths
    sheet.column_dimensions['A'].width = 10
    sheet.column_dimensions['B'].width = 30
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 20

    workbook.save(filename)
    workbook.close()

# Main script
def main():
    input_folder = 'input_files'
    output_folder = 'output'
    archive_folder = 'archive'
    results_file = 'classification_results.xlsx'

    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(archive_folder, exist_ok=True)

    results = []

    for idx, file_name in enumerate(os.listdir(input_folder), start=1):
        file_path = os.path.join(input_folder, file_name)
        text = process_file(file_path)
        classification = classify_text(text)
        
        output_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_output.txt")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"Classification: {classification}\n\n")
            f.write(f"Extracted Text:\n{text}")

        processed_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        results.append([idx, file_name, processed_time, classification])

        archive_path = os.path.join(archive_folder, file_name)
        shutil.move(file_path, archive_path)

        print(f"[INFO] Processed {file_name}, classified as {classification} and saved output to {output_file}")
        print(f"[INFO] Moved {file_name} to archive folder")

    if os.path.exists(results_file):
        existing_df = pd.read_excel(results_file, engine='openpyxl')
        new_df = pd.DataFrame(results, columns=['Serial No', 'File Name', 'File Processed Date and Time', 'Classification'])
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = pd.DataFrame(results, columns=['Serial No', 'File Name', 'File Processed Date and Time', 'Classification'])

    combined_df.to_excel(results_file, index=False, engine='openpyxl')
    style_excel(results_file)  # Apply styling after saving to Excel
    print(f"[INFO] Classification results saved to {results_file}")

if __name__ == "__main__":
    main()
