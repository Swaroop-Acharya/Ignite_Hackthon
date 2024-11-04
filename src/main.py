import os
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import shutil
import io
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# Define keywords for classification
compliant_keywords = [
    'complaint', 'grievance', 'mad', 'angry', 'upset', 'legal', 'action', 'legal action', 'lawyer', 'frustration', 'attorney',
    'department of insurance', 'doi', 'lawyer', 'regulatory agency', 'lawsuit', 'hire', 'deadline', 'action',
    'unauthorized', 'inappropriate', 'theft', 'forgery', 'fraud', 'media', 'better business bureau',
    'subpoena', 'attorney letterhead'
]

# Lock to prevent concurrent file processing issues
file_lock = Lock()

def extract_text_from_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text() + "\n\n"
    doc.close()
    return text

def ocr_pdf_page(page):
    pix = page.get_pixmap()
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    text = pytesseract.image_to_string(img)
    return text

def ocr_image_file(image_file):
    text = ""
    img = Image.open(image_file)
    for page in range(img.n_frames):
        img.seek(page)
        text += pytesseract.image_to_string(img) + "\n\n"
    return text

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

def classify_text(text):
    text_lower = text.lower()
    if any(keyword in text_lower for keyword in compliant_keywords):
        return 'Compliant'
    return 'Appeal'

def style_excel(filename):
    workbook = load_workbook(filename)
    sheet = workbook.active
    header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    
    for col in sheet.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
    
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    sheet.column_dimensions['A'].width = 10
    sheet.column_dimensions['B'].width = 30
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 20
    workbook.save(filename)
    workbook.close()

def process_single_file(idx, file_name, input_folder, output_folder, compliant_folder, appeal_folder, unable_to_detect_folder):
    file_path = os.path.join(input_folder, file_name)
    
    # Check if the file has already been archived to avoid duplicate processing
    with file_lock:
        if any(os.path.exists(os.path.join(folder, file_name)) for folder in [compliant_folder, appeal_folder, unable_to_detect_folder]):
            print(f"[INFO] Skipping already processed file: {file_name}")
            return None

    text = process_file(file_path)
    classification = classify_text(text)
    
    output_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_output.txt")
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(f"Classification: {classification}\n\n")
        f.write(f"Extracted Text:\n{text}")

    processed_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    result = [idx, file_name, processed_time, classification]

    # Determine the destination folder based on classification
    if classification == 'Compliant':
        destination_folder = compliant_folder
    elif classification == 'Appeal':
        destination_folder = appeal_folder
    else:
        destination_folder = unable_to_detect_folder

    # Move the processed file to the respective folder in the archive
    archive_path = os.path.join(destination_folder, file_name)
    shutil.move(file_path, archive_path)

    print(f"[INFO] Processed {file_name}, classified as {classification} and saved output to {output_file}")
    print(f"[INFO] Moved {file_name} to {destination_folder}")
    
    return result

def main():
    input_folder = 'input'
    output_folder = 'extracted_text'
    archive_folder = 'archive'
    results_file = 'classification_results.xlsx'

    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(archive_folder, exist_ok=True)

    compliant_folder = os.path.join(archive_folder, 'Compliant')
    appeal_folder = os.path.join(archive_folder, 'Appeal')
    unable_to_detect_folder = os.path.join(archive_folder, 'Unable to Detect')

    os.makedirs(compliant_folder, exist_ok=True)
    os.makedirs(appeal_folder, exist_ok=True)
    os.makedirs(unable_to_detect_folder, exist_ok=True)

    results = []

    # Use ThreadPoolExecutor to process files concurrently
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_file = {
            executor.submit(process_single_file, idx, file_name, input_folder, output_folder, compliant_folder, appeal_folder, unable_to_detect_folder): file_name
            for idx, file_name in enumerate(os.listdir(input_folder), start=1)
        }

        for future in as_completed(future_to_file):
            result = future.result()
            if result:
                results.append(result)

    if os.path.exists(results_file):
        existing_df = pd.read_excel(results_file, engine='openpyxl')
        new_df = pd.DataFrame(results, columns=['Serial No', 'File Name', 'File Processed Date and Time', 'Classification'])
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = pd.DataFrame(results, columns=['Serial No', 'File Name', 'File Processed Date and Time', 'Classification'])

    combined_df.to_excel(results_file, index=False, engine='openpyxl')
    style_excel(results_file)
    print(f"[INFO] Classification results saved to {results_file}")

if __name__ == "__main__":
    main()
