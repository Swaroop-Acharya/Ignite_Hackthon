import os
import fitz  # PyMuPDF for PDF handling
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
import pandas as pd

def extract_text_from_pdf(pdf_file):
    """Extract text from each page of a PDF file."""
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text() + "\n\n"
    doc.close()
    return text

def ocr_image_file(image_file):
    """Extract text from a single image or multi-page TIF file."""
    text = ""
    img = Image.open(image_file)
    for page in range(img.n_frames):
        img.seek(page)
        text += pytesseract.image_to_string(img) + "\n\n"
    return text

def process_file(file_path):
    """Process a file (PDF or TIF) and return the extracted text."""
    file_ext = os.path.splitext(file_path)[1].lower()
    text = ""
    
    if file_ext == '.pdf':
        text = extract_text_from_pdf(file_path)
        if not text.strip():  # If no text found, treat as scanned PDF and use OCR
            images = convert_from_path(file_path)
            for img in images:
                text += pytesseract.image_to_string(img) + "\n\n"
    elif file_ext in ['.tif', '.tiff']:
        text = ocr_image_file(file_path)
    else:
        print(f"[ERROR] Unsupported file format: {file_ext}")
    
    return text

def main():
    input_folder = 'input'
    output_file = 'extracted_text_results.xlsx'
    
    # Create list to hold results
    results = []

    # Process each file in the input folder
    for file_name in os.listdir(input_folder):
        file_path = os.path.join(input_folder, file_name)
        print(f"[INFO] Processing {file_name}")
        
        text = process_file(file_path)
        results.append({"File Name": file_name, "Extracted Text": text})

    # Save results to an Excel file
    df = pd.DataFrame(results)
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"[INFO] Text extraction results saved to {output_file}")

if __name__ == "__main__":
    main()
