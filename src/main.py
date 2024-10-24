import os
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from pdf2image import convert_from_path

# Function to extract text from PDF (text-based)
def extract_text_from_pdf(pdf_file):
    doc = fitz.open(pdf_file)
    text = ""
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text += page.get_text()  # Get the text from the page
    doc.close()
    return text

# Function to OCR scanned PDFs or TIF files
def ocr_image_file(image_file):
    if isinstance(image_file, str):  # If it's a file path
        img = Image.open(image_file)
    else:  # If it's already a PIL Image object
        img = image_file
    text = pytesseract.image_to_string(img)
    return text

# Function to process PDF and TIF files
def process_file(file_path):
    file_ext = os.path.splitext(file_path)[1].lower()
    text = ""
    
    if file_ext == '.pdf':
        text = extract_text_from_pdf(file_path)
        if not text.strip():  # If no text found, treat as scanned PDF
            images = convert_from_path(file_path)
            for img in images:
                text += ocr_image_file(img) + "\n\n"
    elif file_ext == '.tif':
        text = ocr_image_file(file_path)
    else:
        print(f"Unsupported file format: {file_ext}")
    
    return text

# Main script
def main():
    input_folder = 'input_files'
    output_folder = 'output'

    for file_name in os.listdir(input_folder):
        file_path = os.path.join(input_folder, file_name)
        text = process_file(file_path)

        output_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_output.txt")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text)

        print(f"Processed {file_name} and saved output to {output_file}")

if __name__ == "__main__":
    main()