import fitz  # PyMuPDF for PDF handling
import easyocr  # EasyOCR for text recognition
from PIL import Image
import os
import io

# Initialize EasyOCR once to avoid loading the model multiple times
reader = easyocr.Reader(['en'])

# Function to check if the document contains handwritten text
def is_handwritten(file_path):
    try:
        # Check if the file is PDF or TIF
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.pdf':
            # Process each page of the PDF
            doc = fitz.open(file_path)
            for page in doc:
                # Convert PDF page to image
                pix = page.get_pixmap()
                img = Image.open(io.BytesIO(pix.tobytes("png")))

                # Perform OCR and check for handwritten text
                if check_handwritten(img):
                    return True  # Handwritten text found
            doc.close()

        elif file_ext == '.tif':
            # Open TIF image and process each page
            img = Image.open(file_path)
            for page in range(img.n_frames):
                img.seek(page)

                # Perform OCR and check for handwritten text
                if check_handwritten(img):
                    return True  # Handwritten text found

        return False  # No handwritten text detected in document

    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return False

# Helper function to perform OCR on an image and check if any text is likely handwritten
def check_handwritten(img):
    # Convert PIL image to bytes for EasyOCR
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='PNG')
    img_bytes = img_bytes.getvalue()

    # Perform OCR and analyze confidence
    result = reader.readtext(img_bytes)
    for _, text, prob in result:
        if prob > 0.5:  # Adjust confidence threshold as needed
            return True  # Considered as handwritten text
    return False

# Example usage
file_path = r"C:\Users\X8382\Documents\LTCIndexing\src\archive\191079198 Appeal.tiff"
if os.path.exists(file_path):
    if is_handwritten(file_path):
        print("The document contains handwritten text.")
    else:
        print("The document does not contain handwritten text.")
else:
    print(f"The file {file_path} does not exist.")
