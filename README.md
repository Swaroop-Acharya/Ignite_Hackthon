# 🤖 DocuSort – Smart Document Classifier & Summarizer

**DocuSort** is an intelligent automation system that classifies and summarizes scanned documents such as customer correspondence in PDF or TIF formats. Leveraging OCR, NLP, and summarization pipelines, it streamlines customer service workflows by identifying complaints, appeals, or ambiguous content requiring manual review.

---

## 🧠 Features

- 🧾 **OCR Support**: Extracts text from both text-based and scanned PDFs/TIFs using Tesseract and PyMuPDF.
- 🏷️ **Auto Classification**: Categorizes documents into `Complaint`, `Appeal`, or `Manual` based on customizable keyword sets.
- 🧠 **AI Summarization**: Generates concise summaries of long documents using Hugging Face’s `bart-large-cnn` model.
- 📊 **Excel Report Generation**: Compiles all results in a styled, timestamped Excel sheet.
- 🔁 **Multithreading Support**: Processes multiple files concurrently for maximum performance.
- 📂 **File Archival**: Automatically organizes processed files into categorized archive folders.

---

## 🛠️ Tech Stack

- **Python 3.9+**
- **OCR**: PyMuPDF (`fitz`), Tesseract, PIL, `pdf2image`
- **NLP**: HuggingFace Transformers (BART model)
- **Excel Styling**: openpyxl
- **Concurrency**: ThreadPoolExecutor + Locks
- **Data Handling**: pandas

---

## 🚀 How to Use

1. 📥 **Place input files**: Drop `.pdf` or `.tif` files into the `/input` folder.
2. ▶️ **Run the script**: Execute `main()` from the script to start processing.
3. 📊 **Get results**:
    - Extracted text saved in `/extracted_text/<Category>/`
    - Original files archived in `/archive/<Category>/`
    - Classification report saved as an Excel file in `/classification_reports/`

---

## 📁 Directory Structure

```
📂 input/
📂 archive/
    ├── Complaint/
    ├── Appeal/
    └── Manual/
📂 extracted_text/
    ├── Complaint/
    ├── Appeal/
    └── Manual/
📂 classification_reports/
📂 keywords/
    ├── complaint_keywords.txt
    └── appeal_keywords.txt
```

---

## 🧪 Sample Output

Each row in the Excel report includes:
- Serial Number
- File Name
- Processed Timestamp
- Classification Result
- AI-Generated Summary

---

## 🛡️ Customization

- Add new keywords to the `/keywords/complaint_keywords.txt` or `/keywords/appeal_keywords.txt` files to tailor classification.
- Plug in other transformer models if domain-specific summarization is required.

---

## 🧬 Use Cases

- Prioritize sensitive customer communication
- Automate triaging of service requests
- Digitize and analyze scanned records with high efficiency

---

Built with ❤️ to make document handling intelligent, fast, and hassle-free.
