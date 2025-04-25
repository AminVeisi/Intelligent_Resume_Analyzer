# Resume Processing and Relevance Scoring System

This Python script processes resumes in PDF format, extracts and structures their content, and calculates a relevance score based on a given job description using NLP techniques.

## Features

- **PDF Text Extraction**: Extracts text from PDF resumes using `pdfplumber` (including table data)
- **OCR Support**: Can process scanned resumes using `pytesseract` (though not implemented in main flow)
- **Text Cleaning**: Removes special characters and normalizes text
- **Section Classification**: Identifies and categorizes resume sections (Contact Info, Profile, Skills, etc.)
- **Relevance Scoring**: 
  - Uses TF-IDF vectorization to convert text to numerical features
  - Calculates cosine similarity between resume and job description
- **Batch Processing**: Processes all PDFs in a folder and outputs scores to Excel

## Installation

1. **Prerequisites**:
   - Python 3.7+
   - Tesseract OCR installed on your system (for OCR functionality)

2. **Install dependencies**:
   ```bash
   pip install pdfplumber pytesseract pillow scikit-learn pandas openpyxl

## Usage

1. **Prepare your files**
   - Place all resumes in PDF format in a folder
   - Create a text file with the job description

2. **Configure paths (in the script or via command line arguments)**:
   ```bash
   folder_path = "path/to/resumes"  # Folder containing PDF resumes
   job_description_path = "path/to/job_description.txt"  # Job description file
3. **Run the python script**
4. **Output**:
   - The script will create an Excel file (resume_scores.xlsx) in the same folder as the resumes
   - Contains filename and relevance score (0-1) for each resume
