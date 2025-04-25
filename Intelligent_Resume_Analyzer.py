import pdfplumber
import pytesseract
from PIL import Image
import re
import json
import os
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd

# Example Usage Path
folder_path = r"E:\NLP\NLP_Project\Intelligent_Resume_Analyzer"  # Path to the folder containing PDF resumes
job_description_path = r"E:\NLP\NLP_Project\Intelligent_Resume_Analyzer\job_description.txt"  # Path to the job description file

# Predefined resume section headers
SECTION_HEADERS = ["CONTACT INFO", "PROFILE", "KEY SKILLS", "WORK EXPERIENCE", "EDUCATION"]

# Step 1: Extract text from PDFs
def extract_text_from_pdf(file_path):
    text = ""
    tables = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
            table = page.extract_table()
            if table:
                tables.append(table)
    return text.strip(), tables

# Step 1: Extract text with OCR for scanned PDFs
def extract_text_with_ocr(image_path):
    return pytesseract.image_to_string(Image.open(image_path))

# Step 1: Clean extracted text
def clean_text(raw_text):
    return re.sub(r"[^a-zA-Z0-9\s•-]", "", raw_text).strip()

# Step 1: Categorize text into structured sections
def classify_resume_text(text):
    structured_data = {}
    current_section = None
    lines = text.split("\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        for header in SECTION_HEADERS:
            if re.search(rf"\b{header}\b", line, re.IGNORECASE):
                current_section = header
                structured_data[current_section] = []
                break

        if current_section:
            structured_data[current_section].append(line)

    return structured_data

# Step 2: Calculate similarity score using TF-IDF and cosine similarity
def calculate_similarity_with_tfidf(resume_data, job_description):
    # Combine resume sections into a single text string
    resume_text = " ".join(
        resume_data.get("KEY SKILLS", []) +
        resume_data.get("WORK EXPERIENCE", []) +
        resume_data.get("PROFILE", [])
    )
    corpus = [resume_text, job_description]

    # Vectorize the corpus using TF-IDF
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform(corpus)

    # Calculate cosine similarity between the resume and job description
    similarity_score = cosine_similarity(tfidf_matrix[0], tfidf_matrix[1])[0][0]

    return round(similarity_score, 2)

# Process and structure resume + calculate relevance
def process_resume(file_path, job_description):
    try:
        # Attempt to extract text with pdfplumber
        text, tables = extract_text_from_pdf(file_path)

        # If text is empty, use OCR for scanned PDFs
        if not text.strip():
            print(f"Using OCR for scanned PDF: {file_path}")
            text = extract_text_with_ocr(file_path)
            tables = []  # OCR doesn't extract tables

        # Clean and structure the extracted text
        cleaned_text = clean_text(text)
        structured_resume = classify_resume_text(cleaned_text)

        # Include tables (if applicable and available)
        if tables:
            structured_resume["TABLES"] = tables

        # Calculate the relevance score
        relevance_score = calculate_similarity_with_tfidf(structured_resume, job_description)

        return file_path, relevance_score

    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return file_path, 0  # Return 0 score for files with errors

# Main function to process all resumes and save scores in an Excel file
def process_all_resumes(folder_path, job_description_path):
    # Read job description from the text file
    with open(job_description_path, "r", encoding="utf-8") as file:
        job_description = file.read().strip()

    # Validate folder existence
    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist.")
        return

    # Process all PDF files in the folder
    scores = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(folder_path, file_name)
            _, relevance_score = process_resume(file_path, job_description)
            scores.append({"File Name": file_name, "Relevance Score": relevance_score})

    # Save scores to an Excel file
    df = pd.DataFrame(scores)
    output_file = os.path.join(folder_path, "resume_scores.xlsx")
    df.to_excel(output_file, index=False)
    print(f"Scores saved to {output_file}")

# Example Usage
process_all_resumes(folder_path, job_description_path)
