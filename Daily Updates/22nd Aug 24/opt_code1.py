import pdfplumber
import pandas as pd
import os
import re
from ollama import Client
import time
import docx2txt
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor

# Initialize the LLM client
client = Client()

# Function to clean the extracted text by removing non-printable characters
def clean_text(text):
    return ''.join(char for char in text if char.isprintable())

# Function to extract text from PDF files
def extract_text_from_pdf(pdf_path):
    text_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                cleaned_text = clean_text(page_text)
                text_data.extend(cleaned_text.split('\n'))
    return text_data

# Function to extract text from .docx files
def extract_text_from_docx(docx_path):
    return docx2txt.process(docx_path).split('\n')

# Function to extract text from .txt files
def extract_text_from_txt(txt_path):
    with open(txt_path, 'r', encoding='utf-8') as file:
        return file.readlines()

# Function to clean text using regular expressions for better formatting
def clean_text_column(text):
    if isinstance(text, str):
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        text = re.sub(r'(\d)([A-Z])', r'\1 \2', text)
        text = re.sub(r'([A-Z])([A-Z][a-z])', r'\1 \2', text)
        text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
    return text

# Function to process each resume in parallel
def process_resume(file_path, job_description_text):
    # Extract text from the current file
    resume_text = "\n".join(extract_text_from_file(file_path))
    
    # Clean the extracted text
    cleaned_text = clean_text_column(resume_text)
    
    # Batch request to LLM to extract all information at once
    prompt = f"""
    Extract the following information from the resume:
    1. Name and Location.
    2. Phone number.
    3. GitHub and LinkedIn links.
    4. Fitment summary based on the job description.
    5. Total experience.
    6. Suitability score (1-100).

    Resume Text: {cleaned_text}

    Job Description: {job_description_text}
    """
    
    response = client.generate(model="llama3:latest", prompt=prompt)
    extracted_text = response.get('response', 'No response text found.')
    
    extracted_info = {
        "Name": "Not mentioned",
        "Location": "Not mentioned",
        "Phone Number": "Not mentioned",
        "Github Links": "Not mentioned",
        "LinkedIn Links": "Not mentioned",
        "Fitment Summary": "Not mentioned",
        "Total Experience": "Fresher or Not mentioned",
        "Score": 0
    }
    
    # Parse LLM response for different fields
    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Phone Number:" in line:
            extracted_info["Phone Number"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "GitHub Links:" in line:
            extracted_info["Github Links"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "LinkedIn Links:" in line:
            extracted_info["LinkedIn Links"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Fitment Summary:" in line:
            extracted_info["Fitment Summary"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Experience:" in line:
            extracted_info["Total Experience"] = line.split(":", 1)[1].strip() or "Fresher or Not mentioned"
        elif "Score:" in line:
            try:
                score = int(line.split(":", 1)[1].strip())
                extracted_info["Score"] = max(1, min(score, 100))  # Ensure score is between 1 and 100
            except ValueError:
                extracted_info["Score"] = 0
    
    return extracted_info

# Function to determine file type and extract text accordingly
def extract_text_from_file(file_path):
    ext = Path(file_path).suffix.lower()
    if ext == '.pdf':
        return extract_text_from_pdf(file_path)
    elif ext == '.docx':
        return extract_text_from_docx(file_path)
    elif ext == '.txt':
        return extract_text_from_txt(file_path)
    else:
        return []

# Main function to extract, clean, and process resumes
def pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, final_excel_path):
    all_extracted_data = []

    # Extract the job description text
    job_description_text = "\n".join(extract_text_from_file(job_description_file))
    
    # Record the start time
    start_time = time.time()

    # Create a list of resume file paths
    resume_files = [os.path.join(resume_folder, f) for f in os.listdir(resume_folder) if f.endswith((".pdf", ".docx", ".txt"))]

    # Use parallel processing to handle multiple resumes
    with ProcessPoolExecutor() as executor:
        results = list(executor.map(process_resume, resume_files, [job_description_text]*len(resume_files)))
    
    # Collect results
    for result in results:
        all_extracted_data.append(result)

    # Convert the list of extracted info into a DataFrame
    extracted_df = pd.DataFrame(all_extracted_data)

    # Save the extracted information to a final Excel file
    extracted_df.to_excel(final_excel_path, index=False)

    # Record the end time
    end_time = time.time()

    # Calculate the duration
    duration = end_time - start_time

    print(f"Final data saved to {final_excel_path}")
    print(f"Time taken: {duration:.2f} seconds")

if __name__ == "__main__":
    # Ensure that multiprocessing is handled correctly
    import multiprocessing
    multiprocessing.freeze_support()

    # Specify the folder containing the resumes, the job description file, and the final Excel file path
    resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
    job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"
    final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\opt_code_result.xlsx"

    # Call the function to process the resumes and store the final data in an Excel file
    pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, final_excel_file)
