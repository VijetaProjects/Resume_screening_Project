import pdfplumber
import pandas as pd
import os
import time
import docx2txt
from pathlib import Path
from ollama import Client

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

# Function to calculate a suitability score using LLM
def calculate_score(resume_text, job_description):
    if not isinstance(resume_text, str) or not isinstance(job_description, str):
        return 0

    prompt = f"""
    Based on the following resume and job description, rate this candidate's suitability for the role on a scale of 1-100.

    Resume Text: {resume_text}

    Job Description: {job_description}

    Provide the score in this format:
    - Score: [Score]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    score = 0
    
    # Extract the numeric score and ensure it's between 1 and 100
    lines = extracted_text.split('\n')
    for line in lines:
        if "Score:" in line:
            try:
                score = int(line.split(":", 1)[1].strip())
                score = max(1, min(score, 100))  # Ensure score is between 1 and 100
            except ValueError:
                score = 0
    
    return score

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

# Main function to extract, clean, and process resumes with scores
def process_resumes_with_scores(resume_folder, job_description_file, final_excel_path):
    all_extracted_data = []

    # Extract the job description text
    job_description_text = "\n".join(extract_text_from_file(job_description_file))
    
    # Record the start time
    start_time = time.time()

    # Loop through all files in the specified folder
    for filename in os.listdir(resume_folder):
        if filename.endswith((".pdf", ".docx", ".txt")):
            file_path = os.path.join(resume_folder, filename)
            print(f"Processing: {filename}")
            
            # Extract text from the current file
            resume_text = "\n".join(extract_text_from_file(file_path))
            
            # Calculate Suitability Score
            score = calculate_score(resume_text, job_description_text)
            
            # Collect all extracted information
            all_extracted_data.append({
                "Filename": filename,
                "Score": score
            })

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

# Specify the folder containing the resumes, the job description file, and the final Excel file path
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"
final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\resumes_with_scores.xlsx"

# Call the function to process the resumes and store the final data in an Excel file
process_resumes_with_scores(resume_folder, job_description_file, final_excel_file)
