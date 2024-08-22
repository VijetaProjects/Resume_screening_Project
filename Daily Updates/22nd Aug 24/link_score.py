import pdfplumber
import pandas as pd
import os
import re
from ollama import Client
import time
import docx2txt
from pathlib import Path

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

# Function to clean the 'Extracted Text' column in a DataFrame
def clean_extracted_text_column(df, column_name='Extracted Text'):
    df[column_name] = df[column_name].apply(clean_text_column)
    return df

from ollama import Client

# Initialize the LLM client
client = Client()

# Function to extract GitHub and LinkedIn links using LLM
def extract_links(text):
    github_links = set()
    linkedin_links = set()

    prompt = f"""
    Extract all GitHub and LinkedIn URLs from the following text. The links can be plain text or embedded as hyperlinks.

    Text: {text}

    Provide the extracted links in this format:
    - GitHub Links: [comma-separated GitHub links]
    - LinkedIn Links: [comma-separated LinkedIn links]
    """

    # Generate the response using the LLM
    response = client.generate(model="llama3:latest", prompt=prompt)
    extracted_text = response.get('response', '')

    # Split the response to get GitHub and LinkedIn links
    github_links_str = ""
    linkedin_links_str = ""
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "GitHub Links:" in line:
            github_links_str = line.split(":", 1)[1].strip()
        elif "LinkedIn Links:" in line:
            linkedin_links_str = line.split(":", 1)[1].strip()

    # Convert the comma-separated string to a list and remove any empty entries
    if github_links_str:
        github_links = set([link.strip() for link in github_links_str.split(",") if link.strip()])
    if linkedin_links_str:
        linkedin_links = set([link.strip() for link in linkedin_links_str.split(",") if link.strip()])

    return list(github_links), list(linkedin_links)


# Function to generate a suitability score using LLM
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

# Main function to extract, clean, and process resumes
def pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, final_excel_path):
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
            
            # Clean the extracted text
            cleaned_text = clean_text_column(resume_text)
            
            # Extract GitHub and LinkedIn links from the text
            github_links, linkedin_links = extract_links(cleaned_text)
            
            # Calculate Suitability Score
            score = calculate_score(cleaned_text, job_description_text)
            
            # Join multiple links into a single string (comma-separated)
            github_links_str = ', '.join(github_links) if github_links else "Not mentioned"
            linkedin_links_str = ', '.join(linkedin_links) if linkedin_links else "Not mentioned"
            
            # Collect all extracted information
            all_extracted_data.append({
                "Filename": filename,
                "Github Links": github_links_str,
                "LinkedIn Links": linkedin_links_str,
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
final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\link_score_extraction.xlsx"

# Call the function to process the resumes and store the final data in an Excel file
pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, final_excel_file)
