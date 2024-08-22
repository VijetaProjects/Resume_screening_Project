import pdfplumber
import pandas as pd
import os
import time
import docx2txt
from pathlib import Path
from urllib.parse import urlparse

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

# Function to extract URLs and partial URLs from text
def extract_links_from_text(text):
    github_links = set()
    linkedin_links = set()

    # Split the text into words
    words = text.split()

    for word in words:
        parsed_url = urlparse(word)
        # If the word is a complete URL, use the netloc to identify the domain
        if parsed_url.scheme in ["http", "https"]:
            if "github.com" in parsed_url.netloc:
                github_links.add(word)
            elif "linkedin.com" in parsed_url.netloc:
                linkedin_links.add(word)
        # If it's not a complete URL, look for patterns that indicate a partial URL
        elif "github.com" in word:
            github_links.add(word)
        elif "linkedin.com/in/" in word:
            linkedin_links.add(word)

    return list(github_links), list(linkedin_links)

# Function to extract hyperlinks from PDF annotations (for embedded links)
def extract_hyperlinks_from_pdf(pdf_path):
    github_links = set()
    linkedin_links = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        url = annot['uri']
                        parsed_url = urlparse(url)
                        if parsed_url.scheme in ["http", "https"]:
                            if "github.com" in parsed_url.netloc:
                                github_links.add(url)
                            elif "linkedin.com" in parsed_url.netloc:
                                linkedin_links.add(url)

    return list(github_links), list(linkedin_links)

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
            
            # Extract GitHub and LinkedIn links from the text and PDF annotations
            github_links_text, linkedin_links_text = extract_links_from_text(resume_text)
            github_links_pdf, linkedin_links_pdf = extract_hyperlinks_from_pdf(file_path)

            # Combine links from text and PDF annotations
            github_links = set(github_links_text).union(github_links_pdf)
            linkedin_links = set(linkedin_links_text).union(linkedin_links_pdf)
            
            # Join multiple links into a single string (comma-separated)
            github_links_str = ', '.join(github_links) if github_links else "Not mentioned"
            linkedin_links_str = ', '.join(linkedin_links) if linkedin_links else "Not mentioned"
            
            # Collect all extracted information
            all_extracted_data.append({
                "Filename": filename,
                "Github Links": github_links_str,
                "LinkedIn Links": linkedin_links_str,
                # Assuming score calculation is done elsewhere or omitted for now
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
final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\links_extraction.xlsx"

# Call the function to process the resumes and store the final data in an Excel file
pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, final_excel_file)
