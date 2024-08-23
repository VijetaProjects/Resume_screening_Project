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

# Function to extract information using LLM prompt engineering
def extract_information_llm(resume_text):
    if not isinstance(resume_text, str):
        return {"Name": "Not mentioned", "Location": "Not mentioned"}

    prompt = f"""
    Extract the name and location from the following resume text. The name may appear anywhere in the text and may be capitalized.

    Resume Text: {resume_text}

    Provide the extracted information in this format:
    - Name: [Extracted Name]
    - Location: [Extracted Location]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    extracted_info = {
        "Name": "Not mentioned",
        "Location": "Not mentioned"
    }

    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip() or "Not mentioned"
    
    return extracted_info

# Function to extract phone number using LLM prompt engineering
def extract_phone_number(resume_text):
    if not isinstance(resume_text, str):
        return "Not mentioned"

    prompt = f"""
    Extract the phone number from the following resume text. The phone number may appear in various formats, such as international or local.

    Resume Text: {resume_text}

    Provide the extracted phone number in this format:
    - Phone Number: [Extracted Phone Number]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    phone_number = "Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Phone Number:" in line:
            extracted_value = line.split(":", 1)[1].strip()
            if extracted_value.lower() == "none" or not extracted_value:
                phone_number = "Not mentioned"
            else:
                phone_number = extracted_value
    
    return phone_number

# Function to extract all GitHub and LinkedIn hyperlinks from the PDF using pdfplumber
def extract_hyperlinks(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = set()  # Ensure LinkedIn links are unique
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        url = annot['uri']
                        if url:
                            if 'github.com' in url:
                                github_links.add(url)  # Add to set to remove duplicates
                            elif 'linkedin.com' in url:
                                linkedin_links.add(url)
    
    return list(github_links), list(linkedin_links)

# Function to generate a fitment summary using LLM
def fitment_summary(resume_text, job_description):
    if not isinstance(resume_text, str) or not isinstance(job_description, str):
        return "Not mentioned"
    
    prompt = f"""
    Based on the following resume and job description, provide a summary of how this candidate is suitable for the role in 50 words.

    Resume Text: {resume_text}

    Job Description: {job_description}

    Provide the fitment summary in this format:
    - Summary: [50-word Summary]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    summary = "Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Summary:" in line:
            summary = line.split(":", 1)[1].strip() or "Not mentioned"
    
    return summary

# Function to calculate total experience using LLM
def total_experience(resume_text):
    if not isinstance(resume_text, str):
        return "Fresher or Not mentioned"

    prompt = f"""
    Calculate the total years of experience from the work experience section of the following resume text. If no work experience section is available, return 'Fresher or Not mentioned.'

    Resume Text: {resume_text}

    Provide the total experience in this format:
    - Experience: [Total Experience in Years]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    experience = "Fresher or Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Experience:" in line:
            experience = line.split(":", 1)[1].strip() or "Fresher or Not mentioned"
    
    return experience

# Function to calculate a suitability score using LLM
def calculate_score(resume_text, job_description):
    if not isinstance(resume_text, str) or not isinstance(job_description, str):
        return "Not mentioned"

    prompt = f"""
    Based on the job description below, evaluate the suitability of the candidate described in the resume. 
    Provide a score from 1 to 100 on how suitable the candidate is for the job role.

    Job Description:
    {job_description}

    Candidate Resume:
    {resume_text}

    Please respond with only the numerical score.
    """

    try:
        response = client.generate(model="llama3:latest", prompt=prompt)
        score = int(response['response'].strip())
    except (ValueError, KeyError):
        score = None

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

# Function to evaluate candidate's resume against the extracted skills
def evaluate_candidate(skill, resume_text):
    prompt = f"Does the candidate have the skill '{skill}'? If so, briefly describe their experience with it.\n\nCandidate Resume:\n{resume_text}"
    response = client.generate(model='llama3:latest', prompt=prompt)
    return response.get('response', "Error: Candidate observation generation failed.")

# Main function to extract, clean, and process resumes
def pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, skills_file, final_excel_path):
    all_extracted_data = []

    # Extract the job description text
    job_description_text = "\n".join(extract_text_from_file(job_description_file))
    
    # Load skills and requirements from the provided Excel sheet
    skills_df = pd.read_excel(skills_file)
    skills = skills_df['Skills'].tolist()  # Assuming the column is named 'Skills'
    
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
            
            # Extract Name and Location using LLM
            extracted_info = extract_information_llm(cleaned_text)
            
            # Extract Phone Number
            phone_number = extract_phone_number(cleaned_text)
            
            # Extract GitHub and LinkedIn links
            github_links, linkedin_links = extract_hyperlinks(file_path)
            
            # Generate Fitment Summary
            summary = fitment_summary(cleaned_text, job_description_text)
            
            # Calculate Total Experience
            experience = total_experience(cleaned_text)
            
            # Calculate Suitability Score
            score = calculate_score(cleaned_text, job_description_text)
            
            # Join multiple links into a single string (comma-separated)
            github_links_str = ', '.join(github_links) if github_links else "Not mentioned"
            linkedin_links_str = ', '.join(linkedin_links) if linkedin_links else "Not mentioned"
            
            # Initialize a dictionary to store all extracted information
            extracted_data = {
                "Filename": filename,
                "Name": extracted_info["Name"],
                "Location": extracted_info["Location"],
                "Phone Number": phone_number,
                "Github Links": github_links_str,
                "LinkedIn Links": linkedin_links_str,
                "Total Experience": experience,
                "Fitment Summary": summary,
                "Score": score
            }
            
            # Generate candidate observations for each skill
            for skill in skills:
                observation = evaluate_candidate(skill, cleaned_text)
                extracted_data[skill] = observation
            
            # Append the extracted data to the list
            all_extracted_data.append(extracted_data)

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

# Specify the folder containing the resumes, the job description file, the skills file, and the final Excel file path
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"
skills_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\Evaluation Criteria Sheet.xlsx"
final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\23 Aug\final_info_with_skills.xlsx"

# Call the function to process the resumes and store the final data in an Excel file
pdfs_to_cleaned_and_extracted_excel(resume_folder, job_description_file, skills_file, final_excel_file)
