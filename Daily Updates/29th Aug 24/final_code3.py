import pdfplumber
import re
import os
import pandas as pd
import time
import docx2txt
from pathlib import Path
from ollama import Client
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS
import nltk

# Download the required NLTK resources
nltk.download('stopwords')
from nltk.corpus import stopwords

# Initialize the LLM client
client = Client()

# Function to clean the extracted text by removing non-printable characters and stopwords
def clean_text(text):
    text = ''.join(char for char in text if char.isprintable())
    stop_words = set(stopwords.words('english')).union(set(ENGLISH_STOP_WORDS))
    cleaned_text = ' '.join(word for word in text.split() if word.lower() not in stop_words)
    return cleaned_text

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
    return clean_text(docx2txt.process(docx_path)).split('\n')

# Function to extract text from .txt files
def extract_text_from_txt(txt_path):
    with open(txt_path, 'r', encoding='utf-8') as file:
        return clean_text(file.read()).splitlines()

# Function to clean text using regular expressions for better formatting
def clean_text_column(text):
    if isinstance(text, str):
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        text = re.sub(r'(\d)([A-Z])', r'\1 \2', text)
        text = re.sub(r'([A-Z])([A-Z][a-z])', r'\1 \2', text)
        text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
    return text

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

# Function to extract links using pdfplumber
def extract_links_pdfplumber(pdf_path):
    github_links = set()
    linkedin_links = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        url = annot['uri']
                        if url:
                            if 'github.com' in url:
                                github_links.add(url)
                            elif 'linkedin.com' in url:
                                linkedin_links.add(url)
    
    return list(github_links), list(linkedin_links)

# Fallback function to extract links using regex if pdfplumber didn't find any
def extract_links_regex(pdf_path):
    github_links = set()
    linkedin_links = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                urls = re.findall(r'(https?://[^\s]+|www\.[^\s]+)', page_text)
                for url in urls:
                    if 'github.com' in url:
                        github_links.add(url)
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

# Function to extract text from different file types
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
    prompt = f"Does the candidate have the skill '{skill}'? If so, briefly describe their experience with it in 50-100 words.\n\nCandidate Resume:\n{resume_text}"
    response = client.generate(model='llama3:latest', prompt=prompt)
    result = response.get('response', "Not mentioned")
    
    if "not mentioned" in result.lower() or "does not have" in result.lower():
        return "Not mentioned"
    return result

# Main function to extract, clean, and process resumes
def pdfs_to_cleaned_and_extracted_excel(folder_path, job_description_file, skills_file, existing_excel_path, final_excel_path):
    start_time = time.time()

    # Load the job description and skills into memory once
    job_description_text = "\n".join(extract_text_from_file(job_description_file))
    skills_df = pd.read_excel(skills_file)
    skills = skills_df['Skills'].tolist()  # Assuming the column is named 'Skills'

    # Load the existing Excel file if it exists
    if os.path.exists(existing_excel_path):
        df_existing = pd.read_excel(existing_excel_path)
    else:
        df_existing = pd.DataFrame()

    all_extracted_data = []

    # Process each resume in the folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)

        if Path(file_path).suffix.lower() not in ['.pdf', '.docx', '.txt']:
            continue

        # Extract the resume text
        resume_text = ' '.join(extract_text_from_file(file_path))

        # Extract name and location using LLM
        extracted_info = extract_information_llm(resume_text)
        name = clean_text_column(extracted_info.get("Name", "Not mentioned"))
        location = clean_text_column(extracted_info.get("Location", "Not mentioned"))

        # Extract phone number using LLM
        phone_number = clean_text_column(extract_phone_number(resume_text))

        # Extract GitHub and LinkedIn links
        github_links, linkedin_links = extract_links_pdfplumber(file_path)
        if not github_links and not linkedin_links:
            github_links, linkedin_links = extract_links_regex(file_path)

        # Generate fitment summary using LLM
        fitment_summary_text = fitment_summary(resume_text, job_description_text)

        # Calculate total experience using LLM
        total_experience_years = total_experience(resume_text)

        # Calculate the suitability score using LLM
        suitability_score = calculate_score(resume_text, job_description_text)

        # Evaluate candidate's skills using LLM
        skill_matches = {skill: evaluate_candidate(skill, resume_text) for skill in skills}

        # Append the extracted data to the list
        extracted_data = {
            "Filename": filename,
            "Name": name,
            "Location": location,
            "Phone Number": phone_number,
            "GitHub Links": ", ".join(github_links),
            "LinkedIn Links": ", ".join(linkedin_links),
            "Fitment Summary": fitment_summary_text,
            "Total Experience": total_experience_years,
            "Suitability Score": suitability_score,
        }
        extracted_data.update(skill_matches)

        all_extracted_data.append(extracted_data)

    # Combine existing and new data
    df_combined = pd.concat([df_existing, pd.DataFrame(all_extracted_data)], ignore_index=True)

    # Write the final output to an Excel file
    df_combined.to_excel(final_excel_path, index=False)

    end_time = time.time()
    print(f"Processing completed in {end_time - start_time:.2f} seconds.")

# Call the main function to start processing
folder_path = r'C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\10 Resumes'
job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"
skills_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\Evaluation Criteria Sheet.xlsx"
existing_excel_path = "existing_data.xlsx"
final_excel_path = "final_output3.xlsx"

pdfs_to_cleaned_and_extracted_excel(folder_path, job_description_file, skills_file, existing_excel_path, final_excel_path)
