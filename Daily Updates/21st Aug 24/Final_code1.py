import pdfplumber
import pandas as pd
import os
import re
from ollama import Client
import time

# Initialize the LLM client
client = Client()

# Function to clean the extracted text by removing non-printable characters
def clean_text(text):
    return ''.join(char for char in text if char.isprintable())

# Function to extract text from a single PDF and return it as a list
def extract_text_from_pdf(pdf_path):
    text_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                cleaned_text = clean_text(page_text)
                text_data.extend(cleaned_text.split('\n'))
    return text_data

# Function to clean text using regular expressions for better formatting
def clean_text_column(text):
    if isinstance(text, str):  # Check if the input is a string
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
            phone_number = line.split(":", 1)[1].strip() or "Not mentioned"
    
    return phone_number

# Function to extract all GitHub and LinkedIn hyperlinks from the PDF using pdfplumber
def extract_hyperlinks(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = []
    
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
                                linkedin_links.append(url)
    
    return list(github_links), linkedin_links

# Main function to extract, clean, and process resumes
def pdfs_to_cleaned_and_extracted_excel(pdf_folder, final_excel_path):
    all_extracted_data = []

    # Record the start time
    start_time = time.time()

    # Loop through all PDF files in the specified folder
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            print(f"Processing: {filename}")
            
            # Extract text from the current PDF
            pdf_text = "\n".join(extract_text_from_pdf(pdf_path))
            
            # Clean the extracted text
            cleaned_text = clean_text_column(pdf_text)
            
            # Extract Name and Location using LLM
            extracted_info = extract_information_llm(cleaned_text)
            
            # Extract Phone Number
            phone_number = extract_phone_number(cleaned_text)
            
            # Extract GitHub and LinkedIn links
            github_links, linkedin_links = extract_hyperlinks(pdf_path)
            
            # Join multiple links into a single string (comma-separated)
            github_links_str = ', '.join(github_links) if github_links else "Not mentioned"
            linkedin_links_str = ', '.join(linkedin_links) if linkedin_links else "Not mentioned"
            
            # Collect all extracted information
            all_extracted_data.append({
                "Filename": filename,
                "Name": extracted_info["Name"],
                "Location": extracted_info["Location"],
                "Phone Number": phone_number,
                "Github Links": github_links_str,
                "LinkedIn Links": linkedin_links_str
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

# Specify the folder containing the PDFs and the final Excel file path
pdf_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
final_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\final_resume_info1.xlsx"

# Call the function to process the resumes and store the final data in an Excel file
pdfs_to_cleaned_and_extracted_excel(pdf_folder, final_excel_file)
