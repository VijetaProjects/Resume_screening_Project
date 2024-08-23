import pdfplumber
import re
import os
import pandas as pd

# Function to extract links using pdfplumber
def extract_links_pdfplumber(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = set()  # Ensure LinkedIn links are unique

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extracting annotations
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

# Fallback function to extract links using regex if pdfplumber didn't find any
def extract_links_regex(pdf_path):
    github_links = set()
    linkedin_links = set()

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract text and search for links using regex
            page_text = page.extract_text()
            if page_text:
                # Regex to capture URLs
                urls = re.findall(r'(https?://[^\s]+|www\.[^\s]+)', page_text)
                for url in urls:
                    if 'github.com' in url:
                        github_links.add(url)
                    elif 'linkedin.com' in url:
                        linkedin_links.add(url)
    
    return list(github_links), list(linkedin_links)

# Function to process multiple resumes and save the output to an Excel sheet
def process_resumes_and_save_to_excel(resume_folder, output_excel_path):
    results = []

    for filename in os.listdir(resume_folder):
        if filename.endswith(".pdf"):
            file_path = os.path.join(resume_folder, filename)
            print(f"Processing: {filename}")

            # First try extracting links using pdfplumber
            github_links, linkedin_links = extract_links_pdfplumber(file_path)

            # If no links were found, fall back to regex extraction
            if not github_links and not linkedin_links:
                print(f"No links found using pdfplumber for {filename}. Falling back to regex extraction.")
                github_links, linkedin_links = extract_links_regex(file_path)
            
            # Collect results for this resume
            results.append({
                "Filename": filename,
                "GitHub Links": ", ".join(github_links) if github_links else "No links",
                "LinkedIn Links": ", ".join(linkedin_links) if linkedin_links else "No links",
            })

    # Convert the results to a DataFrame
    df = pd.DataFrame(results)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

# Specify the folder containing the resumes and the output Excel file path
resume_folder = r'C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes'
output_excel_path = r'extracted_links3.xlsx'

# Process the resumes and save the results to an Excel file
process_resumes_and_save_to_excel(resume_folder, output_excel_path)
