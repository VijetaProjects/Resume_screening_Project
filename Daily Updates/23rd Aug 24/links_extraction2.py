import pdfplumber
import re
import os
import pandas as pd

# Function to extract links using pdfplumber and fallback to regex if needed
def extract_hyperlinks_combined(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = set()  # Ensure LinkedIn links are unique
    text_links = set()

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

            # Fallback: Using regex to find links in the page's text
            page_text = page.extract_text()
            if page_text:
                # Regex to capture URLs
                urls = re.findall(r'(https?://[^\s]+|www\.[^\s]+)', page_text)
                for url in urls:
                    if 'github.com' in url:
                        github_links.add(url)
                    elif 'linkedin.com' in url:
                        linkedin_links.add(url)
                    else:
                        text_links.add(url)  # Store any other detected URLs
    
    return list(github_links), list(linkedin_links), list(text_links)

# Function to process multiple resumes and save the output to an Excel sheet
def process_resumes_and_save_to_excel(resume_folder, output_excel_path):
    results = []

    for filename in os.listdir(resume_folder):
        if filename.endswith(".pdf"):
            file_path = os.path.join(resume_folder, filename)
            print(f"Processing: {filename}")

            # Extract links from the current resume
            github_links, linkedin_links, other_links = extract_hyperlinks_combined(file_path)
            
            # Collect results for this resume
            results.append({
                "Filename": filename,
                "GitHub Links": ", ".join(github_links) if github_links else "No links",
                "LinkedIn Links": ", ".join(linkedin_links) if linkedin_links else "No links",
                "Other Links": ", ".join(other_links) if other_links else "No links"
            })

    # Convert the results to a DataFrame
    df = pd.DataFrame(results)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

# Specify the folder containing the resumes and the output Excel file path
resume_folder = r'C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes'
output_excel_path = r'extracted_links2.xlsx'

# Process the resumes and save the results to an Excel file
process_resumes_and_save_to_excel(resume_folder, output_excel_path)
