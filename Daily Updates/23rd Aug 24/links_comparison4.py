import re
import time
import os
import pdfplumber
from flashtext import KeywordProcessor
import ahocorasick
import pandas as pd

# Function to extract text from PDF using pdfplumber
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text

### Method 1: Regular Expressions (re)
def extract_links_re(text):
    github_links = set(re.findall(r'github\.com/[^\s]+', text))
    linkedin_links = set(re.findall(r'linkedin\.com/in/[^\s]+', text))
    return github_links, linkedin_links

### Method 2: FlashText
def extract_links_flashtext(text):
    github_links = set()
    linkedin_links = set()
    
    keyword_processor = KeywordProcessor()
    keyword_processor.add_keyword("github.com")
    keyword_processor.add_keyword("linkedin.com")
    
    keywords_found = keyword_processor.extract_keywords(text)
    
    for word in keywords_found:
        if "github.com" in word:
            github_links.add(word)
        elif "linkedin.com" in word:
            linkedin_links.add(word)
    
    return github_links, linkedin_links

### Method 3: Aho-Corasick
def extract_links_ahocorasick(text):
    github_links = set()
    linkedin_links = set()
    
    A = ahocorasick.Automaton()
    A.add_word("github.com", "github.com")
    A.add_word("linkedin.com", "linkedin.com")
    A.make_automaton()
    
    for end_index, found_word in A.iter(text):
        if "github.com" in found_word:
            github_links.add(found_word)
        elif "linkedin.com" in found_word:
            linkedin_links.add(found_word)
    
    return github_links, linkedin_links

# Function to benchmark each method and return the results
def benchmark_method(method, method_name, pdf_path):
    text = extract_text_from_pdf(pdf_path)
    
    start_time = time.time()
    github_links, linkedin_links = method(text)
    end_time = time.time()
    
    # You can set expected values here for accuracy checking if known
    expected_github_links = set()  # Set expected values here
    expected_linkedin_links = set()  # Set expected values here
    
    github_accuracy = (github_links == expected_github_links) if expected_github_links else "Not checked"
    linkedin_accuracy = (linkedin_links == expected_linkedin_links) if expected_linkedin_links else "Not checked"
    
    # Return the results as a dictionary
    return {
        "Method": method_name,
        "Time Taken (s)": round(end_time - start_time, 6),
        "GitHub Links": ", ".join(github_links) if github_links else "No links",
        "LinkedIn Links": ", ".join(linkedin_links) if linkedin_links else "No links",
        "GitHub Accuracy": github_accuracy,
        "LinkedIn Accuracy": linkedin_accuracy
    }

# Main function to process all PDFs and store results in Excel
def process_pdfs_and_save_to_excel(pdf_folder, output_excel_path):
    all_results = []

    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            print(f"Processing: {filename}")
            
            # Run benchmarks for each method and store the results
            for method, method_name in [
                (extract_links_re, "Regular Expressions (re)"),
                (extract_links_flashtext, "FlashText"),
                (extract_links_ahocorasick, "Aho-Corasick")
            ]:
                result = benchmark_method(method, method_name, pdf_path)
                result["Filename"] = filename  # Add the filename to the results
                all_results.append(result)

    # Convert the results to a DataFrame
    df = pd.DataFrame(all_results)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

# Specify the folder containing the PDF resumes and the output Excel file path
pdf_folder = r'C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes'
output_excel_path = r'links_benchmark_results.xlsx'

# Process the PDFs and save the results to an Excel file
process_pdfs_and_save_to_excel(pdf_folder, output_excel_path)
