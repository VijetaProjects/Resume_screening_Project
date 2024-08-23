import pdfplumber
import fitz  # PyMuPDF
import os
import time
import pandas as pd

# Function to extract links using pdfplumber
def extract_links_pdfplumber(pdf_path):
    links = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        links.append(annot['uri'])
    return [link for link in links if link]  # Filter out None values

# Function to extract links using PyMuPDF (fitz)
def extract_links_fitz(pdf_path):
    links = []
    pdf_document = fitz.open(pdf_path)
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        links.extend([link['uri'] for link in page.get_links() if 'uri' in link])
    return [link for link in links if link]  # Filter out None values

# Function to compare the speed and accuracy of both methods and store the results in an Excel sheet
def compare_extraction_methods_and_save(resume_folder, output_excel_path):
    results = []

    for filename in os.listdir(resume_folder):
        if filename.endswith(".pdf"):
            file_path = os.path.join(resume_folder, filename)
            print(f"Processing: {filename}")

            # Extracting links using pdfplumber
            start_pdfplumber = time.time()
            links_pdfplumber = extract_links_pdfplumber(file_path)
            end_pdfplumber = time.time()

            # Extracting links using PyMuPDF (fitz)
            start_fitz = time.time()
            links_fitz = extract_links_fitz(file_path)
            end_fitz = time.time()

            # Compare the results
            same_links = set(links_pdfplumber) == set(links_fitz)
            pdfplumber_time = end_pdfplumber - start_pdfplumber
            fitz_time = end_fitz - start_fitz

            # Collect results for this PDF
            results.append({
                "Filename": filename,
                "pdfplumber_Links": ", ".join(links_pdfplumber) if links_pdfplumber else "No links",
                "fitz_Links": ", ".join(links_fitz) if links_fitz else "No links",
                "pdfplumber_Time (s)": round(pdfplumber_time, 4),
                "fitz_Time (s)": round(fitz_time, 4),
                "Same_Links": same_links
            })

    # Convert results to a DataFrame and save to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

# Specify the folder containing the PDFs and the output Excel file path
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
output_excel_path = r"extracted_links_comparison.xlsx"

# Compare the extraction methods on the PDFs and save the results to an Excel file
compare_extraction_methods_and_save(resume_folder, output_excel_path)
