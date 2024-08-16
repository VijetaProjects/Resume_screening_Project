import pdfplumber
import pandas as pd
import os

# Function to clean the extracted text by removing non-printable characters
def clean_text(text):
    # Keep only printable characters
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

# Function to extract text from multiple PDFs and store it in an Excel sheet
def pdfs_to_excel(pdf_folder, excel_path):
    all_text_data = []
    
    # Loop through all PDF files in the specified folder
    for filename in os.listdir(pdf_folder):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            print(f"Extracting text from: {filename}")
            
            # Extract text from the current PDF and append it to the list
            pdf_text = extract_text_from_pdf(pdf_path)
            all_text_data.append({"Filename": filename, "Extracted Text": "\n".join(pdf_text)})
    
    # Create a pandas DataFrame from the collected text data
    df = pd.DataFrame(all_text_data)
    
    # Save the DataFrame to an Excel file
    df.to_excel(excel_path, index=False)



# Specify the folder containing the PDFs and the output Excel file path
pdf_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"

excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_pdfplumber.xlsx"

# Call the function to extract text from all PDFs and store it in the Excel file
pdfs_to_excel(pdf_folder, excel_file)
