import os
import pdfplumber
import pandas as pd
import openpyxl

# Function to clean the extracted text by removing illegal characters for Excel
def clean_text(text):
    # Define illegal characters that are not allowed in Excel cells
    illegal_characters = [
        chr(0), chr(1), chr(2), chr(3), chr(4), chr(5), chr(6), chr(7), chr(8), chr(9),
        chr(10), chr(11), chr(12), chr(13), chr(14), chr(15), chr(16), chr(17), chr(18), chr(19),
        chr(20), chr(21), chr(22), chr(23), chr(24), chr(25), chr(26), chr(27), chr(28), chr(29),
        chr(30), chr(31)
    ]
    for char in illegal_characters:
        text = text.replace(char, "")
    return text

# Function to extract text from a single PDF file using pdfplumber
def extract_text_from_pdf(pdf_path):
    text_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                cleaned_text = clean_text(page_text)  # Clean the extracted text
                text_data.append(cleaned_text)
    return "\n".join(text_data)

# Function to process all PDFs in a directory and save the extracted text to an Excel file
def extract_text_from_pdfs_to_excel(pdf_directory, output_excel_file):
    # List to hold all the extracted data
    extracted_data = []

    # Loop through all PDF files in the specified directory
    for filename in os.listdir(pdf_directory):
        if filename.endswith(".pdf"):
            file_path = os.path.join(pdf_directory, filename)
            print(f"Processing: {filename}")
            
            try:
                # Extract text from the PDF file
                extracted_text = extract_text_from_pdf(file_path)
                
                # Append the filename and extracted text to the list
                extracted_data.append({"Filename": filename, "Extracted Text": extracted_text})
            
            except Exception as e:
                # Print an error message and skip to the next file
                print(f"Error processing {filename}: {e}")
                continue

    # Convert the list of extracted data to a pandas DataFrame
    df = pd.DataFrame(extracted_data)

    # Save the DataFrame to an Excel file
    df.to_excel(output_excel_file, index=False)

    print(f"Extraction complete. Data saved to {output_excel_file}")

# Example usage
pdf_directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\100 Resumes"  # Path to the directory containing the PDF resumes
output_excel_file = r"100 resumes_text.xlsx"  # Path to the output Excel file

extract_text_from_pdfs_to_excel(pdf_directory, output_excel_file)
