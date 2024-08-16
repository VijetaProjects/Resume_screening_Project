
# Extract text using PyPDF2

import os
import pandas as pd
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
from tika import parser

# Define the directory containing the PDF resumes
directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"

# List to store the extracted text and filenames
extracted_data = []

# Function to extract text using PyPDF2
def extract_text_pypdf2(file_path):
    try:
        reader = PdfReader(file_path)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        print(f"PyPDF2 failed for {file_path}: {e}")
        return None

# Function to extract text using pdfminer
def extract_text_pdfminer(file_path):
    try:
        text = extract_text(file_path)
        return text
    except Exception as e:
        print(f"PDFMiner failed for {file_path}: {e}")
        return None

# Function to extract text using Tika
def extract_text_tika(file_path):
    try:
        parsed = parser.from_file(file_path)
        return parsed['content']
    except Exception as e:
        print(f"Tika failed for {file_path}: {e}")
        return None

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(directory, filename)
        
        # Try to extract text using different methods
        text = extract_text_pypdf2(file_path)
        
        # If PyPDF2 fails, try PDFMiner
        if text is None or text.strip() == "":
            text = extract_text_pdfminer(file_path)
        
        # If PDFMiner fails, try Tika
        if text is None or text.strip() == "":
            text = extract_text_tika(file_path)
        
        # If any text was successfully extracted, add it to the list
        if text and text.strip():
            extracted_data.append({'filename': filename, 'text': text})
        else:
            print(f"Failed to extract text from {filename}")

# Create a DataFrame from the extracted data
df = pd.DataFrame(extracted_data)

# Specify the output Excel file
output_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_pypdf1.xlsx"

# Write the DataFrame to an Excel file
df.to_excel(output_excel_file, index=False)
print(f"Data successfully written to {output_excel_file}")
