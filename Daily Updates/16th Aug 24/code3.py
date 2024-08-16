
# extract text , convert to pdf to doc

import os
import pandas as pd
from docx import Document

# Define the directory containing the DOCX resumes
directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\TempDocx"

# List to store the extracted text and filenames
extracted_data = []

# Function to extract text from a DOCX file using python-docx
def extract_text_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        print(f"Failed to extract text from {file_path}: {e}")
        return None

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.docx'):
        file_path = os.path.join(directory, filename)
        
        # Extract text from the DOCX file
        text = extract_text_docx(file_path)
        
        # If any text was successfully extracted, add it to the list
        if text and text.strip():
            extracted_data.append({'filename': filename, 'text': text})
        else:
            print(f"Failed to extract text from {filename}")

# Create a DataFrame from the extracted data
df = pd.DataFrame(extracted_data)

# Specify the output Excel file
output_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_docx.xlsx"

# Write the DataFrame to an Excel file
df.to_excel(output_excel_file, index=False)
print(f"Data successfully written to {output_excel_file}")
