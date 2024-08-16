import os
import pandas as pd
from PyPDF2 import PdfReader

# Specify the directory containing the PDF resumes
directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"

# List to store the extracted text and filenames
extracted_data = []

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.pdf'):
        file_path = os.path.join(directory, filename)
        
        # Extract text from the PDF file
        try:
            reader = PdfReader(file_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            
            # Append the extracted text and filename to the list
            extracted_data.append({'filename': filename, 'text': text})
        
        except Exception as e:
            print(f"Error reading {filename}: {e}")

# Create a DataFrame from the extracted data
df = pd.DataFrame(extracted_data)

# Specify the output Excel file
output_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_pypdf.xlsx"

# Write the DataFrame to an Excel file
df.to_excel(output_excel_file, index=False)
print(f"Data successfully written to {output_excel_file}")
