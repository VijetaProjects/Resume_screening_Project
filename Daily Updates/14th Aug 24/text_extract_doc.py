import os
import pandas as pd
from docx import Document
from pdf2docx import Converter
from ollama import Client

# Initialize the client
client = Client()

# Function to convert PDF to DOCX using pdf2docx
def convert_pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

# Function to extract text from DOCX
def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# Define a function to extract information using prompt engineering
def extract_information(resume_text):
    prompt = f"""
    The following is a resume text. Extract and list the following information if available:
    1. Name
    2. Location
    3. GitHub Link
    4. LinkedIn Link
    5. Phone Number
    6. Total Experience
    
    Resume Text:
    {resume_text}
    
    Please provide the extracted information in a clear and concise manner.
    """
    
    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    extracted_info = {
        "Name": "",
        "Location": "",
        "GitHub Link": "",
        "LinkedIn Link": "",
        "Phone Number": "",
        "Total Experience": ""
    }

    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip()
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip()
        elif "GitHub Link:" in line:
            extracted_info["GitHub Link"] = line.split(":", 1)[1].strip()
        elif "LinkedIn Link:" in line:
            extracted_info["LinkedIn Link"] = line.split(":", 1)[1].strip()
        elif "Phone Number:" in line:
            extracted_info["Phone Number"] = line.split(":", 1)[1].strip()
        elif "Total Experience:" in line:
            extracted_info["Total Experience"] = line.split(":", 1)[1].strip()
    
    return extracted_info

# Function to process multiple resumes
def process_resumes(directory_path, temp_docx_dir):
    data = []
    
    for filename in os.listdir(directory_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory_path, filename)
            docx_path = os.path.join(temp_docx_dir, filename.replace('.pdf', '.docx'))
            
            convert_pdf_to_docx(pdf_path, docx_path)
            resume_text = extract_text_from_docx(docx_path)
            extracted_info = extract_information(resume_text)
            data.append(extracted_info)
    
    return data

# Function to save data to an Excel file
def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

# Directory containing resumes
directory_path = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
temp_docx_dir = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\TempDocx"
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\extracted_info_bydoc.xlsx"

if not os.path.exists(temp_docx_dir):
    os.makedirs(temp_docx_dir)

data = process_resumes(directory_path, temp_docx_dir)
save_to_excel(data, output_file)

print(f"Data saved to {output_file}")
