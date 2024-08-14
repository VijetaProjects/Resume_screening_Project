
# Use Pdfplumber for text extraction

import pdfplumber
import os
import pandas as pd
from ollama import Client
import time

# Start time measurement
start_time = time.time()

# Initialize the client
client = Client()

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
    
    # Use the generate function to specify the model and pass the prompt
    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    # Split the extracted information into a structured dictionary
    extracted_info = {
        "Name": "",
        "Location": "",
        "GitHub Link": "",
        "LinkedIn Link": "",
        "Phone Number": "",
        "Total Experience": ""
    }

    # Parse the extracted text
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

# Function to extract text from a PDF using pdfplumber
def extract_text_pdfplumber(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Function to process multiple resumes
def process_resumes(directory_path):
    data = []
    
    for filename in os.listdir(directory_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory_path, filename)
            resume_text = extract_text_pdfplumber(pdf_path)
            extracted_info = extract_information(resume_text)
            data.append(extracted_info)
    
    return data

# Function to save data to an Excel file
def save_to_excel(data, output_file):
    # Convert the list of extracted info into a DataFrame
    df = pd.DataFrame(data)
    
    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)

# Directory containing resumes
directory_path = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\extracted_info_pdfplumber.xlsx"

# Process resumes and save the extracted data to Excel
data = process_resumes(directory_path)
save_to_excel(data, output_file)

print(f"Data saved to {output_file}")
# End time measurement
end_time = time.time()

# Calculate response time
response_time = end_time - start_time

# Print the response time
print(f"Response Time: {response_time:.2f} seconds")
