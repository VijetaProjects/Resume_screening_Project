import os
import re
import pandas as pd
from ollama import Client

# Initialize the client
client = Client()

# Function to extract information from text using prompt engineering
def extract_information(resume_text):
    prompt = f"""
    The following is a resume text. Extract and list the following information if available:
    1. Name
    2. Location
    3. Phone Number

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
        "Phone Number": ""
    }

    # Parse the extracted text
    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip()
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip()
        elif "Phone Number:" in line:
            extracted_info["Phone Number"] = line.split(":", 1)[1].strip()

    return extracted_info

# Function to read data from Excel sheet and extract relevant information
def process_excel_file(input_file):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # List to store the extracted information
    extracted_data = []

    # Loop through each resume in the DataFrame
    for index, row in df.iterrows():
        resume_text = row['text']
        extracted_info = extract_information(resume_text)
        extracted_data.append(extracted_info)

    return extracted_data

# Function to save extracted data to a new Excel file
def save_to_excel(data, output_file):
    # Convert the list of extracted info into a DataFrame
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    df.to_excel(output_file, index=False)

# Input Excel file with resume text
input_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_docx.xlsx"

# Output Excel file to save extracted information
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\extracted_info_docx.xlsx"

# Process resumes and save the extracted data to a new Excel file
extracted_data = process_excel_file(input_file)
save_to_excel(extracted_data, output_file)

print(f"Data saved to {output_file}")
