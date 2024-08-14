
# In this code separate prompts are used. 

import pdfplumber
import os
import pandas as pd
from ollama import Client

# Initialize the client
client = Client()

# Define functions to create specific prompts for each piece of information

def create_name_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the full name of the candidate from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the full name.
    """

def create_location_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the location (city, state, and/or country) where the candidate is based from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the location.
    """

def create_github_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the GitHub profile link from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the GitHub profile link.
    """

def create_linkedin_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the LinkedIn profile link from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the LinkedIn profile link.
    """

def create_phone_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the phone number from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the phone number.
    """

def create_experience_prompt(resume_text):
    return f"""
    The following is a resume text. Extract the total years of experience the candidate has from the resume text.

    Resume Text:
    {resume_text}
    
    Please provide only the total years of experience.
    """

# Function to extract text from a PDF using pdfplumber
def extract_text_pdfplumber(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

# Function to extract information from a resume using specific prompts
def extract_information_from_resume(resume_text):
    # Generate prompts for each piece of information
    name_prompt = create_name_prompt(resume_text)
    location_prompt = create_location_prompt(resume_text)
    github_prompt = create_github_prompt(resume_text)
    linkedin_prompt = create_linkedin_prompt(resume_text)
    phone_prompt = create_phone_prompt(resume_text)
    experience_prompt = create_experience_prompt(resume_text)
    
    # Extract information using the ollama client
    name = client.generate(model="llama3:latest", prompt=name_prompt).get('response', 'No name found.')
    location = client.generate(model="llama3:latest", prompt=location_prompt).get('response', 'No location found.')
    github = client.generate(model="llama3:latest", prompt=github_prompt).get('response', 'No GitHub link found.')
    linkedin = client.generate(model="llama3:latest", prompt=linkedin_prompt).get('response', 'No LinkedIn link found.')
    phone = client.generate(model="llama3:latest", prompt=phone_prompt).get('response', 'No phone number found.')
    experience = client.generate(model="llama3:latest", prompt=experience_prompt).get('response', 'No experience found.')
    
    # Return a structured dictionary with the extracted information
    return {
        "Name": name.strip(),
        "Location": location.strip(),
        "GitHub Link": github.strip(),
        "LinkedIn Link": linkedin.strip(),
        "Phone Number": phone.strip(),
        "Total Experience": experience.strip()
    }

# Function to process multiple resumes
def process_resumes(directory_path):
    data = []
    
    for filename in os.listdir(directory_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory_path, filename)
            resume_text = extract_text_pdfplumber(pdf_path)
            extracted_info = extract_information_from_resume(resume_text)
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
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\extracted_info_sepearte_prompts.xlsx"

# Process resumes and save the extracted data to Excel
data = process_resumes(directory_path)
save_to_excel(data, output_file)

print(f"Data saved to {output_file}")
