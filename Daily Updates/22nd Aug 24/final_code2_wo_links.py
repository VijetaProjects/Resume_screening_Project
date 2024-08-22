import pandas as pd
import re
from ollama import Client
import time

# Initialize the LLM client
client = Client()

# Function to clean text using regular expressions for better formatting
def clean_text_column(text):
    if isinstance(text, str):
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        text = re.sub(r'(\d)([A-Z])', r'\1 \2', text)
        text = re.sub(r'([A-Z])([A-Z][a-z])', r'\1 \2', text)
        text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
    return text

# Function to clean the 'Extracted Text' column in a DataFrame
def clean_extracted_text_column(df, column_name='Extracted Text'):
    df[column_name] = df[column_name].apply(clean_text_column)
    return df

# Function to extract information using LLM prompt engineering
def extract_information_llm(resume_text):
    if not isinstance(resume_text, str):
        return {"Name": "Not mentioned", "Location": "Not mentioned"}

    prompt = f"""
    Extract the name and location from the following resume text. The name may appear anywhere in the text and may be capitalized.

    Resume Text: {resume_text}

    Provide the extracted information in this format:
    - Name: [Extracted Name]
    - Location: [Extracted Location]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    extracted_info = {
        "Name": "Not mentioned",
        "Location": "Not mentioned"
    }

    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip() or "Not mentioned"
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip() or "Not mentioned"
    
    return extracted_info

# Function to extract phone number using LLM prompt engineering
def extract_phone_number(resume_text):
    if not isinstance(resume_text, str):
        return "Not mentioned"

    prompt = f"""
    Extract the phone number from the following resume text. The phone number may appear in various formats, such as international or local.

    Resume Text: {resume_text}

    Provide the extracted phone number in this format:
    - Phone Number: [Extracted Phone Number]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    phone_number = "Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Phone Number:" in line:
            extracted_value = line.split(":", 1)[1].strip()
            if extracted_value.lower() == "none" or not extracted_value:
                phone_number = "Not mentioned"
            else:
                phone_number = extracted_value
    
    return phone_number

# Function to generate a fitment summary using LLM
def fitment_summary(resume_text, job_description):
    if not isinstance(resume_text, str) or not isinstance(job_description, str):
        return "Not mentioned"
    
    prompt = f"""
    Based on the following resume and job description, provide a summary of how this candidate is suitable for the role in 50 words.

    Resume Text: {resume_text}

    Job Description: {job_description}

    Provide the fitment summary in this format:
    - Summary: [50-word Summary]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    summary = "Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Summary:" in line:
            summary = line.split(":", 1)[1].strip() or "Not mentioned"
    
    return summary

# Function to calculate total experience using LLM
def total_experience(resume_text):
    if not isinstance(resume_text, str):
        return "Fresher or Not mentioned"

    prompt = f"""
    Calculate the total years of experience from the work experience section of the following resume text. If no work experience section is available, return 'Fresher or Not mentioned.'

    Resume Text: {resume_text}

    Provide the total experience in this format:
    - Experience: [Total Experience in Years]
    """

    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    experience = "Fresher or Not mentioned"
    
    lines = extracted_text.split('\n')
    for line in lines:
        if "Experience:" in line:
            experience = line.split(":", 1)[1].strip() or "Fresher or Not mentioned"
    
    return experience

# Function to calculate a suitability score using LLM
def calculate_score(resume_text, job_description):
    if not isinstance(resume_text, str) or not isinstance(job_description, str):
        return "Not mentioned"

    prompt = f"""
    Based on the job description below, evaluate the suitability of the candidate described in the resume. 
    Provide a score from 1 to 100 on how suitable the candidate is for the job role.

    Job Description:
    {job_description}

    Candidate Resume:
    {resume_text}

    Please respond with only the numerical score.
    """

    # Query the Llama3 model using the Ollama client
    try:
        response = client.generate(model="llama3:latest", prompt=prompt)
        score = int(response['response'].strip())  # Extract the score from the model's response
    except (ValueError, KeyError):
        score = None  # Handle cases where the response is not a valid integer

    return score

# Main function to process the resumes from an Excel file and output to a new Excel file
def process_resumes_with_llm(excel_file, job_description_file, output_excel_file):
    # Load the Excel file with resume data
    df = pd.read_excel(excel_file)
    
    # Read the job description from the text file
    with open(job_description_file, 'r') as file:
        job_description = file.read()
    
    # Clean the 'Extracted Text' column
    df = clean_extracted_text_column(df)
    
    # Initialize lists to store extracted information
    names = []
    locations = []
    phone_numbers = []
    experiences = []
    summaries = []
    scores = []

    # Loop through each resume in the DataFrame
    for index, row in df.iterrows():
        resume_text = row['Extracted Text']
        print(f"Processing resume: {row['Filename']}")  # Print progress
        
        # Extract information using LLM
        extracted_info = extract_information_llm(resume_text)
        phone_number = extract_phone_number(resume_text)
        experience = total_experience(resume_text)
        summary = fitment_summary(resume_text, job_description)
        score = calculate_score(resume_text, job_description)
        
        # Append the extracted information to the lists
        names.append(extracted_info["Name"])
        locations.append(extracted_info["Location"])
        phone_numbers.append(phone_number)
        experiences.append(experience)
        summaries.append(summary)
        scores.append(score)
    
    # Add the extracted information as new columns in the DataFrame
    df['Name'] = names
    df['Location'] = locations
    df['Phone Number'] = phone_numbers
    df['Total Experience'] = experiences
    df['Fitment Summary'] = summaries
    df['Suitability Score'] = scores
    
    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_excel_file, index=False)

    print(f"Processing complete. Data saved to {output_excel_file}")

# Example usage
excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\22 Aug\100 resumes_text.xlsx"  # Path to your Excel file with resumes
job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"  # Path to your job description text file
output_excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\22 Aug\processed_100resumes_with_info_wolinks.xlsx"  # Path to your output Excel file

# Call the function to process the resumes and store the final data in an Excel file
process_resumes_with_llm(excel_file, job_description_file, output_excel_file)
