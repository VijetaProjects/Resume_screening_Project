# In below code Used different dataset

import pandas as pd
from ollama import Client

df = pd.read_csv("hf://datasets/brackozi/Resume/UpdatedResumeDataSet.csv")

# Initialize the client
client = Client()

# Define the specific prompts for each field
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

# Function to extract information from a single resume
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


# Process each resume and extract information
extracted_data = []
for index, row in df.iterrows():
    resume_text = row['Resume']
    extracted_info = extract_information_from_resume(resume_text)
    extracted_data.append(extracted_info)

# Convert extracted data to a DataFrame
extracted_df = pd.DataFrame(extracted_data)

# Save to an Excel file
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\14 Aug\extracted_resume_info.xlsx"
extracted_df.to_excel(output_file, index=False)

print(f"Data saved to {output_file}")
