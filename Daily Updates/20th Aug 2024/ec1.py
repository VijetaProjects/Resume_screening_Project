import pandas as pd
import ollama
import os

# Initialize the client (without specifying the model here)
client = ollama.Client()

# Function to extract skills from job description
def extract_skills(job_description):
    prompt = f"Extract a list of technical skills required for the following job description:\n\n{job_description}"
    
    # Generate the response
    response = client.generate(model='llama3:latest', prompt=prompt)
    
    # Print the response for debugging purposes
    print("Response from API:", response)
    
    # Access the response content using the correct key
    if 'response' in response:
        # Extract skills by splitting the response by newline and filtering out unnecessary lines
        skills = response['response'].split('\n')
        # Remove any empty lines and strip numbering from skills
        skills = [skill.split('. ', 1)[-1].strip() for skill in skills if skill.strip() and not skill.startswith("Note")]
        return skills
    else:
        raise KeyError("The key 'response' was not found in the response.")

# Function to evaluate candidate's resume against the extracted skills
def evaluate_candidate(skills, resume_text):
    observations = []
    for skill in skills:
        prompt = f"Does the following resume contain the skill '{skill}'?\n\nResume:\n{resume_text}"
        
        # Generate the response
        response = client.generate(model='llama3:latest', prompt=prompt)
        
        # Check if 'response' exists in the response
        if 'response' in response:
            observations.append(response['response'].strip())
        else:
            observations.append(f"Error: 'response' not found in response for skill '{skill}'")
    
    return observations

# Load job description from txt file
with open(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt", 'r') as file:
    job_description = file.read()

# Load resume data from Excel file
resume_df = pd.read_excel(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx")

# Select only the first record from the DataFrame
first_record = resume_df.iloc[0]

# Extract skills from job description
skills = extract_skills(job_description)

# Extract filename and resume text from the first record
filename = first_record['Filename']
resume_text = first_record['Extracted Text']

# Evaluate the candidate's resume based on the extracted skills
observations = evaluate_candidate(skills, resume_text)

# Create an evaluation sheet (Pandas DataFrame) for this candidate
evaluation_sheet = pd.DataFrame({
    'Filename': [filename] * len(skills),
    'Evaluation Parameter': skills,
    'Requirement': [f"Proficiency in {skill}" for skill in skills],
    'Candidate Observation': observations
})

# Define the file path for saving the evaluation sheet
output_file = "evaluation_sheet_first_record1.xlsx"

try:
    # Save the evaluation sheet to an Excel file
    evaluation_sheet.to_excel(output_file, index=False)
    print(f"Evaluation sheet for the first record has been saved to '{output_file}'.")
    
    # Check if the file was actually created
    if os.path.exists(output_file):
        print(f"File saved successfully at: {os.path.abspath(output_file)}")
    else:
        print("Error: File was not saved.")
except Exception as e:
    print(f"An error occurred while saving the Excel file: {e}")
