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

# Function to generate skill requirements from job description
def generate_requirements(skill, job_description):
    prompt = f"Given the following job description, explain the skill requirement for the role '{skill}'.\n\nJob Description:\n{job_description}"
    
    # Generate the response
    response = client.generate(model='llama3:latest', prompt=prompt)
    
    # Access the response content using the correct key
    return response.get('response', "Error: Requirement generation failed.")

# Function to evaluate candidate's resume against the extracted skills
def evaluate_candidate(skill, resume_text):
    prompt = f"How does the candidate have alignment with the skill and requirement '{skill}'?\n\nCandidate Resume:\n{resume_text}"
    
    # Generate the response
    response = client.generate(model='llama3:latest', prompt=prompt)
    
    # Access the response content using the correct key
    return response.get('response', "Error: Candidate observation generation failed.")

# Load job description from txt file
with open(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt", 'r') as file:
    job_description = file.read()

# Load resume data from Excel file
resume_df = pd.read_excel(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx")

# Select only the first record from the DataFrame
first_record = resume_df.iloc[0]

# Extract skills from job description
skills = extract_skills(job_description)

# Extract resume text from the first record
resume_text = first_record['Extracted Text']

# Initialize lists to store requirements and candidate observations
requirements = []
observations = []

# Generate requirements and candidate observations for each skill
for skill in skills:
    requirement = generate_requirements(skill, job_description)
    observation = evaluate_candidate(skill, resume_text)
    
    requirements.append(requirement)
    observations.append(observation)

# Create an evaluation sheet (Pandas DataFrame) for this candidate
evaluation_sheet = pd.DataFrame({
    'Evaluation Parameter': skills,
    'Requirement': requirements,
    'Candidate Observation': observations
})

# Define the file path for saving the evaluation sheet
output_file = "evaluation_sheet_first_record3.xlsx"

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
