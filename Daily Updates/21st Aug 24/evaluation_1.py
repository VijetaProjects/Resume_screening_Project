import pandas as pd
import ollama
import os

# Initialize the client (without specifying the model here)
client = ollama.Client()

# Function to evaluate candidate's resume against the extracted skills
def evaluate_candidate(skill, resume_text):
    prompt = f"Does the candidate have the skill '{skill}'? If so, briefly describe their experience with it.\n\nCandidate Resume:\n{resume_text}"
    
    # Generate the response
    response = client.generate(model='llama3:latest', prompt=prompt)
    
    # Access the response content using the correct key
    return response.get('response', "Error: Candidate observation generation failed.")

# Load skills and requirements from the provided Excel sheet
skills_df = pd.read_excel(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\21 Aug\Final Code 0\Evaluation Criteria Sheet.xlsx")  # Replace with actual path to the Excel sheet

# Print the column names to debug
print("Columns in skills_df:", skills_df.columns)

# Load resume data from Excel file
resume_df = pd.read_excel(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx")

# Select only the first record from the DataFrame
first_record = resume_df.iloc[0]

# Extract resume text from the first record
resume_text = first_record['Extracted Text']

# Initialize a list to store candidate observations
observations = []

# Generate candidate observations for each skill
# Ensure the correct column name is used here based on the printed column names
for skill in skills_df['Skills']:  # If the column name is different, update this line accordingly
    observation = evaluate_candidate(skill, resume_text)
    observations.append(observation)

# Add candidate observations to the DataFrame
skills_df['Candidate Observation'] = observations

# Define the file path for saving the evaluation sheet
output_file = "evaluation_sheet_first_record_updated.xlsx"

try:
    # Save the evaluation sheet to an Excel file
    skills_df.to_excel(output_file, index=False)
    print(f"Evaluation sheet for the first record has been saved to '{output_file}'.")
    
    # Check if the file was actually created
    if os.path.exists(output_file):
        print(f"File saved successfully at: {os.path.abspath(output_file)}")
    else:
        print("Error: File was not saved.")
except Exception as e:
    print(f"An error occurred while saving the Excel file: {e}")
