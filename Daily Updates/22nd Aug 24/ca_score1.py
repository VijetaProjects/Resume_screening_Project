import pandas as pd
import ollama

def calculate_job_fit_score(candidate_resume, job_description):
    # Prepare the prompt for Llama3
    prompt = f"""
    Based on the job description below, evaluate the suitability of the candidate described in the resume. 
    Provide a score from 1 to 100 on how suitable the candidate is for the job role.

    Job Description:
    {job_description}

    Candidate Resume:
    {candidate_resume}

    Please respond with only the numerical score.
    """

    # Query the Llama3 model using the Ollama client
    try:
        # Use the correct method for generating the response
        response = ollama.generate(model="llama3:latest", prompt=prompt)
        
        # Extract the score from the model's response
        score = int(response['response'].strip())  # 'response' key contains the generated text
    except (ValueError, KeyError):
        score = None  # Handle cases where the response is not a valid integer

    return score

def process_resumes(excel_file, job_description_file, output_file):
    # Load the Excel file with resume data
    df = pd.read_excel(excel_file)

    # Read the job description from the text file
    with open(job_description_file, 'r') as file:
        job_description = file.read()

    # Initialize a list to store scores
    scores = []

    # Loop through each resume in the DataFrame
    for index, row in df.iterrows():
        resume_text = row['Extracted Text']
        print(f"Processing resume: {row['Filename']}")  # Optional: Print progress
        score = calculate_job_fit_score(resume_text, job_description)
        
        # Append the score (or a default value if None) to the scores list
        scores.append(score if score is not None else "Score Not Available")

    # Add the scores as a new column in the DataFrame
    df['Suitability Score'] = scores

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_file, index=False)

    print(f"Scores have been calculated and saved to {output_file}")

# Example usage
excel_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx"  # Path to your Excel file with resumes
job_description_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Role Description.txt"  # Path to your job description text file
output_file = "resumes_with_scores1.xlsx"  # Output Excel file with scores

process_resumes(excel_file, job_description_file, output_file)
