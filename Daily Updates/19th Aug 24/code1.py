import time  # Import the time module
import pandas as pd
from ollama import Client

# Initialize the LLM client
client = Client()

# Load the Excel file containing extracted resume text
input_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx"
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\19 Aug\name_location_extraction_sheet1.xlsx"

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# Define a function to extract information using LLM prompt engineering
def extract_information_llm(resume_text):
    # Skip if the text is not a string
    if not isinstance(resume_text, str):
        return {"Name": "", "Location": ""}

    # Improved prompt for extracting name and location, handling complex cases
    prompt = f"""
    The following text contains semi-structured information related to personal details, education, work experience, projects, skills, and additional information. Your task is to extract key personal information.

    Resume Text:
    {resume_text}
    
    Task:
    Extract the following categories:
    1. Personal Information: Name and location. The name may be embedded within the text, capitalized, or surrounded by other text. Ensure that you accurately identify and extract the full name by recognizing patterns like capitalization and common formatting around names. For example, if the name appears in the middle of the text or is attached to other information, ensure only the proper name is extracted.
    
    Organize the information: Format the extracted information into a structured and readable format with proper headings.
    
    Output Format:

    - Name: [Extracted Name]
    - Location: [Extracted Location]

    Provide the extracted information clearly.
    """

    # Use the LLM model to generate a response
    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    # Initialize a dictionary to store the extracted information
    extracted_info = {
        "Name": "",
        "Location": ""
    }

    # Parse the extracted text for the required fields
    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip()
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip()
    
    return extracted_info

# Record the start time
start_time = time.time()

# Create a list to store the extracted information
extracted_data = []

# Process each row in the DataFrame to extract the required information
for index, row in df.iterrows():
    resume_text = row.get('Extracted Text', "")  # Assuming the resume text is in this column
    extracted_info = extract_information_llm(resume_text)
    extracted_info["Filename"] = row.get('Filename', "")  # Add filename to the extracted data
    extracted_data.append(extracted_info)

# Convert the list of extracted info into a DataFrame
extracted_df = pd.DataFrame(extracted_data)

# Save the extracted information to a new Excel file
extracted_df.to_excel(output_file, index=False)

# Record the end time
end_time = time.time()

# Calculate the duration
duration = end_time - start_time

print(f"Extracted data saved to {output_file}")
print(f"Time taken: {duration:.2f} seconds")
