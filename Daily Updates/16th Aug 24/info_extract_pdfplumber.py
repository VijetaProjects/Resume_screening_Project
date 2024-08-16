import pandas as pd
from ollama import Client

# Initialize the LLM client
client = Client()

# Load the Excel file containing extracted resume text
input_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_pdfplumber.xlsx"
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\extracted_resume_info_llm.xlsx"

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# Define a function to extract information using LLM prompt engineering
def extract_information_llm(resume_text):
    # Skip if the text is not a string
    if not isinstance(resume_text, str):
        return {"Name": "", "Location": "", "Phone Number": ""}

    # Create the prompt for the LLM model
    prompt = f"""
    The following is a resume text. Extract and list the following information if available:
    1. Name
    2. Location
    3. Phone Number
    
    Resume Text:
    {resume_text}
    
    Please provide the extracted information in a clear and concise manner.
    """

    # Use the LLM model to generate a response
    response = client.generate(model="llama3:latest", prompt=prompt)
    
    extracted_text = response.get('response', 'No response text found.')
    
    # Split the extracted information into a structured dictionary
    extracted_info = {
        "Name": "",
        "Location": "",
        "Phone Number": ""
    }

    # Parse the extracted text for the required fields
    lines = extracted_text.split('\n')
    for line in lines:
        if "Name:" in line:
            extracted_info["Name"] = line.split(":", 1)[1].strip()
        elif "Location:" in line:
            extracted_info["Location"] = line.split(":", 1)[1].strip()
        elif "Phone Number:" in line:
            extracted_info["Phone Number"] = line.split(":", 1)[1].strip()
    
    return extracted_info

# Create a list to store the extracted information
extracted_data = []

# Process each row in the DataFrame to extract the required information
for index, row in df.iterrows():
    resume_text = row['Extracted Text']  # Assuming the resume text is in this column
    extracted_info = extract_information_llm(resume_text)
    extracted_info["Filename"] = row['Filename']  # Add filename to the extracted data
    extracted_data.append(extracted_info)

# Convert the list of extracted info into a DataFrame
extracted_df = pd.DataFrame(extracted_data)

# Save the extracted information to a new Excel file
extracted_df.to_excel(output_file, index=False)

print(f"Extracted data saved to {output_file}")
