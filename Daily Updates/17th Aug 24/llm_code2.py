import pandas as pd
from ollama import Client

# Initialize the LLM client
client = Client()

# Load the Excel file containing extracted resume text
input_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx"
output_file = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\clean_text_llm1.xlsx"

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(input_file)

# Define a function to extract information using LLM prompt engineering
def extract_information_llm(resume_text):
    # Skip if the text is not a string
    if not isinstance(resume_text, str):
        return {"Name": "", "Location": "", "Phone Number": ""}

    # Enhanced prompt to handle cases with no spaces, missing keywords, and heavily attached text
    prompt = f"""
    The following is a resume text. The text might be compressed, lacking spaces, or missing clear keywords. Please carefully infer the following information:
    1. Name: Look for potential names, considering capitalization patterns and common name structures.
    2. Location: Try to identify location-related information based on typical address patterns (e.g., city names, postal codes, etc.).
    3. Phone Number: Look for numeric sequences that resemble phone numbers.

    If the text is too compressed, use your best judgment to infer these details based on patterns like capitalization and numeric sequences.

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
