import PyPDF2
import os
import ollama
from ollama import Client

# Initialize the client without specifying the model
client = Client()

# Define a function to extract information using prompt engineering
def extract_information(resume_text):
    prompt = f"""
    The following is a resume text. Extract and list the following information if available:
    1. Name
    2. Location
    3. GitHub Link
    4. LinkedIn Link
    5. Phone Number
    
    Resume Text:
    {resume_text}
    
    Please provide the extracted information in a clear and concise manner.
    """
    
    # Use the generate function to specify the model and pass the prompt
    response = client.generate(model="llama3:latest", prompt=prompt)
    
    # Print the entire response to understand its structure
    #print("Full Response:", response)
    
    # Return the extracted information from the 'response' key
    return response.get('response', 'No response text found.')

def extract_text_pypdf(pdf_path):
  with open(pdf_path, 'rb') as pdf_file:
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    num_pages = len(pdf_reader.pages)
    text = ""
    for page_num in range(num_pages):
      page = pdf_reader.pages[page_num]
      text += page.extract_text()
    return text

pdf_path = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Vijeta _V _Wasnik Resume.pdf"

resume_text = extract_text_pypdf(pdf_path)
extracted_info = extract_information(resume_text)
print(extracted_info)

