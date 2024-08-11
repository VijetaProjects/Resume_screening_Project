import os
import pdfplumber
import spacy
import pandas as pd

# Initialize spaCy model
nlp = spacy.load("en_core_web_sm")

# Folder containing resumes in PDF format
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
# List to hold extracted data
extracted_data = []

# Function to extract text from PDF using pdfplumber
def extract_text_from_pdf(pdf_path):
    resume_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            resume_text += page.extract_text()
    return resume_text

# Function to extract details using spaCy
def extract_details(text):
    doc = nlp(text)
    details = {
        "Name": None,
        "Location": None,
        "Phone Number": None,
        "Email": None,
        "GitHub Link": None,
        "LinkedIn Link": None
    }

    # Extract entities
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            details["Name"] = ent.text
        elif ent.label_ == "GPE":
            details["Location"] = ent.text
        # You can use more sophisticated techniques or custom models for phone numbers and emails

    # For simplicity, we'll use basic string matching for emails and URLs
    if "@" in text:
        details["Email"] = [word for word in text.split() if "@" in word and "." in word]
    if "github.com" in text:
        details["GitHub Link"] = [word for word in text.split() if "github.com" in word]
    if "linkedin.com" in text:
        details["LinkedIn Link"] = [word for word in text.split() if "linkedin.com" in word]
    
    # Check for phone numbers with a placeholder approach
    phone_number_candidates = [word for word in text.split() if word.isdigit() and len(word) >= 10]
    if phone_number_candidates:
        details["Phone Number"] = phone_number_candidates[0]  # Simplified approach
    
    return details

# Process each PDF resume in the folder
for resume_filename in os.listdir(resume_folder):
    if resume_filename.endswith(".pdf"):
        pdf_path = os.path.join(resume_folder, resume_filename)
        
        # Extract text from PDF
        resume_text = extract_text_from_pdf(pdf_path)
        
        # Extract details from resume text
        details = extract_details(resume_text)
        details["Filename"] = resume_filename
        extracted_data.append(details)

# Convert extracted data to DataFrame and save to Excel
df = pd.DataFrame(extracted_data)
df.to_excel("extracted_resume_details_pdfplumber.xlsx", index=False)

print("Extraction complete. Data saved to 'extracted_resume_details.xlsx'")
