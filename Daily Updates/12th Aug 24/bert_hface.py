from transformers import pipeline
import PyPDF2
import pandas as pd
import os

# Load pre-trained NER model
ner_pipeline = pipeline("ner", model="dbmdz/bert-large-cased-finetuned-conll03-english", tokenizer="dbmdz/bert-large-cased-finetuned-conll03-english")

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        resume_text = ""
        for page in range(len(reader.pages)):
            resume_text += reader.pages[page].extract_text()
    return resume_text

# Function to extract details using transformer-based NER
def extract_details(text):
    entities = ner_pipeline(text)
    details = {
        "Name": None,
        "Location": None,
        "Phone Number": None,
        "Email": None,
        "GitHub Link": None,
        "LinkedIn Link": None
    }

    for entity in entities:
        if entity['entity'] == 'B-PER':
            details["Name"] = entity['word']
        elif entity['entity'] == 'B-LOC':
            details["Location"] = entity['word']
    
    # Simple string matching for emails and URLs
    if "@" in text:
        details["Email"] = [word for word in text.split() if "@" in word and "." in word]
    if "github.com" in text:
        details["GitHub Link"] = [word for word in text.split() if "github.com" in word]
    if "linkedin.com" in text:
        details["LinkedIn Link"] = [word for word in text.split() if "linkedin.com" in word]
    
    phone_number_candidates = [word for word in text.split() if word.isdigit() and len(word) >= 10]
    if phone_number_candidates:
        details["Phone Number"] = phone_number_candidates[0]
    
    return details

# Process resumes
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
extracted_data = []

for resume_filename in os.listdir(resume_folder):
    if resume_filename.endswith(".pdf"):
        pdf_path = os.path.join(resume_folder, resume_filename)
        resume_text = extract_text_from_pdf(pdf_path)
        details = extract_details(resume_text)
        details["Filename"] = resume_filename
        extracted_data.append(details)

# Save to Excel
df = pd.DataFrame(extracted_data)
df.to_excel("extracted_resume_details.xlsx", index=False)

print("Extraction complete. Data saved to 'extracted_resume_details.xlsx'")
