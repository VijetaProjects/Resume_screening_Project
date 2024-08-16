from transformers import pipeline
import PyPDF2
import os
import pandas as pd

# Load pre-trained classification model
classifier = pipeline("text-classification", model="distilbert-base-uncased-finetuned-sst-2-english")

# Function to extract text from PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        resume_text = ""
        for page in range(len(reader.pages)):
            resume_text += reader.pages[page].extract_text()
    return resume_text

# Function to classify text sections
def classify_text(text):
    result = classifier(text)
    return result[0]['label']

# Process resumes
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
extracted_data = []

for resume_filename in os.listdir(resume_folder):
    if resume_filename.endswith(".pdf"):
        pdf_path = os.path.join(resume_folder, resume_filename)
        resume_text = extract_text_from_pdf(pdf_path)
        # Example: Classify text and extract details based on labels
        classification = classify_text(resume_text)
        details = {
            "Filename": resume_filename,
            "Classification": classification
        }
        extracted_data.append(details)

# Save to Excel
df = pd.DataFrame(extracted_data)
df.to_excel("classified_resume_details.xlsx", index=False)

print("Classification complete. Data saved to 'classified_resume_details.xlsx'")
