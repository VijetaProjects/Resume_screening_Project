import pandas as pd
import re

# Function to clean text
def clean_text_column(text):
    if isinstance(text, str):  # Check if the input is a string
        # Adding space after certain keywords for better formatting
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        text = re.sub(r'(\d)([A-Z])', r'\1 \2', text)
        text = re.sub(r'([A-Z])([A-Z][a-z])', r'\1 \2', text)
        text = re.sub(r'(\d)([a-zA-Z])', r'\1 \2', text)
    return text

# Function to clean the 'Extracted Text' column in a DataFrame
def clean_extracted_text_column(df, column_name='Extracted Text'):
    # Apply the clean_text function to the specified column
    df[column_name] = df[column_name].apply(clean_text_column)
    return df

# Example usage:
# Load the Excel file with proper escaping for backslashes
df = pd.read_excel(r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\16 Aug\resumes_text_pdfplumber.xlsx")

# Clean the 'Extracted Text' column
cleaned_df = clean_extracted_text_column(df, 'Extracted Text')

# Save the cleaned data to a new Excel file
cleaned_df.to_excel(r'C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Daily Work\17 Aug\cleaned_resume_info.xlsx', index=False)
