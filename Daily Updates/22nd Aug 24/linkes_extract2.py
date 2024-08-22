import os
import pdfplumber
import pandas as pd


# Function to extract all GitHub and LinkedIn hyperlinks from the PDF using pdfplumber
def extract_hyperlinks(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = set()  # Ensure LinkedIn links are unique
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        url = annot['uri']
                        if url:
                            if 'github.com' in url:
                                github_links.add(url)  # Add to set to remove duplicates
                            elif 'linkedin.com' in url:
                                linkedin_links.add(url)
    
    return list(github_links), list(linkedin_links)


def process_pdfs_in_directory(directory):
    # List to store all the data
    data = []

    # Iterate over all files in the directory
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(directory, filename)
            github_link, linkedin_link = extract_hyperlinks(pdf_path)
            
            # Add data to the list
            data.append({"Filename": filename, "Github Link": github_link, "Linkedin Link": linkedin_link})

    # Convert the data to a pandas DataFrame
    df = pd.DataFrame(data, columns=["Filename", "Github Link", "Linkedin Link"])

    # Save the DataFrame to an Excel file
    df.to_excel("extracted_links2.xlsx", index=False)

    print("Extraction complete. Data saved to extracted_links.xlsx")

# Example usage
directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
process_pdfs_in_directory(directory)
