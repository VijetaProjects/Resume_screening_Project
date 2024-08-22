import os
import pdfplumber
import pandas as pd

def extract_hyperlinks(pdf_path):
    github_links = set()  # Use a set to ensure uniqueness
    linkedin_links = set()  # Ensure LinkedIn links are unique
    
    # Helper function to find hyperlinks in text
    def find_links(text, base_url):
        links = set()
        start = 0
        while start < len(text):
            start = text.find(base_url, start)
            if start == -1:
                break
            # Find the end of the URL (space, newline, etc.)
            end = start
            while end < len(text) and text[end] not in (' ', '\n', '\r', ')', '(', '<', '>'):
                end += 1
            # Extract the URL
            url = text[start:end]
            links.add(url)
            start = end
        return links

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract annotations (hyperlinks embedded in the PDF)
            if page.annots:
                for annot in page.annots:
                    if 'uri' in annot:
                        url = annot['uri']
                        if url:
                            if 'github.com' in url:
                                github_links.add(url)  # Add to set to remove duplicates
                            elif 'linkedin.com' in url:
                                linkedin_links.add(url)
            
            # Extract text from the page and search for URLs manually
            page_text = page.extract_text()
            if page_text:
                # Manually find GitHub and LinkedIn links in the text
                github_links.update(find_links(page_text, 'github.com/'))
                linkedin_links.update(find_links(page_text, 'linkedin.com/in/'))

    # Convert sets to lists and handle missing links
    github_link = list(github_links)[0] if github_links else "Not Mentioned"
    linkedin_link = list(linkedin_links)[0] if linkedin_links else "Not Mentioned"

    return github_link, linkedin_link

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
    df.to_excel("extracted_links1.xlsx", index=False)

    print("Extraction complete. Data saved to extracted_links.xlsx")

# Example usage
directory = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
process_pdfs_in_directory(directory)
