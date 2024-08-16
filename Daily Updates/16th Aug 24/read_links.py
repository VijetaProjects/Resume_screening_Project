import PyPDF2

# Open the PDF file
file_path = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Experimentation Documents\Vijeta _V _Wasnik Resume.pdf"
file = open(file_path, "rb")

# Read the PDF file
reader = PyPDF2.PdfReader(file)

# Function to extract hyperlinks
def extract_hyperlinks(reader):
    hyperlinks = []

    # Iterate over each page in the PDF
    for page_no in range(len(reader.pages)):
        page = reader.pages[page_no]
        if '/Annots' in page:
            for annot in page['/Annots']:
                annot_obj = annot.get_object()
                if '/A' in annot_obj:
                    link = annot_obj['/A']
                    if '/URI' in link:
                        hyperlinks.append(link['/URI'])
                    elif '/S' in link and link['/S'] == '/GoToR':
                        # This is for links that go to a destination
                        if '/F' in link and '/URI' in link['/F']:
                            hyperlinks.append(link['/F']['/URI'])

    return hyperlinks

# Extract hyperlinks from the PDF
urls = extract_hyperlinks(reader)

# Print the extracted hyperlinks
for url in urls:
    print(url)

# Close the file
file.close()
