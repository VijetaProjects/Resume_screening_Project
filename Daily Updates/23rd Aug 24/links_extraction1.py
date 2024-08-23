import fitz  # PyMuPDF
import os
import time
import pandas as pd

# Function to extract links and images with their coordinates using PyMuPDF (fitz)
def extract_links_and_images_fitz(pdf_path):
    links = []
    images = []
    pdf_document = fitz.open(pdf_path)
    
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        
        # Extract links with coordinates
        for link in page.get_links():
            if 'uri' in link:
                link_info = {
                    "uri": link['uri'],
                    "rect": fitz.Rect(link['from'])  # Get the coordinates of the link
                }
                links.append(link_info)
        
        # Extract images with coordinates using alternative method
        for img in page.get_images(full=True):
            try:
                xref = img[0]  # Reference number of the image
                image_rect = page.get_image_bbox(xref)  # Try to get the bounding box of the image
                if image_rect:
                    image_info = {
                        "xref": xref,
                        "rect": image_rect  # Store the bounding box (coordinates) of the image
                    }
                    images.append(image_info)
            except ValueError:
                # Skip images that do not have a bounding box
                continue
    
    return links, images

# Function to filter out links that overlap with images
def filter_links(links, images):
    filtered_links = []
    
    for link in links:
        link_rect = link['rect']
        # Check if the link's rectangle overlaps with any image rectangle
        overlapping = any(link_rect.intersects(image['rect']) for image in images)
        
        if not overlapping:
            filtered_links.append(link['uri'])
    
    return filtered_links

# Function to extract and filter links from PDFs and save the results to an Excel sheet
def compare_extraction_methods_and_save(resume_folder, output_excel_path):
    results = []

    for filename in os.listdir(resume_folder):
        if filename.endswith(".pdf"):
            file_path = os.path.join(resume_folder, filename)
            print(f"Processing: {filename}")

            # Extract links and images using PyMuPDF (fitz)
            start_fitz = time.time()
            links_fitz, images_fitz = extract_links_and_images_fitz(file_path)
            filtered_links_fitz = filter_links(links_fitz, images_fitz)
            end_fitz = time.time()

            # Collect results for this PDF
            results.append({
                "Filename": filename,
                "Filtered_Links": ", ".join(filtered_links_fitz) if filtered_links_fitz else "No links",
                "fitz_Time (s)": round(end_fitz - start_fitz, 4)
            })

    # Convert results to a DataFrame and save to Excel
    df = pd.DataFrame(results)
    df.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

# Specify the folder containing the PDFs and the output Excel file path
resume_folder = r"C:\Users\vijet\OneDrive\Desktop\Profiling Project\Resumes Data\50 Resumes"
output_excel_path = r"filtered_links_comparison.xlsx"

# Compare the extraction methods on the PDFs and save the results to an Excel file
compare_extraction_methods_and_save(resume_folder, output_excel_path)
