import os
import zipfile
import fitz  # PyMuPDF
import shutil

def extract_zip(zip_path, extract_to):
    """
    Extracts a ZIP file to the specified directory, preserving the directory structure.
    """
    print(f"Extracting {zip_path} to {extract_to}...")
    if os.path.exists(extract_to):
        shutil.rmtree(extract_to)
    os.makedirs(extract_to)

    with zipfile.ZipFile(zip_path, 'r') as zf:
        for old_name in zf.namelist():
            new_name = old_name.encode('cp437').decode('gbk')
            new_path = os.path.join(extract_to, new_name)
            file_size = zf.getinfo(old_name).file_size
            
            if file_size > 0:
                os.makedirs(os.path.dirname(new_path), exist_ok=True)
                with open(new_path, 'wb') as f:
                    f.write(zf.read(old_name))
                print(f"Extracted file: {new_path}")
            else:
                os.makedirs(new_path, exist_ok=True)
                print(f"Created directory: {new_path}")

    print(f"Extraction completed.")

def convert_pdf_to_images(pdf_path, output_folder):
    """
    Converts a PDF file to images using PyMuPDF and saves them to the specified folder.
    """
    print(f"Converting {pdf_path} to images...")
    try:
        # Open the PDF document
        pdf_document = fitz.open(pdf_path)
        
        for page_number in range(len(pdf_document)):
            # Load the page
            page = pdf_document.load_page(page_number)
            
            # Render the page to a pixmap (image)
            pix = page.get_pixmap()
            
            # Define the output image file path
            image_filename = os.path.join(output_folder, f"page_{page_number + 1}.png")
            
            # Save the pixmap as an image
            pix.save(image_filename)
            print(f"Saved {image_filename}")
        
        # Close the PDF document
        pdf_document.close()
    except Exception as e:
        # Log error to a text file
        with open('/data/Finance/failed_pdfs.txt', 'a') as f:
            f.write(f"Failed to convert {pdf_path}. Error: {e}\n")
        print(f"Failed to convert {pdf_path} to images. Error: {e}")

def read_processed_files(processed_file_list_path):
    """
    Reads the list of already processed files.
    """
    if os.path.exists(processed_file_list_path):
        with open(processed_file_list_path, 'r') as f:
            return set(line.strip() for line in f)
    return set()

def write_processed_file(processed_file_list_path, file_name):
    """
    Appends the processed file name to the processed files list.
    """
    with open(processed_file_list_path, 'a') as f:
        f.write(f"{file_name}\n")

def get_pdf_files(root_dir):
    """
    Returns a list of all PDF files found in the specified directory.
    """
    pdf_files = []
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

def is_folder_processed(output_folder):
    """
    Checks if the output folder contains PNG files, indicating that the folder has been processed.
    """
    if not os.path.exists(output_folder):
        return False
    for root, dirs, files in os.walk(output_folder):
        if any(file.lower().endswith('.png') for file in files):
            return True
    return False

def batch_process_files(pdf_files, output_folder, processed_file_list_path, batch_size):
    """
    Processes PDF files in batches.
    """
    processed_files = read_processed_files(processed_file_list_path)

    for i in range(0, len(pdf_files), batch_size):
        batch = pdf_files[i:i + batch_size]
        print(f"Processing batch {i // batch_size + 1}...")
        
        for pdf_path in batch:
            pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.join(output_folder, pdf_base_name)
            
            if is_folder_processed(output_dir):
                print(f"Skipping already processed PDF: {pdf_path}")
                continue
            
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            print(f"Processing PDF: {pdf_path}")
            try:
                convert_pdf_to_images(pdf_path, output_dir)
                write_processed_file(processed_file_list_path, os.path.basename(pdf_path))
            except Exception as e:
                print(f"Error processing file {pdf_path}: {e}")

def process_zip_files(zip_path, extract_to, output_folder, processed_file_list_path, batch_size):
    """
    Extracts ZIP files and processes the PDFs contained within.
    """
    extract_zip(zip_path, extract_to)
    pdf_files = get_pdf_files(extract_to)
    batch_process_files(pdf_files, output_folder, processed_file_list_path, batch_size)

def main():
    zip_path = '/data/金融研报/研报.zip'
    extract_to = '/data/extracted/研报/'
    output_folder = '/data/Finance/研报/'
    processed_file_list_path = '/data/Finance/研报_processed_files.txt'
    batch_size = 10  # Number of PDFs to process in each batch

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    process_zip_files(zip_path, extract_to, output_folder, processed_file_list_path, batch_size)

if __name__ == "__main__":
    main()
