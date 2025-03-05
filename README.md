# PDF to Images Conversion Tool

## Introduction
This Python tool extracts PDF files from ZIP archives and converts them to images. It supports batch processing, maintains a list of processed files to avoid duplicate conversions, and provides error logging functionality.

## Features
- Extracts PDF files from ZIP archives while preserving directory structure
- Converts PDF pages to images using PyMuPDF (fitz)
- Supports configurable batch processing
- Maintains a list of processed files to avoid reprocessing
- Provides error logging functionality

## Requirements
- Python 3.x
- PyMuPDF (fitz) library
- zipfile and shutil libraries (included in Python's standard library)

## Installation
1. Clone the repository to your local machine:
```bash
git clone https://github.com/your-username/pdf_to_images.git
cd pdf_to_images
```
## Install the required dependencies:

pip install fitz
Usage
Configure the script:
Modify the zip_path, extract_to, output_folder, and processed_file_list_path variables in the main() function to match your directory structure.
Set the batch_size variable to control the number of PDF files processed in each batch.
## Run the script:
```bash
python main.py
```
## Output
- **Converted Images**: Extracted PDF files will be converted to images and saved in the specified `output_folder`
- **Processed Files Log**: Processed files will be logged in the `processed_file_list_path` file
- **Error Log**: Errors during conversion will be logged to `/data/Finance/failed_pdfs.txt`

## Directory Structure
The script expects the following directory structure:

`
data
├── 金融研报
│   └── 研报.zip
├── extracted
│   └── 研报
└── Finance
   ├── 研报
   └── 研报\_processed_files.txt
`

## Notes
- **Directory Verification**: Ensure that all specified directories exist or are created before running the script
- **ZIP Structure Assumption**: The script assumes the ZIP file contains PDF files. Adjust the script if the ZIP structure differs
- **Error Log Maintenance**: The error log file (`failed_pdfs.txt`) will append new errors with each script execution

## Contributing
We welcome contributions! If you find any bugs or have suggestions for improvements, please:
- Submit an issue to report bugs
- Create a pull request to propose improvements
