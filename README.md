# OCR Project: Time Sheet Recognition from Low-Quality Scans

## Overview

This project leverages Optical Character Recognition (OCR) techniques to extract data from poorly scanned or photographed time sheets into an excel file structure. The primary objective is to automate the extraction of critical information such as employee names, hours worked, dates, and job codes, even when the source documents are of low quality.

## Features

- **Time Sheet Parsing**: Extracts structured data from time sheet images.
- **PDF to Excel Conversion**: Converts scanned PDF time sheets into editable Excel files.
- **Image Preprocessing**: Applies techniques like thresholding and noise reduction to enhance OCR accuracy.
- **GUI Integration**: Provides a graphical user interface for ease of use.

## Installation

### Setup

1. Clone the repository:

   ```bash
   git clone https://github.com/sdening/OCR_project.git
   cd OCR_project
Install required Python packages:
pip install -r requirements.txt
Install Tesseract-OCR:
Windows: Download and install from Tesseract at UB Mannheim.
macOS: Install via Homebrew:
brew install tesseract
Linux: Install using apt:
sudo apt-get install tesseract-ocr
Configure Tesseract path (if necessary):
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows path
Usage

Run the main script to process a time sheet image:

python ocr_an_img.py -i path_to_image.jpg
This command will process the specified image and output the extracted data to a text file.

For PDF files, use the PDF to Excel conversion script:

python pdf_to_excel.py -i path_to_pdf.pdf
This will convert the scanned PDF into an Excel file with extracted data.

File Descriptions

ocr_an_img.py: Processes individual image files to extract text.
pdf_to_excel.py: Converts scanned PDF time sheets into Excel format.
Test_in_Env_GUI.ipynb: Jupyter Notebook for testing the OCR functionality in a GUI environment.
alte_Excel/: Directory containing sample Excel files for testing.
alte_PDF/: Directory containing sample PDF files for testing.
