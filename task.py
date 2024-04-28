import os
import re
import glob
import docx
import PyPDF2
import textract
import pandas as pd

# Glob all the docx and pdf files
resume = glob.glob("Sample2/*.docx")
resume2 = glob.glob("Sample2/*.pdf")

# Function to open docx documents
def open_docx_documents(file_path):
    doc = open(file_path, "rb")
    return docx.Document(doc)

# Function to extract text from docx files
def doc_to_text(file_path):
    text = textract.process(file_path).decode('utf-8')
    return text

# Function to open pdf documents
def open_pdf_documents(file_path):
    text = ''
    with open(file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to search for email addresses in a string
def search_email(string):
    return re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", string)

# Function to search for mobile numbers in a string
def search_mobile(string):
    return re.findall(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]', string)

# Function to open a pdf file
def open_pdf(file):
    try:
        pdfFileObj = open(file, "rb")
        return PyPDF2.PdfReader(pdfFileObj)
    except PyPDF2.errors.PdfReadError:
        print(f"Could not open {file}")
        return None

# Function to read text from a pdf file
def read_pdf(pdfReader):
    num_pages = pdfReader.numPages
    content = ""
    for i in range(0, num_pages):
        content += pdfReader.getPage(i).extract_text() + "\n"
    content = " ".join(content.replace(u"\xa0", " ").strip().split())
    return content

# Function to read email addresses from a pdf file
def read_email(pdfReader):
    num_pages = len(pdfReader.pages)
    for i in range(0, num_pages):
        return search_email(pdfReader.pages[i].extract_text())

# Function to read mobile numbers from a pdf file
def read_mobile(pdfReader):
    num_pages = len(pdfReader.pages)
    for i in range(0, num_pages):
        return search_mobile(pdfReader.pages[i].extract_text())

# Lists to store extracted emails and mobile numbers
emails = []
mobile = []

# Dictionary to store file information
file_dict = []

# Iterate over docx and pdf files
for file_list in [resume, resume2]:
    for file in file_list:
        # Open the file and extract information
        if file.endswith('.pdf'):
            file_info = open_pdf_documents(file)
        else:
            file_info = doc_to_text(file)
        
        # Search for emails and mobile numbers in the extracted text
        emails = search_email(file_info)
        mobile = search_mobile(file_info)

        # Extract the file name
        file_name = os.path.basename(file)

        # Append file information to the list
        file_dict.append(
            {'File_Name': file_name, 'Emails': emails[0] if emails else "Not Found",
             'Mobile': mobile[0] if mobile else "Not Found", 'Text': file_info})

# Convert list to DataFrame and save to Excel
file_info_dict = pd.DataFrame(file_dict)
file_info_dict.to_excel('output.xlsx', index=False)
print("Successfully extracted data and saved to output.xlsx")
