import os
from PyPDF2 import PdfReader
import re
import pandas as pd
from openpyxl import load_workbook

def convert_pdf_to_txt(file_path):
    pdf_reader = PdfReader(file_path)
    texts = []
    for page in pdf_reader.pages:
        texts.append(page.extract_text())
    return texts

def write_to_txt_and_extract_invoice_number_and_email(file_path, text, page_number):
    invoice_number = None
    email = None
    for line in text.split('\n'):
        if "Faktura č.:" in line:
            match = re.search(r'Faktura č.:\s*(\d+)', line)
            if match:
                invoice_number = match.group(1)
        if "E-mail:" in line:
            match = re.search(r'(?i)E-mail:\s*([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)', line)
            if match:
                email = match.group(1)

    if invoice_number:
        # Create 'txt' directory inside 'faktury' if it doesn't exist
        if not os.path.exists('faktury/txt'):
            os.makedirs('faktury/txt')

        # Save .txt file in 'faktury/txt' directory
        with open(f"faktury/txt/{invoice_number}.txt", 'w', encoding='utf-8') as txt_file_obj:
            txt_file_obj.write(text)

    return invoice_number, email


def write_to_txt(file_path, data):
    with open(file_path, 'w', encoding='utf-8') as txt_file_obj:
        for item in data:
            txt_file_obj.write(f"{item}\n")

def write_to_excel(file_path, emails, invoice_numbers):
    invoice_numbers = [str(invoice) + '.pdf' for invoice in invoice_numbers]
    df = pd.DataFrame({
        'Email': emails,
        'Attachment': invoice_numbers
    })
    
    # Create 'mail_tool' directory if it doesn't exist
    if not os.path.exists('mail_tool'):
        os.makedirs('mail_tool')

    df.to_excel(file_path, index=False)
    book = load_workbook(file_path)
    sheet = book.active
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 11
    book.save(file_path)

# Get all PDF files in the 'faktury' directory
pdf_files = [f for f in os.listdir('faktury') if f.endswith('.pdf')]

invoice_numbers = []
emails = []

# Process each PDF file
for pdf_file in pdf_files:
    pdf_path = os.path.join('faktury', pdf_file)
    texts = convert_pdf_to_txt(pdf_path)
    for i, text in enumerate(texts):
        invoice_number, email = write_to_txt_and_extract_invoice_number_and_email(pdf_path, text, i+1)
        if invoice_number:
            invoice_numbers.append(invoice_number)
        if email:
            emails.append(email)

invoice_txt_path = 'invoice_numbers.txt'
email_txt_path = 'emails.txt'
write_to_txt(invoice_txt_path, invoice_numbers)
write_to_txt(email_txt_path, emails)

excel_path = 'mail_tool/recipients.xlsx'  # Save Excel file in 'mail_tool' directory
write_to_excel(excel_path, emails, invoice_numbers)

