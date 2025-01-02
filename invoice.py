import re
import pandas as pd
import PyPDF2
from datetime import datetime
import fitz  # PyMuPDF for reading PDFs
import csv

# Task 1: Pattern Matching from PDFs

# Paths to PDF files
pdf1_path = "sample_invoice_1.pdf"
pdf2_path = "sample_invoice_2.pdf"

# Define patterns to extract specific text from PDFs
pattern_pdf1 = r"Gross Amount incl\. VAT\s+[\d,]+\s\u20AC"  # Pattern for PDF 1
pattern_pdf2 = r"Total\s*USD\s+\$[\d,]+\.\d{2}"             # Pattern for PDF 2

# Extract text from the first page of PDF 1
with fitz.open(pdf1_path) as pdf1:
    page = pdf1[0]  # Access the first page (0-based index)
    text_pdf1 = page.get_text()
    match_pdf1 = re.search(pattern_pdf1, text_pdf1)
    if match_pdf1:
        print("PDF 1 Match:", match_pdf1.group())
    else:
        print("No match found in PDF 1.")

# Extract text from the first page of PDF 2
with fitz.open(pdf2_path) as pdf2:
    page1 = pdf2[0]  # Access the first page (0-based index)
    text_pdf2 = page1.get_text()
    match_pdf2 = re.search(pattern_pdf2, text_pdf2)
    if match_pdf2:
        print("PDF 2 Match:", match_pdf2.group())
    else:
        print("No match found in PDF 2.")


# Task 2: Extract Data and Write to Excel

# Function to extract data from PDF 1
def extract_data_from_pdf1(text):
    """
    Extracts date and gross amount from the text of PDF 1.
    
    Args:
        text (str): Extracted text from the PDF.

    Returns:
        tuple: File name, extracted date, and gross amount.
    """
    file_name = "sample_invoice_1.pdf"
    date_match = re.search(r"Date\s*(\d{1,2}\.\s*\w+\s*\d{4})", text)
    date = date_match.group(1) if date_match else None
    gross_amount_match = re.search(r"Gross Amount incl\. VAT\s+([\d,]+)\s\u20AC", text)
    gross_amount = float(gross_amount_match.group(1).replace(',', '')) if gross_amount_match else None
    return file_name, date, gross_amount

# Function to extract data from PDF 2
def extract_data_from_pdf2(text):
    """
    Extracts invoice date and total USD from the text of PDF 2.
    
    Args:
        text (str): Extracted text from the PDF.

    Returns:
        tuple: File name, extracted date, and total USD.
    """
    file_name = "sample_invoice_2.pdf"
    date_match = re.search(r"Invoice date:\s*(\w+\s\d{1,2},\s\d{4})", text)
    total_usd_match = re.search(r"Total USD \$([\d,.]+)", text)
    date = datetime.strptime(date_match.group(1), "%b %d, %Y").strftime("%Y-%m-%d") if date_match else None
    total_usd = float(total_usd_match.group(1).replace(',', '')) if total_usd_match else None
    return file_name, date, total_usd

# Read the PDFs and extract data
pdf_files = ['sample_invoice_1.pdf', 'sample_invoice_2.pdf']
data = []

for pdf_file in pdf_files:
    with open(pdf_file, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        text = "\n".join(page.extract_text() for page in reader.pages)

        if pdf_file == "sample_invoice_1.pdf":
            data.append(extract_data_from_pdf1(text))
        elif pdf_file == "sample_invoice_2.pdf":
            data.append(extract_data_from_pdf2(text))

# Create an Excel file with two sheets
sheet1_data = pd.DataFrame(data, columns=["File Name", "Date", "Value"])
pivot_table = sheet1_data.pivot_table(index="Date", values="Value", columns="File Name", aggfunc="sum", fill_value=0)

with pd.ExcelWriter("output.xlsx") as writer:
    sheet1_data.to_excel(writer, index=False, sheet_name="Sheet1")
    pivot_table.to_excel(writer, sheet_name="Sheet2")

print("Excel file 'output.xlsx' created successfully!")

# Task 3: Extract Comprehensive Data from PDFs


def extract_text_from_pdf(pdf_path):
    """
    Extracts all text from a PDF file.

    Args:
        pdf_path (str): Path to the PDF file.

    Returns:
        str: The concatenated text from all pages in the PDF.
    """
    text = ""
    try:
        with fitz.open(pdf_path) as pdf:
            for page in pdf:
                text += page.get_text()
    except Exception as e:
        print(f"Error reading PDF: {e}")
    return text

def extract_info_from_text(text):
    """
    Extracts specific information from the provided text using regex patterns.

    Args:
        text (str): Text extracted from the PDF.

    Returns:
        dict: A dictionary containing extracted data fields.
    """
    data = {}

    # Extracting company name
    company_pattern = re.compile(r"\b[A-Z][\w\s,&.\-()]*\s(?:Ltd|GmbH|Pty\sLtd|Inc)\b", re.IGNORECASE)
    company_match = company_pattern.search(text)
    data['company_name'] = company_match.group(0).strip() if company_match else None

    # Extracting company address
    address_pattern = re.compile(r"(\d+\s+[\w\s,.]+(?:\s+[\w\s,.]+)*)", re.MULTILINE)
    address_matches = address_pattern.findall(text)
    data['company_address'] = address_matches[0].strip() if address_matches else None

    # Extracting customer name
    name_pattern = re.compile(r"\b(Mr\.?|Mrs\.?|Ms\.?|Dr\.?)\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*", re.IGNORECASE)
    name_match = name_pattern.search(text)
    data['customer_name'] = name_match.group(0).strip() if name_match else None

    # Extracting phone number
    phone_pattern = re.compile(r"\+?\d{1,3}[\s-]?\(?\d+\)?[\s-]?\d+[\s-]?\d+")
    phone_match = phone_pattern.search(text)
    data['phone_number'] = phone_match.group(0).strip() if phone_match else None

    # Extracting customer address
    customer_address_pattern = re.compile(r"(?:Street|St|Strasse|Str)\s*[:\s]*(.+)", re.IGNORECASE)
    customer_address_match = customer_address_pattern.search(text)
    data['customer_address'] = customer_address_match.group(1).strip() if customer_address_match else None

    # Extracting invoice number
    invoice_no_pattern = re.compile(r"Invoice\s*(Number|No)[:\s]*([\w-]+)", re.IGNORECASE)
    invoice_no_match = invoice_no_pattern.search(text)
    data['invoice_no'] = invoice_no_match.group(2).strip() if invoice_no_match else None

    # Extracting invoice period
    period_pattern = re.compile(r"Period:\s*(\d{2}\.\d{2}\.\d{4})\s*to\s*(\d{2}\.\d{2}\.\d{4})", re.IGNORECASE)
    from_until_pattern = re.compile(r"From\s*.*?Until\s*.*?\n.*?\n.*?\n(\w+\s\d{1,2},\s\d{4})\n(\w+\s\d{1,2},\s\d{4})", re.IGNORECASE)

    from_date, until_date = None, None
    period_match = period_pattern.search(text)
    if period_match:
        from_date, until_date = period_match.groups()
    else:
        from_until_match = from_until_pattern.search(text)
        if from_until_match:
            from_date, until_date = from_until_match.groups()

    data['from_date'] = from_date
    data['until_date'] = until_date

    # Extracting invoice date
    invoice_date_pattern = re.compile(r'(?:Invoice\s*Date|Date)[:\s]*(\d{1,2}\.\s\w+\s\d{4}|[A-Za-z]{3}\s\d{1,2},\s\d{4})', re.IGNORECASE)
    invoice_date_match = invoice_date_pattern.search(text)
    data['invoice_date'] = invoice_date_match.group(1).strip() if invoice_date_match else None

    # Extracting total amount
    total_amount_pattern = re.compile(r"Total\s*(?:USD|EUR|INR)?[\s$]*([\d,.]+)", re.IGNORECASE)
    total_match = total_amount_pattern.search(text)
    data['total_amount'] = total_match.group(1).strip() if total_match else None

    # Extracting gross amount
    gross_amount_pattern = re.compile(r"Gross\s*Amount(?: incl\. VAT)?[:\s]+([\d,.]+)", re.IGNORECASE)
    gross_match = gross_amount_pattern.search(text)
    data['gross_amount'] = gross_match.group(1).strip() if gross_match else None

    return data

def clean_data_for_csv(data):
    """
    Cleans extracted data for CSV storage by replacing newlines with spaces.

    Args:
        data (dict): Data dictionary.

    Returns:
        dict: Cleaned data dictionary.
    """
    for key, value in data.items():
        if isinstance(value, str):
            data[key] = value.replace('\n', ' ')
    return data

def save_data_to_csv(data_list, output_csv):
    """
    Saves extracted data to a CSV file.

    Args:
        data_list (list): List of data dictionaries.
        output_csv (str): Path to the output CSV file.
    """
    keys = data_list[0].keys() if data_list else []
    try:
        with open(output_csv, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=keys, delimiter=';')
            writer.writeheader()
            cleaned_data_list = [clean_data_for_csv(data) for data in data_list]
            writer.writerows(cleaned_data_list)
    except Exception as e:
        print(f"Error writing to CSV: {e}")

if __name__ == "__main__":
    # Define PDF files to process
    pdf_files = ["sample_invoice_1.pdf", "sample_invoice_2.pdf"]
    extracted_data_list = []

    # Extract data from each PDF
    for pdf_file in pdf_files:
        print(f"Extracting data from: {pdf_file}")
        text = extract_text_from_pdf(pdf_file)
        extracted_data = extract_info_from_text(text)
        extracted_data_list.append(extracted_data)

    # Save the data to a CSV file
    output_csv_path = "extracted_invoice_data.csv"
    save_data_to_csv(extracted_data_list, output_csv_path)
    print(f"Data saved to {output_csv_path}")
