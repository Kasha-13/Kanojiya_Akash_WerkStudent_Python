# PDF Invoice Data Extraction and Analysis

## Overview

This project processes invoice data from PDF files to extract relevant information, summarize it, and save the results in easy-to-use formats like Excel and CSV. The primary goal is to automate the tedious process of manual data extraction from invoices for quick analysis and record-keeping.

## Tasks Performed

### 1. Extracting Specific Information from PDFs
**Purpose**: Identify and extract key details like:
- Gross Amount (from the first invoice in EUR).
- Total USD (from the second invoice in USD).

**Method**: Using pattern-matching techniques to locate relevant text in PDFs.

### 2. Data Extraction and Excel File Creation
**Purpose**: Extract and store important details, such as invoice dates and amounts, into an Excel file for analysis.

**Details**:
- Data from both invoices is compiled into a single sheet.
- A pivot table summarizes totals by date for easier comparison.

**Output**: The Excel file (`output.xlsx`) includes:
- **Sheet 1**: Raw extracted data.
- **Sheet 2**: Summary pivot table.

### 3. Comprehensive Data Extraction for CSV
**Purpose**: Extract as much detail as possible from invoices, including:
- Company name and address.
- Customer name and address.
- Invoice number, date, and periods.
- Total and gross amounts.

**Method**: Advanced text recognition with defined patterns to locate details.

**Output**: A structured CSV file (`extracted_invoice_data.csv`) containing extracted data for each invoice.

## How It Works

### PDF Reading:
- PDFs are scanned for text using specialized libraries.
- Only the relevant details are extracted using predefined rules.

### Data Processing:
- Text is cleaned and transformed into a structured format.

### Output Generation:
- Two files are created:
  - `output.xlsx` for an organized overview.
  - `extracted_invoice_data.csv` for comprehensive record storage.

## Technical Details

### Tools Used:
- **Python Libraries**:
  - `PyPDF2` and `fitz` (for reading PDFs).
  - `re` (for text pattern matching).
  - `pandas` (for Excel creation).
  - `csv` (for generating CSV files).

### Approach:
- Scan the PDF content page by page.
- Look for specific patterns like dates, amounts, and names.
- Save extracted information into clear and structured files.

## Outputs

### Excel File: `output.xlsx`
- **Sheet 1**: Raw data extracted from invoices.
- **Sheet 2**: Pivot table summarizing totals by date.

### CSV File: `extracted_invoice_data.csv`
- Includes all extracted details, ready for further analysis.
