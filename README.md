# ETL-data-pipeline
# Insurance PDF to Excel Parser 🧾➡️📊

This project extracts structured content and tables from insurance PDF documents (specifically IndiaFirst Life's product forms) and converts them into clean, analyzable Excel sheets.

## Features

- Extracts text between specific sections (e.g., 6.2 to 6.3)
- Skips page headers and footers
- Detects subheaders like `a)`, `b)`, `i)` and groups content accordingly
- Parses inline tables and exports them as properly formatted Excel columns
- No hardcoding — works dynamically for structured PDFs

## Tech Stack

- Python
- `pdfplumber` for PDF parsing
- `pandas` for data structuring
- `openpyxl` for Excel writing
- `re` for regex-based subheader detection

## Usage

1. Install dependencies:
   ```bash
   pip install pdfplumber pandas openpyxl
