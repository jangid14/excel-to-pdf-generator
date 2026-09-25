# Excel to PDF Generator

A Python automation tool that reads records from an Excel file
and generates individual PDF documents using ReportLab.

## Features

- Read Excel data using Pandas
- Validate required columns
- Generate individual PDFs
- Format rent values
- Sanitize filenames
- Error handling
- Logging
- Success/failure summary

## Tech Stack

- Python
- Pandas
- OpenPyXL
- ReportLab

## Project Structure

excel_to_pdf_generator/
├── main.py
├── requirements.txt
├── config/
├── input/
├── output/
└── logs/

## Installation

Create virtual environment:

python -m venv venv

Activate:

venv\Scripts\activate

Install dependencies:

pip install -r requirements.txt

## Run

python main.py

## Input

The Excel file should contain:

- Flat No
- Name
- Society Name
- Member
- Contact
- Email
- Rent

## Output

The application generates one PDF per Excel record.

Example:

A101_Rahul_Sharma.pdf

## Architecture

Excel
 ↓
Pandas
 ↓
Validation
 ↓
Row Processing
 ↓
ReportLab
 ↓
PDF

## Future Improvements

- FastAPI API
- Kafka
- PostgreSQL
- Redis
- Docker
- AWS S3
- Microservices architecture
