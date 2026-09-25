from pathlib import Path


# Project root directory
BASE_DIR = Path(__file__).resolve().parent.parent


# Directories
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"


# Excel input file
INPUT_FILE = INPUT_DIR / "rent_data.xlsx"


# Required Excel columns
REQUIRED_COLUMNS = [
    "Flat No",
    "Name",
    "Society Name",
    "Member",
    "Contact",
    "Email",
    "Rent"
]


# PDF title
PDF_TITLE = "Rent Details"