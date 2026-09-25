import logging
import re
from pathlib import Path

import pandas as pd

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle
)
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

from config.config import (
    INPUT_FILE,
    OUTPUT_DIR,
    LOG_DIR,
    REQUIRED_COLUMNS,
    PDF_TITLE
)


# ============================================================
# 1. CREATE DIRECTORIES
# ============================================================

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

LOG_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# 2. LOGGING CONFIGURATION
# ============================================================

LOG_FILE = LOG_DIR / "automation.log"

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logger = logging.getLogger(__name__)


# ============================================================
# 3. HELPER FUNCTIONS
# ============================================================

def clean_filename(value):
    """
    Removes invalid characters from a filename.
    """

    value = str(value)

    value = re.sub(
        r'[<>:"/\\|?*]',
        '',
        value
    )

    value = value.strip()

    value = value.replace(
        " ",
        "_"
    )

    return value


def format_rent(value):
    """
    Formats rent value with commas.
    """

    try:
        return f"₹ {float(value):,.0f}"

    except (ValueError, TypeError):

        return str(value)


# ============================================================
# 4. VALIDATE EXCEL
# ============================================================

def validate_excel(df):

    missing_columns = [
        column
        for column in REQUIRED_COLUMNS
        if column not in df.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Missing Excel columns: {missing_columns}"
        )


# ============================================================
# 5. CREATE PDF
# ============================================================

def create_pdf(row, output_path):

    document = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=20 * mm,
        bottomMargin=20 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontSize=20,
        leading=24,
        alignment=TA_CENTER,
        spaceAfter=20
    )

    heading_style = ParagraphStyle(
        "HeadingStyle",
        parent=styles["Heading2"],
        fontSize=12,
        leading=16,
        spaceAfter=10
    )

    normal_style = ParagraphStyle(
        "NormalStyle",
        parent=styles["Normal"],
        fontSize=10,
        leading=14
    )

    story = []

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            PDF_TITLE,
            title_style
        )
    )

    story.append(
        Paragraph(
            f"<b>Society:</b> {row['Society Name']}",
            heading_style
        )
    )

    story.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    data = [

        ["Field", "Details"],

        [
            "Flat No",
            str(row["Flat No"])
        ],

        [
            "Name",
            str(row["Name"])
        ],

        [
            "Members",
            str(row["Member"])
        ],

        [
            "Contact",
            str(row["Contact"])
        ],

        [
            "Email",
            str(row["Email"])
        ],

        [
            "Monthly Rent",
            format_rent(row["Rent"])
        ]

    ]

    # --------------------------------------------------------
    # TABLE
    # --------------------------------------------------------

    table = Table(
        data,
        colWidths=[
            50 * mm,
            110 * mm
        ]
    )

    table.setStyle(
        TableStyle([

            # Header
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            # First column
            (
                "FONTNAME",
                (0, 1),
                (0, -1),
                "Helvetica-Bold"
            ),

            # Font
            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                10
            ),

            # Grid
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),

            # Alignment
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            # Padding
            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                8
            )
        ])
    )

    story.append(table)

    story.append(
        Spacer(1, 20)
    )

    story.append(
        Paragraph(
            "This document was generated automatically from Excel data.",
            normal_style
        )
    )

    # --------------------------------------------------------
    # GENERATE
    # --------------------------------------------------------

    document.build(story)


# ============================================================
# 6. MAIN AUTOMATION
# ============================================================

def main():

    logger.info(
        "========== Automation Started =========="
    )

    print(
        "\nExcel to PDF Automation Started\n"
    )

    # --------------------------------------------------------
    # CHECK INPUT FILE
    # --------------------------------------------------------

    if not Path(INPUT_FILE).exists():

        raise FileNotFoundError(
            f"Excel file not found: {INPUT_FILE}"
        )

    print(
        f"Reading Excel: {INPUT_FILE}"
    )

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    df = pd.read_excel(
        INPUT_FILE
    )

    print(
        f"Total records found: {len(df)}"
    )

    logger.info(
        f"Excel records: {len(df)}"
    )

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------

    validate_excel(df)

    print(
        "Excel validation successful."
    )

    # --------------------------------------------------------
    # PROCESS EACH ROW
    # --------------------------------------------------------

    success_count = 0
    failed_count = 0

    for index, row in df.iterrows():

        try:

            flat_no = clean_filename(
                row["Flat No"]
            )

            name = clean_filename(
                row["Name"]
            )

            filename = (
                f"{flat_no}_{name}.pdf"
            )

            output_path = (
                OUTPUT_DIR / filename
            )

            # Create PDF
            create_pdf(
                row,
                output_path
            )

            success_count += 1

            logger.info(
                f"PDF generated: {filename}"
            )

            print(
                f"[SUCCESS] {filename}"
            )

        except Exception as error:

            failed_count += 1

            logger.error(
                f"Row {index + 2} failed: {error}"
            )

            print(
                f"[FAILED] Row {index + 2}: {error}"
            )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    print(
        "\n======================================"
    )

    print(
        "Automation Completed"
    )

    print(
        "======================================"
    )

    print(
        f"Total records : {len(df)}"
    )

    print(
        f"Successful    : {success_count}"
    )

    print(
        f"Failed        : {failed_count}"
    )

    print(
        f"Output folder : {OUTPUT_DIR}"
    )

    print(
        f"Log file      : {LOG_FILE}"
    )

    logger.info(
        f"Completed | Total={len(df)} "
        f"Success={success_count} "
        f"Failed={failed_count}"
    )


# ============================================================
# 7. PROGRAM ENTRY POINT
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logger.exception(
            "Automation failed"
        )

        print(
            f"\nERROR: {error}"
        )