from __future__ import annotations

import os
from io import BytesIO
from typing import List, Any

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
import barcode
from barcode.writer import ImageWriter



def generate_code128_barcode(value: str, width: int = 80, height: int = 25) -> Any:
    """
    Generate a Code128 barcode as a ReportLab Image object.

    Parameters
    ----------
    value : str
        The label value to encode.
    width : int
        Width of the generated barcode image.
    height : int
        Height of the generated barcode image.

    Returns
    -------
    Image | str
        Returns a ReportLab Image object if successful, otherwise an empty string.
    """
    cleaned = str(value).strip().replace(" ", "")

    if not cleaned or cleaned.lower() == "nan":
        return ""

    try:
        buffer = BytesIO()
        code = barcode.get("code128", cleaned, writer=ImageWriter())
        code.write(buffer, options={"module_height": 5.0, "quiet_zone": 1})
        buffer.seek(0)

        return Image(buffer, width=width, height=height)

    except Exception as exc:
        print(f"⚠️ Warning: Failed to generate barcode for '{value}' → {exc}")
        return ""


def excel_to_barcode_pdf(input_excel: str, output_pdf: str = "output_labels.pdf") -> None:
    """
    Convert an Excel spreadsheet to a barcode PDF table.

    Parameters
    ----------
    input_excel : str
        Path to the input Excel file.
    output_pdf : str
        Path to the output PDF file.

    Returns
    -------
    None
    """

    try:
        df = pd.read_excel(input_excel)
    except Exception as exc:
        raise FileNotFoundError(f"Unable to read Excel file '{input_excel}': {exc}")

    df = df.fillna("").astype(str)

    # Remove Label Number 5 if present
    if "Label Number 5" in df.columns:
        df = df.drop(columns=["Label Number 5"])

    expected_cols = [
        "S.No.", "Applicant No", "Artisan Name",
        "Label Number 1", "Label Number 2",
        "Label Number 3", "Label Number 4",
    ]

    available_cols = [col for col in expected_cols if col in df.columns]

    if not available_cols:
        raise ValueError("❌ No expected columns found in Excel file.")

    df = df[available_cols]

    # Prepare table header
    barcode_headers = [f"Label {i} Barcode" for i in range(1, 5)]
    header_row = available_cols + barcode_headers

    table_data = [header_row]

    # Create table rows
    for _, row in df.iterrows():
        row_cells: List[Any] = [row.get(col, "") for col in available_cols]

        # Add generated barcode objects
        for i in range(1, 5):
            col = f"Label Number {i}"
            barcode_img = generate_code128_barcode(row.get(col, ""))
            row_cells.append(barcode_img)

        table_data.append(row_cells)

    # Create PDF
    pdf = SimpleDocTemplate(output_pdf, pagesize=landscape(A3))
    table = Table(table_data, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 6),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
    ]))

    pdf.build([table])

    print(f"✅ PDF successfully created: {output_pdf}")


if __name__ == "__main__":

    input_excel = "path/to/input_labels.xlsx"
    output_pdf = "output_labels.pdf"

    if not os.path.exists(input_excel):
        print(f"❌ Excel file not found: {input_excel}")
    else:
        excel_to_barcode_pdf(input_excel, output_pdf)
