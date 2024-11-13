import pandas as pd
import glob
import os
from fpdf import FPDF
from pathlib import Path

DATA_SOURCE = "invoices"
LOGO_PATH = "pythonhow.png"
OUTPUT_PATH = "PDFs"
PRODUCT_ID_KEY = "product_id"
PRODUCT_NAME_KEY = "product_name"
AMOUNT_PURCHASED_KEY = "amount_purchased"
PRICE_PER_UNIT_KEY = "price_per_unit"
TOTAL_PRICE_KEY = "total_price"

if not Path(OUTPUT_PATH).exists():
    os.mkdir(OUTPUT_PATH)

invoice_paths = glob.glob(f"{DATA_SOURCE}/*.xlsx")

for path in invoice_paths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    file_name = Path(path).stem
    invoice_number, date = file_name.split("-")

    pdf.add_page()

    # Invoice number
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoide nr. {invoice_number}", ln=1)

    # Date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)

    # Table header
    data = pd.read_excel(path, sheet_name="Sheet 1")
    columns = [c.replace("_", " ").title() for c in data.columns]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Table rows
    for index, row in data.iterrows():

        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row[PRODUCT_ID_KEY]), border=1)
        pdf.cell(w=70, h=8, txt=str(row[PRODUCT_NAME_KEY]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[AMOUNT_PURCHASED_KEY]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[PRICE_PER_UNIT_KEY]), border=1)
        pdf.cell(w=30, h=8, txt=str(row[TOTAL_PRICE_KEY]), border=1, ln=1)

    # Total price
    total_price = data[TOTAL_PRICE_KEY].sum()

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)

    # Total price sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_price}", ln=1)
    
    # Company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"PythonHow")
    pdf.image(LOGO_PATH, w=10)

    pdf.output(f"{OUTPUT_PATH}/{file_name}.pdf")