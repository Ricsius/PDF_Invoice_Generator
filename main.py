import pandas as pd
import glob
import os
from fpdf import FPDF
from pathlib import Path

DATA_SOURCE = "invoices"
OUTPUT_PATH = "PDFs"

if not Path(OUTPUT_PATH).exists():
    os.mkdir(OUTPUT_PATH)

invoice_paths = glob.glob(f"{DATA_SOURCE}/*.xlsx")

for path in invoice_paths:
    data = pd.read_excel(path, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    file_name = Path(path).stem
    invoice_number = file_name.split("-")[0]

    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoide nr. {invoice_number}")
    pdf.output(f"{OUTPUT_PATH}/{file_name}.pdf")