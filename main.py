import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    inv_num = filename.split("-")[0]
    date = filename.split("-")[1]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0,h=20, txt=f"Invoice Number: {inv_num}")
    pdf.cell(w=0, h=20, txt=f"Date: {date}")
    pdf.output(f"PDFs/{inv_num}.pdf")
