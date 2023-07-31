import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    inv_num = filename.split("-")[0]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Number: {inv_num}", ln=1)

    date = filename.split("-")[1]
    date = date.replace(".", "/")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # header block
    headers = list(df.columns)
    headers = [i.replace("_"," ").title() for i in headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(10, 10, 10)
    pdf.cell(w=20, h=8, txt=headers[0], border=1)
    pdf.cell(w=50, h=8, txt=headers[1], border=1)
    pdf.cell(w=35, h=8, txt=headers[2], border=1)
    pdf.cell(w=35, h=8, txt=headers[3], border=1)
    pdf.cell(w=25, h=8, txt=headers[4], border=1, ln=1)

    # rows under headers
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=20, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=25, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # last row with sum of amounts
    total = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=20, h=8, txt="", border=1)
    pdf.cell(w=50, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt=str(total), border=1, ln=1)

    # Footer info
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=20, h=8, txt=f"The total amount to be paid: {total}", ln=1)

    pdf.output(f"PDFs/{inv_num}.pdf")




