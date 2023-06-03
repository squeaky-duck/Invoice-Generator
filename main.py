from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice #{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add header
    columns = df.columns
    columns = [column.replace('_', ' ').title() for column in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, border=1, txt=columns[0])
    pdf.cell(w=65, h=8, border=1, txt=columns[1])
    pdf.cell(w=35, h=8, border=1, txt=columns[2])
    pdf.cell(w=30, h=8, border=1, txt=columns[3])
    pdf.cell(w=30, h=8, border=1, txt=columns[4], ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, border=1, txt=str(row["product_id"]))
        pdf.cell(w=65, h=8, border=1, txt=str(row["product_name"]))
        pdf.cell(w=35, h=8, border=1, txt=str(row["amount_purchased"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["price_per_unit"]))
        pdf.cell(w=30, h=8, border=1, txt=str(row["total_price"]), ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, border=1, txt="")
    pdf.cell(w=65, h=8, border=1, txt="")
    pdf.cell(w=35, h=8, border=1, txt="")
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, border=1, txt="Total", align="R")
    pdf.cell(w=30, h=8, border=1, txt=str(total_sum), ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=40, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=20, h=8, txt=f"StraightCart")
    pdf.image("cart.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")

