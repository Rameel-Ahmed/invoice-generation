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
    invoice_nr, date = filename.split("-")

    # Invoice header
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    pdf.ln(10)  # Add some space

    # Table header
    columns = df.columns
    columns = [item.replace(" ", "_").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B") 
    pdf.set_text_color(80, 80, 80)
    
    # Adjusted column widths (total width ~190mm for A4 portrait)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)  # product_id
    pdf.cell(w=60, h=8, txt=columns[1], border=1)  # product_name (wider)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)  # amount_purchased (wider)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)  # price_per_unit
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)  # total_price

    # Table rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Total sum row
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Total price text
    pdf.ln(10)  # Add some space
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=0, h=10, txt=f"The total price is {total_sum}", ln=1, align="R")

    pdf.output(f"PDFs/{filename}.pdf")