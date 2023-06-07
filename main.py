import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    df = pd.read_excel(filepath)

    pdf = FPDF(orientation="p", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_num, date = filename.split("-")

    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=50, h=12, txt=f"Invoice num : {invoice_num}", ln=1)

    pdf.set_font(family="Times", style="B")
    pdf.cell(w=50, h=12, txt=f"Date: {date}", ln=1)

    #add a header
    pdf.set_font(family="Times", size=15, style="B")
    pdf.set_text_color(90, 90, 90)

    columns = df.columns

    columns = [item.replace("_", " ").title() for item in columns]


    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=43, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=50, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=35, h=8, txt=str(columns[4]), border=1, ln=1)



    # Add table Rows

    for index,row in df.iterrows():

        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(90, 90, 90)

        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=43, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add total price

    total_sum = df["total_price"].sum()

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(90, 90, 90)

    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=43, h=8,  border=1)
    pdf.cell(w=50, h=8,  border=1)
    pdf.cell(w=40, h=8, border=1)
    pdf.cell(w=35, h=8, txt=str(total_sum), border=1, ln=1)

    # Add logo and total sum

    pdf.set_font(family="Times", size=10, style="B")

    pdf.cell(w=30, h=8, txt=f"Total price is {total_sum}.", ln=2)

    pdf.set_font(family="Times", size=13, style="B")

    pdf.cell(w=30, h=8, txt="ModernMonk27")
    pdf.image("spirit.png", w=10)


    pdf.output(f"PDFs/{filename}.pdf")

