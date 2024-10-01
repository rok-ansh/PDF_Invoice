import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
# print(filepaths)

for filepath in filepaths:
    # Pdf creation
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # getting the filename without its extention
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    # date = filename.split("-")[1]

    # Creating a Invoice number
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr. {invoice_nr}", ln=1)

    # Creating a date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)
    pdf.cell(w=30, h=8, txt="", ln=1)

    # Reading the Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Adding the header
    # Converting the columns name in list format as by default its in Index format
    # columns = list(df.columns)
    columns = df.columns
    # Replacing the _ with space and making it Caps
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[0], border=1)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=30, h=8, txt=columns[0], border=1, ln=1)


    # Adding the rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add the total price
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=70, h=8, txt="")
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)
    pdf.cell(w=30, h=8, txt="", ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"The Total Price is {total_sum}", ln=1)

    # Add Company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt="PythonWorld")
    pdf.image("pythonhow.png", w=8)

    pdf.output(f"PDFS/{filename}.pdf")
