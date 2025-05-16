import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("*.xlsx")

for filepath in filepaths:

    filename = Path(filepath).stem
    invoice_num, invoice_date = filename.split("-")

    df_cust = pd.read_excel(filepath, sheet_name="Customer")
    columns_cust = [item.replace("_", " ").title() for item in df_cust.columns]
    #for i, row in df_cust.iterrows():
    row = df_cust.iloc[0]
    customer_id = row["customer_id"]
    customer_name = row["customer_name"]

    #customer_id, customer_name = df_cust

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=140,h=8, txt=f"Customer: {customer_name}")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Invoice: {invoice_num}", align="R", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=140,h=8, txt=f"Cust Id: {customer_id}")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50,h=8, txt=f"Date: {invoice_date}", align="R", ln=1)

    df = pd.read_excel(filepath, sheet_name="Invoice")
    # print(df)

    # Add header
    #columns = df.columns
    #print(type(df.columns))
    #columns = df.columns
    #print(type(columns))
    columns = [item.replace("_", " ").title() for item in df.columns]
    #print(type(columns))
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=24, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=32, h=8, txt=columns[2], border=1, align="R")
    pdf.cell(w=32, h=8, txt=columns[3], border=1, align="R")
    pdf.cell(w=32, h=8, txt=columns[4], border=1, align="R", ln=1)

    # Add rows to the table
    for i, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=24, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=32, h=8, txt=str(row["amount_purchased"]), border=1, align="R")
        pdf.cell(w=32, h=8, txt=str(row["price_per_unit"]), border=1, align="R")
        pdf.cell(w=32, h=8, txt=str(row["total_price"]), border=1, ln=1, align="R")

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=24, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=32, h=8, txt="", border=1)
    pdf.cell(w=32, h=8, txt="", border=1)
    pdf.cell(w=32, h=8, txt=str(total_sum), border=1, align="R", ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=40, h=8, txt=f"The total price is {total_sum}", ln=1)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=20, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)


    pdf.output(f"{invoice_num}-invoice.pdf")
