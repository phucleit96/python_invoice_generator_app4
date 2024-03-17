import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*xlsx')
for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr = str(filename).split("-")[0]
    date = str(filename).split("-")[1]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", align="L", ln=1)

    pdf.cell(w=0, h=12, txt=f"Date: {date}", align="L", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    df_columns = df.columns.tolist()
    for column in df_columns:
        column_split = [word.title() for word in column.split("_")]
        column_name = " ".join(column_split)
        pdf.set_font(family="Times", style="B", size=12)
        pdf.cell(w=40, h=10, txt=column_name, border=1)
    pdf.ln()
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        for item in row:
            pdf.cell(w=40, h=10, txt=str(item), border=1)
        pdf.ln()


    pdf.output(f"invoices_pdf/{filename}.pdf")

