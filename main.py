import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    print(df)

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Get file name using pathlib
    filename = Path(filepath).stem
    # Split will split on first occurrence
    invoice_nr, invoice_date = filename.split('-')

    # Set title and date
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice nr. {invoice_nr}', ln=1)
    pdf.cell(w=50, h=8, txt=f'Date {invoice_date}', ln=1)
    pdf.ln()

    # Set table column headers
    columns = list(df.columns)
    # List comprehension to format the column headers
    columns = [item.replace('_', ' ').title() for item in columns]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=8, txt=columns[0], border=True)
    pdf.cell(w=60, h=8, txt=columns[1], border=True)
    pdf.cell(w=35, h=8, txt=columns[2], border=True)
    pdf.cell(w=30, h=8, txt=columns[3], border=True)
    pdf.cell(w=30, h=8, txt=columns[4], border=True, ln=1)

    # Draw rows
    total = 0
    pdf.set_font(family='Times', size=10)
    for index, row in df.iterrows():
        print(type(row['amount_purchased']))
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=True)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=True)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), border=True)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=True)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=True, ln=1)
        total += row['total_price']

    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=60, h=8, border=True)
    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=30, h=8, border=True)
    pdf.cell(w=30, h=8, txt=str(total), border=True, ln=1)

    pdf.output(f'PDFs/{filename}.pdf')
