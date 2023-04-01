import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*xlsx')

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    # Get file name using pathlib
    filename = Path(filepath).stem
    # Split will split on first occurrence
    invoice_nr = filename.split('-')[0]  # index to take first item in list
    pdf.add_page()
    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice nr. {invoice_nr}')
    pdf.output(f'PDFs/{filename}.pdf')
