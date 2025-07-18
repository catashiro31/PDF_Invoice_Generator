import numpy as np
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')
pdf = FPDF(orientation='P',unit='mm',format='A4')
for filepath in filepaths:
    filename = Path(filepath).stem
    id_user = filename.split('-')[0]
    cur_date = filename.split('-')[1]

    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    total_price = df['total_price'].sum()
    pdf.add_page()
    # Print Invoice and Date
    pdf.set_font(family='Times',style='B',size=20)
    pdf.set_text_color(0,0,0)
    pdf.cell(w=0, h=12, txt=f'Invoice nr. {id_user}',border=0,ln=1,align='L')
    pdf.cell(w=0, h=12, txt=f'Date: {cur_date}',border=0,ln=1,align='L')
    # Print Columns Name
    pdf.set_font(family='Times', style='B', size=14)
    pdf.set_text_color(80, 80, 80)
    pdf.ln(5)
    pdf.cell( w=30, h=12, txt='Product Id', border=1, ln=0, align='C')
    pdf.cell( w=70, h=12, txt='Product Name', border=1, ln=0, align='C')
    pdf.cell( w=25, h=12, txt='Amount', border=1, ln=0, align='C')
    pdf.cell( w=35, h=12, txt='Price per Unit', border=1, ln=0, align='C')
    pdf.cell( w=30, h=12, txt='Total Price', border=1, ln=0, align='C')
    # Print Table Content
    pdf.set_font(family='Times', style='', size=14)
    pdf.set_text_color(80,80,80)
    for idx, row in df.iterrows():
        pdf.ln(12)
        pdf.cell(w=30, h=12, txt=str(row['product_id']), border=1, ln=0, align='C')
        pdf.cell(w=70, h=12, txt=str(row['product_name']), border=1, ln=0, align='L')
        pdf.cell(w=25, h=12, txt=str(row['amount_purchased']), border=1, ln=0, align='C')
        pdf.cell(w=35, h=12, txt=str(row['price_per_unit']), border=1, ln=0, align='C')
        pdf.cell(w=30, h=12, txt=str(row['total_price']), border=1, ln=0, align='C')
    # Print Total Price in Table
    pdf.ln(12)
    pdf.cell(w=30, h=12, txt='', border=1, ln=0, align='C')
    pdf.cell(w=70, h=12, txt='', border=1, ln=0, align='L')
    pdf.cell(w=25, h=12, txt='', border=1, ln=0, align='C')
    pdf.cell(w=35, h=12, txt='', border=1, ln=0, align='C')
    pdf.cell(w=30, h=12, txt=str(total_price), border=1, ln=0, align='C')
    # Print Total Price out Table
    pdf.ln(17)
    pdf.set_font(family='Times', style='B',size=18)
    pdf.set_text_color(50,50,50)
    pdf.cell(w=0, h=10, txt=f'The total price is {total_price}', border=0,
             ln=1, align='L')
    pdf.cell(w=25, h=10, txt='Built by', border=0, ln=0, align='L')
    pdf.image('pythonhow.png',w=10,h=10)
pdf.output('output.pdf')