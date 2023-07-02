import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Get the file paths of all the XLSX files in the "invoices" directory
filepaths = glob.glob('invoices/*.xlsx')

# Iterate over each file path
for filepath in filepaths:
    # Create a new PDF document
    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()

    # Extract the file name and invoice details
    filename = Path(filepath).stem
    invoice_no, invoice_date = filename.split('-')

    # Set the font style and size for the header
    pdf.set_font(style='B', family='Times', size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice no: {invoice_no}", ln=1)
    pdf.cell(w=1, h=12, txt=f"Date: {invoice_date}", ln=1)

    # Read the Excel file into a DataFrame
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # Format column names for the PDF table
    columns = df.columns
    columns = [item.replace('_', ' ').title() for item in columns]

    # Set the font style and size for the table headers
    pdf.set_text_color(80, 80, 80)
    pdf.set_font(family='Times', size=10, style='B')

    # Add the column headers to the PDF table
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=32, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        # Set the font style and size for the table content
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)

        # Add the data cells to the PDF table
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=32, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    # Calculate the total price
    total_price = df['total_price'].sum()

    # Add the total price to the PDF table
    pdf.set_font(family='Times', size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=70, h=8, txt='', border=1)
    pdf.cell(w=32, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)

    # Add the sum total label to the PDF
    pdf.set_text_color(80, 80, 80)
    pdf.set_font(family='Times', size=12)
    pdf.cell(w=30, h=10, txt=f'Sum Total: {total_price}', ln=1)

    # Add the invoice services label and image to the PDF
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=30, h=10, txt='Invoice services')
    pdf.image('invoice.png', w=10)

    # Save the PDF with the same name as the input file
    pdf.output(f'pdf_invoice/{filename}.pdf')
