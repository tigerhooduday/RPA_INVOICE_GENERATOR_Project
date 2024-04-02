from flask import Flask, render_template, request, flash, send_file
from werkzeug.utils import secure_filename
from io import BytesIO
import pandas as pd
from fpdf import FPDF
import os
import tempfile


app = Flask(__name__, static_url_path='/static', static_folder='static')
app.secret_key = 'your_secret_key'  # Change this to a random string

# Function to parse Excel file
def parse_excel(file):
    df = pd.read_excel(file)
    products = df.iloc[:, 0].tolist()
    amounts = df.iloc[:, 1].tolist()
    return products, amounts

# Function to generate PDF invoice
def generate_invoice(products, amounts, company_name, customer_name, phone_number, date):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Title
    pdf.set_font("Arial", style="B", size=15)
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(200, 5, txt="INVOICE", ln=True, align="C", fill=True)
    pdf.cell(200, 1, txt="_______________________________________________________", ln=True, align="C")
    pdf.ln(5)

    # Company Name
    pdf.cell(200, 5, txt=f'"{company_name} COMPANY"', ln=True, align="C")

    pdf.set_font("Arial", style="I", size=9)  
    pdf.cell(200, 5, txt="  NH05, Chandigarh-Ludhiana Highway,Gharuan, Mohali,Punjab.", ln=True, align="C")
    pdf.set_font("Arial", style="", size=12) 
    pdf.set_text_color(0, 0, 0) 
    pdf.ln(6)

    # Customer Details
    pdf.cell(200, 5, txt=f"Bill to: {customer_name}", ln=True, align="l")
    pdf.cell(200, 5, txt=f"Contact Detail: {phone_number}", ln=True, align="l")
    pdf.cell(200, 5, txt=f"Date : {date}", ln=True, align="l")
    pdf.ln(8)

    # Table Header
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(70, 6, "Product", 1, 0, "C", fill=True)
    pdf.cell(70, 6, "Rate", 1, 1, "C", fill=True)

    # Table Data
    pdf.set_fill_color(255, 255, 255)
    for product, amount in zip(products, amounts):
        pdf.cell(70, 10, product, 1, 0, "L")
        pdf.cell(70, 10, str(amount), 1, 1, "C")

    pdf.ln(1)
    total = sum(amounts)
    pdf.cell(70, 10, "Total amount", 1, 0, "C")
    pdf.cell(70, 10, f"{total:.2f}", 1, 1, "C")
    tax = sum(amounts) * 0.18
    total = sum(amounts) + tax
    pdf.cell(100, 10, txt=f"GST 18% : {tax:.2f}" )
    pdf.ln(7)
    pdf.cell(100, 10, txt=f"Amount to Pay : {total:.2f}" )
    pdf.ln(12)
    pdf.cell(100, 10, txt=f"Thank you for your purchase. We appreciate your business! Have a nice day" )

    temp_dir = tempfile.gettempdir()
    temp_file = os.path.join(temp_dir, 'invoice.pdf')
    pdf.output(temp_file)
    return temp_file

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename == '':
            flash('No file selected. Please choose a file.', 'error')
            return render_template('index.html')
        
        customer_name = request.form['customer_name']
        phone_number = request.form['phone_number']
        date = request.form['date']
        company_name = "CUCHD RPA"
        try:
            products, amounts = parse_excel(file)
            invoice_pdf_path = generate_invoice(products, amounts, company_name, customer_name, phone_number, date)
            return send_file(invoice_pdf_path, as_attachment=True)
        except Exception as e:
            flash(f'Error processing file: {str(e)}', 'error')
            return render_template('index.html')
        
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
