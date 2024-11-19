from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Route for the homepage
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload and conversion
@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    conversion_type = request.form['conversion_type']

    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        if conversion_type == 'word-to-pdf':
            output_path = file_path.replace('.docx', '.pdf')
            convert_word_to_pdf(file_path, output_path)
        elif conversion_type == 'pdf-to-word':
            output_path = file_path.replace('.pdf', '.docx')
            convert_pdf_to_word(file_path, output_path)
        else:
            return "Unsupported conversion type", 400

        return send_file(output_path, as_attachment=True)

# Function to convert Word to PDF
def convert_word_to_pdf(input_path, output_path):
    from fpdf import FPDF  # Lightweight alternative for PDF generation
    pdf = FPDF()
    doc = Document(input_path)
    pdf.set_auto_page_break(auto=True, margin=15)

    for para in doc.paragraphs:
        pdf.add_page()
        pdf.set_font('Arial', size=12)
        pdf.multi_cell(0, 10, para.text)

    pdf.output(output_path)

# Function to convert PDF to Word
def convert_pdf_to_word(input_path, output_path):
    converter = Converter(input_path)
    converter.convert(output_path, start=0, end=None)
    converter.close()

if __name__ == '__main__':
    app.run(debug=True)
