from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os
import subprocess

app = Flask(__name__)

# Set the upload folder and allowed file extensions
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)  # Ensure uploads folder exists


@app.route('/')
def home():
    return render_template('base.html') 

# Route for Word to PDF
@app.route('/word-to-pdf')
def word_to_pdf():
    return render_template('word_to_pdf.html')

# Route for PDF to Word
@app.route('/pdf-to-word')
def pdf_to_word():
    return render_template('pdf_to_word.html')

# Route for File Compression
@app.route('/compress')
def compress():
    return render_template('compress.html')

# Route for Merge PDFs
@app.route('/merge')
def merge():
    return render_template('merge.html')

# Route for Image to PDF
@app.route('/image-to-pdf')
def image_to_pdf():
    return render_template('image_to_pdf.html')

# Route for Split PDFs
@app.route('/split-pdfs')
def split_pdfs():
    return render_template('split_pdfs.html')

# Route to handle file upload and conversion
@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    conversion_type = request.form['conversion_type']

    if file:
        # Ensure the upload directory exists
        upload_folder = app.config['UPLOAD_FOLDER']
        os.makedirs(upload_folder, exist_ok=True)  # Create the directory if it doesn't exist
        
        # Save the file
        filename = secure_filename(file.filename)
        file_path = os.path.join(upload_folder, filename)
        file.save(file_path)
        print(f"Saving file to: {file_path}")  # For debugging purposes

        # Handle conversion type
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
    try:
        # Correct path to soffice.exe (64-bit version of LibreOffice)
        libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"

        # Run the LibreOffice command to convert Word to PDF
        command = [
            libreoffice_path,
            '--headless',  # Runs LibreOffice in headless mode (without GUI)
            '--convert-to', 'pdf',  # Specify PDF output
            '--outdir', os.path.dirname(output_path),  # Set output directory
            input_path  # Input file path
        ]

        # Run the command
        result = subprocess.run(command, check=True)

        # Check if output PDF was created
        if os.path.exists(output_path):
            print(f"PDF successfully created at {output_path}")
        else:
            print("Error: PDF file not created")
    except Exception as e:
        print(f"Error while converting Word to PDF: {e}")

# Function to convert PDF to Word (example using pdf2docx)
def convert_pdf_to_word(input_path, output_path):
    try:
        from pdf2docx import Converter
        converter = Converter(input_path)
        converter.convert(output_path, start=0, end=None)  # Convert the entire PDF
        converter.close()
        print(f"Converted {input_path} to {output_path} successfully.")
    except Exception as e:
        print(f"Error converting PDF to Word: {e}")

if __name__ == '__main__':
    app.run(debug=True)
