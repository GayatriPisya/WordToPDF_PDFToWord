from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import os
import subprocess
from io import BytesIO

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

@app.route('/arrange')
def arrange():
    return render_template('arrange.html')
# Route for Image to PDF
@app.route('/image-to-pdf')
def image_to_pdf():
    return render_template('image_to_pdf.html')


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

@app.route('/compress-image', methods=['POST'])
def compress_image():
    file = request.files['file']
    compression_level = int(request.form['compression_level'])

    if file:
        # Save the uploaded file
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Open the image file with Pillow
        img = Image.open(file_path)

        # Adjust the quality based on the selected compression level
        if compression_level == 25:
            quality = 25
        elif compression_level == 50:
            quality = 50
        elif compression_level == 75:
            quality = 75
        else:
            quality = 100  # No compression

        # Compress the image and save it to a new location
        compressed_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"compressed_{filename}")
        img.save(compressed_file_path, quality=quality)

        # Return the compressed image to the user for download
        return send_file(compressed_file_path, as_attachment=True)
    
@app.route('/arrange', methods=['POST'])
def arrange_files():
    files = request.files.getlist('files')
    file_list = []
    for file in files:
        if file and file.filename.endswith('.pdf'):
            temp_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(temp_path)
            file_list.append({'filename': file.filename, 'path': temp_path})

    return render_template('arrange.html', files=file_list)

@app.route('/merge', methods=['POST'])
def merge_pdfs():
    from PyPDF2 import PdfMerger

    file_order = request.form.getlist('file_order[]')
    merger = PdfMerger()

    for filename in file_order:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        merger.append(file_path)

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'merged.pdf')
    merger.write(output_path)
    merger.close()

    return send_file(output_path, as_attachment=True)

# Route for Image to PDF Conversion
@app.route('/convert-image', methods=['POST'])
def convert_image_to_pdf():
    if 'image' not in request.files:
        return "No file part", 400

    image_file = request.files['image']
    
    if image_file.filename == '':
        return "No selected file", 400

    if image_file:
        # Save the uploaded image
        filename = secure_filename(image_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        image_file.save(file_path)

        # Convert the image to PDF
        try:
            image = Image.open(file_path).convert('RGB')  # Ensure the image is in RGB format
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            image.save(pdf_path, "PDF", quality=100)

            # Remove the original image file to save space
            os.remove(file_path)

            # Send the PDF file back to the user
            return send_file(pdf_path, as_attachment=True)

        except Exception as e:
            print(f"Error: {e}")
            return "Error converting image to PDF", 500

    return "Invalid request", 400

if __name__ == '__main__':
    app.run(debug=True)
