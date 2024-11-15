from flask import Flask, request, send_from_directory, jsonify
from docx2pdf import convert
from pdf2docx import Converter
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

# Create folders if they don't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return "Server is running!"

# Route for handling Word to PDF and PDF to Word conversion
@app.route('/convert', methods=['POST'])
def convert_file():
    try:
        file = request.files['file']
        convert_type = request.form['convert_type']

        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400

        # Save uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # Conversion logic
        output_file_path = ''
        if convert_type == 'word_to_pdf':
            # Convert Word to PDF
            output_file_path = os.path.join(OUTPUT_FOLDER, file.filename.rsplit('.', 1)[0] + '.pdf')
            convert(file_path, output_file_path)
        elif convert_type == 'pdf_to_word':
            # Convert PDF to Word
            output_file_path = os.path.join(OUTPUT_FOLDER, file.filename.rsplit('.', 1)[0] + '.docx')
            cv = Converter(file_path)
            cv.convert(output_file_path, start=0, end=None)

        # Check if the output file is created successfully
        if os.path.exists(output_file_path):
            return jsonify({"message": "Conversion successful", "file_path": output_file_path}), 200
        else:
            return jsonify({"error": "Conversion failed"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Route to serve the converted file for download
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
