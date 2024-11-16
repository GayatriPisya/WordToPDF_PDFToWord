from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from converters.word_to_pdf import word_to_pdf
from converters.pdf_to_word import pdf_to_word

app = Flask(__name__)

# Configure upload and converted directories
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["CONVERTED_FOLDER"] = CONVERTED_FOLDER

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/upload', methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file part", 400

    file = request.files["file"]
    if file.filename == "":
        return "No selected file", 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    conversion_type = request.form.get("conversion_type")
    output_file = os.path.join(app.config["CONVERTED_FOLDER"], f"converted_{file.filename}")

    try:
        if conversion_type == "word_to_pdf":
            word_to_pdf(file_path, output_file)
        elif conversion_type == "pdf_to_word":
            output_file = output_file.replace(".docx", ".pdf")
            pdf_to_word(file_path, output_file)
        else:
            return "Invalid conversion type", 400

        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return f"Conversion failed: {e}", 500

if __name__ == "__main__":
    app.run(debug=True)
