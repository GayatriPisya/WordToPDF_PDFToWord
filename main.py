from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from converters.word_to_pdf import word_to_pdf
from converters.pdf_to_word import pdf_to_word

app = Flask(__name__)

# Configurations
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["CONVERTED_FOLDER"] = CONVERTED_FOLDER
ALLOWED_EXTENSIONS = {"pdf", "docx"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    if file.filename == "":
        return "No file selected", 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(file_path)

        conversion_type = request.form.get("conversion_type")
        output_file = os.path.join(app.config["CONVERTED_FOLDER"], f"converted_{filename}")

        try:
            if conversion_type == "word_to_pdf":
                output_file = output_file.replace(".docx", ".pdf")
                word_to_pdf(file_path, output_file)
            elif conversion_type == "pdf_to_word":
                output_file = output_file.replace(".pdf", ".docx")
                pdf_to_word(file_path, output_file)
            else:
                return "Invalid conversion type", 400

            return send_file(output_file, as_attachment=True)
        except Exception as e:
            return f"Conversion failed: {e}", 500
    else:
        return "Invalid file type", 400

if __name__ == "__main__":
    app.run(debug=True)
