from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from converters.word_to_pdf import word_to_pdf
from converters.pdf_to_word import pdf_to_word
import io

app = Flask(__name__, static_folder="static")

# Allowed file extensions
ALLOWED_EXTENSIONS = {"pdf", "docx"}

def allowed_file(filename):
    """Check if the uploaded file has a valid extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/")
def index():
    """Render the main landing page."""
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    """Handle file uploads and conversion."""
    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    if file.filename == "":
        return "No file selected", 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        conversion_type = request.form.get("conversion_type")

        try:
            # Create an in-memory output file
            output = io.BytesIO()
            
            if conversion_type == "word_to_pdf":
                word_to_pdf(file, output)  # Process Word to PDF
                output_filename = filename.rsplit(".", 1)[0] + ".pdf"
            elif conversion_type == "pdf_to_word":
                pdf_to_word(file, output)  # Process PDF to Word
                output_filename = filename.rsplit(".", 1)[0] + ".docx"
            else:
                return "Invalid conversion type", 400

            # Set pointer to the beginning of the output
            output.seek(0)
            return send_file(output, as_attachment=True, download_name=output_filename)
        except Exception as e:
            return f"Conversion failed: {e}", 500
    else:
        return "Invalid file type", 400

if __name__ == "__main__":
    app.run(debug=True)
