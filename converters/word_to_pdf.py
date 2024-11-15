from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

def word_to_pdf(input_path, output_path):
    try:
        doc = Document(input_path)
        c = canvas.Canvas(output_path, pagesize=letter)
        width, height = letter

        for paragraph in doc.paragraphs:
            text = paragraph.text
            c.drawString(72, height - 72, text)
            height -= 18  # Adjust for the next line
            if height < 72:  # Start a new page
                c.showPage()
                height = letter[1] - 72

        c.save()
        print(f"Successfully converted '{input_path}' to '{output_path}'.")
    except Exception as e:
        print(f"Error: {e}")
