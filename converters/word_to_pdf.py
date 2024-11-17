from reportlab.pdfgen import canvas

def word_to_pdf(input_file, output_stream):
    """Convert Word document to PDF."""
    c = canvas.Canvas(output_stream)
    c.drawString(100, 750, "This is a placeholder for Word to PDF conversion.")
    c.save()
