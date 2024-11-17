from docx import Document

def pdf_to_word(input_file, output_stream):
    """Convert PDF to Word."""
    doc = Document()
    doc.add_paragraph("This is a placeholder for PDF to Word conversion.")
    doc.save(output_stream)
