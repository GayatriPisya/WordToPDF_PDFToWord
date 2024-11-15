from pdf2docx import Converter

def pdf_to_word(input_path, output_path):
    try:
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)  # Convert the entire PDF
        cv.close()
        print(f"Successfully converted '{input_path}' to '{output_path}'.")
    except Exception as e:
        print(f"Error: {e}")
