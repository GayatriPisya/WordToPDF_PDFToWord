from converters.word_to_pdf import word_to_pdf
from converters.pdf_to_word import pdf_to_word

def main():
    print("Choose an option:")
    print("1. Convert Word to PDF")
    print("2. Convert PDF to Word")
    choice = input("Enter your choice (1/2): ")

    if choice == "1":
        word_file = input("Enter the path of the Word file: ")
        output_pdf = input("Enter the output PDF file name: ")
        word_to_pdf(word_file, output_pdf)
    elif choice == "2":
        pdf_file = input("Enter the path of the PDF file: ")
        output_docx = input("Enter the output Word file name: ")
        pdf_to_word(pdf_file, output_docx)
    else:
        print("Invalid choice!")

if __name__ == "__main__":
    main()
