import PyPDF2
from reportlab.pdfgen import canvas
import os

class PdfHandler():
    temp_file_name = 'temp.pdf'

    def __init__(self):
        pass

    @staticmethod
    def get_numbering_format(number):
        return '-' + str(number) + '-'

    @staticmethod
    def create_numbering_pdf(number, page_width, page_height):
        new_pdf = canvas.Canvas(filename=PdfHandler.temp_file_name, pagesize=(page_width, page_height))

        text_x = int(page_width / 2)
        text_y = 1

        text = PdfHandler.get_numbering_format(number)

        new_pdf.drawCentredString(x=text_x, y=text_y, text=text)
        new_pdf.save()

    @staticmethod
    def insert_page_number(input_file, output_file, start_page_number=1):
        output_pdf = PyPDF2.PdfWriter()

        with open(input_file, 'rb') as input_stream:
            origin_pdf = PyPDF2.PdfReader(input_stream)

            for page in origin_pdf.pages:
                page_width = page.mediabox.width
                page_height = page.mediabox.height

                PdfHandler.create_numbering_pdf(start_page_number, page_width, page_height)

                with open(PdfHandler.temp_file_name, 'rb') as temp_stream:
                    numbering_pdf = PyPDF2.PdfReader(temp_stream)
                    page.merge_page(numbering_pdf.pages[0])
                    output_pdf.add_page(page)

                start_page_number += 1
                os.remove(PdfHandler.temp_file_name)

        with open(output_file, 'wb') as output_stream:
            output_pdf.write(output_stream)

    @staticmethod
    def add_bookmark(input_file, title_array, page_num_array, output_file):
        reader = PyPDF2.PdfReader(input_file)
        writer = PyPDF2.PdfWriter()

        page_num = 0

        for page in range(len(reader.pages)):
            writer.add_page(reader.pages[page])

        for i in range(len(title_array)):
            writer.add_outline_item(title_array[i], page_num)
            page_num += page_num_array[i]

        with open(output_file, "wb") as output_stream:
            writer.write(output_stream)

    @staticmethod
    def extract_page_num(input_file):
        pdf_reader = PyPDF2.PdfReader(input_file)

        return len(pdf_reader.pages)

    @staticmethod
    def extract_page(input_pdf, document_file):
        pdf_reader = PyPDF2.PdfReader(input_pdf)
        pdf_writer = PyPDF2.PdfWriter()

        load_file = os.path.basename(document_file)
        if len(pdf_reader.pages) > 1:
            user_input = input(
                f"Enter the sheet number to save {load_file} [Total pages : {len(pdf_reader.pages)}] (separate with commas if multiple pages, press Enter to save the entire page): ")

            if len(user_input) > 1:
                selected_page = [int(num) - 1 for num in user_input.split(',')]
                if len(selected_page) != len(set(selected_page)):
                    print("Duplicate sheet number found")
                else:
                    for page_num in selected_page:
                        try:
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                            print(f"Saved page {page_num + 1}.")
                        except IndexError:
                            print("To save the entire page")
                            for page in pdf_reader.pages:
                                pdf_writer.add_page(page)

            elif len(user_input) == 1:
                pdf_writer.add_page(pdf_reader.pages[int(user_input) - 1])
            else:
                print("To save the entire page")
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)

            with open(input_pdf, 'wb') as out:
                pdf_writer.write(out)


if __name__ == '__main__':
    handler = PdfHandler()
    handler.insert_page_number('../../output.pdf', '../../output2.pdf')
    title_array = ['커버페이지', '목차페이지', 'Sample C_01.pdf', 'Sample C_02.pdf', 'Sample C_03.pdf']
    handler.add_bookmark('../../output.pdf', )
