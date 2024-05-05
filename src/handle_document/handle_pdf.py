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

    def insert_page_number(self, input_file, output_file, start_page_number=1):
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

if __name__ == '__main__':
    handler = PdfHandler()
    handler.insert_page_number('../../sample_data/pdf_sample/Sample C_01.pdf', '../../output.pdf')
