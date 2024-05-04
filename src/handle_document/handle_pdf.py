import PyPDF2
from reportlab.pdfgen import canvas
import os

class PdfHandler():
    temp_file_name = 'temp.pdf'

    def __init__(self):
        pass

    @staticmethod
    def get_number_format(number):
        return '-' + str(number) + '-'

    @staticmethod
    def create_number_pdf(number):
        text = PdfHandler.get_number_format(number)

        new_pdf = canvas.Canvas(PdfHandler.temp_file_name)
        new_pdf.drawString(x=10, y=10, text=text)
        new_pdf.save()

    def insert_page_number(self, input_file, output_file, start_page_number=1):
        output_pdf = PyPDF2.PdfWriter()

        with open(input_file, 'rb') as input_stream:
            origin_pdf = PyPDF2.PdfReader(input_stream)

            for page in origin_pdf.pages:
                PdfHandler.create_number_pdf(start_page_number)

                with open(PdfHandler.temp_file_name, 'rb') as temp_stream:
                    number_pdf = PyPDF2.PdfReader(temp_stream)
                    page.merge_page(number_pdf.pages[0])
                    output_pdf.add_page(page)

                start_page_number += 1
                os.remove(PdfHandler.temp_file_name)

        with open(output_file, 'wb') as output_stream:
            output_pdf.write(output_stream)

if __name__ == '__main__':
    handler = PdfHandler()
    handler.insert_page_number('../../sample_data/pdf_sample/Sample C_01.pdf', '../../output.pdf')
