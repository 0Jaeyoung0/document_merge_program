import sys

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
    def extract_page(input_file, selected_page, output_file):
        pdf_reader = PyPDF2.PdfReader(input_file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_num in selected_page:
            pdf_writer.add_page(pdf_reader.pages[page_num])  # 0부터 첫페이지임

        pdf_writer.write(output_file)



if __name__ == '__main__':
    handler = PdfHandler()
    selected_page = [1, 3]
    handler.extract_page(input_file='../../sample_data/pdf_sample/Sample C_01.pdf', selected_page=selected_page, output_file='../../sample_data/pdf_sample/test.pdf')
