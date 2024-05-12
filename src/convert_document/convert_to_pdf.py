import os
from docx2pdf import convert
import comtypes.client
import xlwings as xw
from PyPDF2 import PdfReader, PdfWriter

class ToPdfConverter:
    def __init__(self):
        pass

    def extract_page(self, input_pdf, document_file):
        pdf_reader = PdfReader(input_pdf)
        pdf_writer = PdfWriter()

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
                            print(f"Saved page {page_num+1}.")
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


    def docx2pdf(self, input_file, output_file):
        convert(input_file, output_file)

    def doc2pdf(self, input_file, output_file):
        doc_file = input_file.replace('/', '\\')

        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False

        doc = word.Documents.Open(doc_file)

        doc_output_file = output_file.replace('/', '\\')

        doc.SaveAs(doc_output_file, FileFormat=17)
        doc.Close()
        word.Quit()

    def excel2pdf(self, input_file, output_file):
        app = xw.App(visible=False)
        sheet_name = []
        try:
            book = xw.Book(input_file)
            for sheet in book.sheets:
                sheet_name.append(sheet.name)
                report_sheet = book.sheets[0]
                report_sheet.api.ExportAsFixedFormat(0, output_file)
            # if(len(book.sheets) > 1):
            #     sheet_number = input(f"Please write down the order of one sheet you want[A blank will bring the first sheet]\n Excel : {book.name} -> Sheets : {sheet_name}")
            #     if sheet_number:
            #         print(f"{sheet_name[int(sheet_number) - 1]} sheet saved.")
            #         report_sheet = book.sheets[int(sheet_number)-1]
            #     else:
            #         print("You entered a space to get the first sheet.")
            #         report_sheet = book.sheets[0]
            #
            #     report_sheet.api.ExportAsFixedFormat(0, output_file)
            # else:
            #     print(f"{book.sheets[0].name} sheet saved.")
            #     report_sheet = book.sheets[0]
            #     report_sheet.api.ExportAsFixedFormat(0, output_file)

        finally:
            app.quit()

    def convert_to_pdf(self, input_file, output_file):
        if input_file.endswith('.docx'):
            self.docx2pdf(input_file, output_file)
        elif input_file.endswith('.doc'):
            self.doc2pdf(input_file, output_file)
        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            self.excel2pdf(input_file, output_file)
        else:
            print("::ERROR::")