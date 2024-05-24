import win32com.client

class ToPdfConverter:
    def __init__(self):
        self.word = win32com.client.Dispatch('Word.Application')
        self.excel = win32com.client.Dispatch('Excel.Application')

    def __del__(self):
        if self.word:
            self.word.Quit()
        if self.excel:
            self.excel.Quit()

    def word2pdf(self, input_file, output_file):
        wdExportFormatPDF = 17

        doc_file = input_file.replace('/', '\\')
        doc = self.word.Documents.Open(doc_file)

        doc.ExportAsFixedFormat(output_file, wdExportFormatPDF)

        doc.Close()

    def excel2pdf(self, input_file, output_file, sheet_name=None):
        xlExportFormatPDF = 0

        wb = self.excel.Workbooks.Open(input_file)

        if sheet_name is None:
            ws = wb.Worksheets(sheet_name)
            ws.ExportAsFixedFormat(xlExportFormatPDF, output_file)
        else:
            wb.ExportAsFixedFormat(xlExportFormatPDF, output_file)

        wb.Close()

    def convert_to_pdf(self, input_file, output_file, sheet_name=None):
        if input_file.endswith('.docx') or input_file.endswith('.doc'):
            self.word2pdf(input_file, output_file)
        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            self.excel2pdf(input_file, output_file, sheet_name)
        else:
            print("::ERROR::")

if __name__ == '__main__':
    converter = ToPdfConverter()
    converter.convert_to_pdf('../../sample_data/xlsx_sample/test3.xlsx', '../../sample_data/xlsx_sample/test3.pdf')
