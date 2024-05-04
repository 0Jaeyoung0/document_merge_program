import win32com.client

class ToPdfConverter:
    def __init__(self):
        pass

    def word2pdf(self, input_file, output_file):
        doc_file = input_file.replace('/', '\\')

        wdFormatPDF = 17

        word = win32com.client.Dispatch('Word.Application')
        doc = word.Documents.Open(doc_file)

        doc_output_file = output_file.replace('/', '\\')

        doc.SaveAs(doc_output_file, FileFormat=wdFormatPDF)

        doc.Close()
        word.Quit()

    def excel2pdf(self, input_file, output_file):
        excel = win32com.client.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(input_file)

        for ws in wb.Worksheets:
            ws.Select()
            wb.ActiveSheet.ExportAsFixedFormat(0, output_file)

        wb.Close()
        excel.Quit()

    def convert_to_pdf(self, input_file, output_file):
        if input_file.endswith('.docx') or input_file.endswith('.doc'):
            self.word2pdf(input_file, output_file)
        elif input_file.endswith('.xlsx') or input_file.endswith('.xls'):
            self.excel2pdf(input_file, output_file)
        else:
            print("::ERROR::")
