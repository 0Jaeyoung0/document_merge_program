import PyPDF2 as pp2

class PdfMerger:
    def __init__(self):
        pass

    def merge_pdf(self, input_files, output_file):
        merger = pp2.PdfMerger()
        for input_file in input_files:
            merger.append(input_file)
        merger.write(output_file)
        merger.close()
