import PyPDF2 as pp2

class PdfMerger:
    def __init__(self):
        pass

    @staticmethod
    def merge_pdf(input_files, output_file):
        merger = pp2.PdfMerger()
        for input_file in input_files:
            merger.append(input_file)
        merger.write(output_file)
        merger.close()


if __name__ == '__main__':
    input_files = ['../../sample_data/pdf_sample/Sample C_01.pdf', '../../sample_data/pdf_sample/Sample C_02.pdf', '../../sample_data/pdf_sample/Sample C_03.pdf']
    PdfMerger.merge_pdf(input_files, '../../output.pdf')