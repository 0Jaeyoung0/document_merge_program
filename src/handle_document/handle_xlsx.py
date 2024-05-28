import pandas as pd


class ExcelHandler(object):
    def __init__(self):
        pass

    @staticmethod
    def extract_sheet_names(input_file):
        try:
            sheet_names = pd.ExcelFile(input_file).sheet_names
            return sheet_names
        except Exception as e:
            print(f"Error: {e}")
            return []