import tkinter as tk
from tkinter import filedialog
import os

class FileIO:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def open_files(self):
        file_paths = []

        while True:
            file_path = filedialog.askopenfilenames(initialdir=f'{os.getcwd()}', title='Merged File Selection Window',
                                                    filetypes=[('ALL', '*.docx'), ('ALL', '*.doc'), ('ALL', '*.xlsx'), ('ALL', '*.xls'), ('ALL', '*.pdf')])
            if not file_path:
                break
            for path in file_path:
                file_paths.append(path)
                print(path)

        print(f"file_paths : {file_paths}")

        return file_paths

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialdir=f'{os.getcwd()}', title='File Storage Location Selection Window',
                                                 filetypes=[('PDF', '*.pdf')])
        return file_path