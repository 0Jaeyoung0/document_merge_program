import src.file_io.file_io as file_io
import src.convert_document.convert_to_pdf as convert_to_pdf
import src.handle_document.handle_docx as handle_docx
import src.handle_document.handle_pdf as handle_pdf
import src.merge_document.merge_pdf as merge_pdf

from tkinter import *
from tkinter.messagebox import askokcancel

import tempfile

import os

class App:
    def __init__(self):
        self.input_files = []
        self.selected_pages = []
        self.file_selector = file_io.FileIO()
        self.to_pdf_converter = convert_to_pdf.ToPdfConverter()
        self.pdf_handler = handle_pdf.PdfHandler()
        self.word_handler = handle_docx.WordHandler()
        self.pdf_merger = merge_pdf.PdfMerger()

    def btn_load_click(self):
        files = self.file_selector.open_files()
        for file in files:
            file_basename = os.path.basename(file)
            self.input_files.append(file)
            listbox.insert(END, file_basename)

    def btn_delete_click(self):
        selected_items = listbox.curselection()
        if askokcancel("Delete", f"Are you sure you want to delete it?"):
            for i in reversed(selected_items):
                listbox.delete(i)
                del self.input_files[i]

    def btn_merge_click(self):
        file_names_without_ext = []
        converted_files = []

        with tempfile.TemporaryDirectory() as temp_dir:
            # 파일 일괄 변환
            for input_file, selected_page in self.input_files, self.selected_pages:
                # 임시 파일 경로 생성
                file_name = os.path.basename(input_file)
                file_name_without_ext = os.path.splitext(file_name)[0]  # 확장명을 제외한 파일명 가져오기
                converted_file = os.path.join(temp_dir, f'{file_name_without_ext}.pdf')

                if input_file.endswith('.doc') or input_file.endswith('.docx'):
                    self.to_pdf_converter.word2pdf(input_file, converted_file)

                    if selected_page != 0:
                        self.pdf_handler.extract_page(converted_file, selected_page, converted_file)
                elif input_file.endswith('.xls') or input_file.endswith('.xlsx'):
                    if selected_page != 0:
                        self.to_pdf_converter.excel2pdf(input_file, converted_file, selected_page)
                    else:
                        self.to_pdf_converter.excel2pdf(input_file, converted_file)
                else:
                    print("error")

                file_names_without_ext.append(file_name_without_ext)
                converted_files.append(converted_file)

            # 각 파일 당 페이지 수 저장
            files_page_num = []

            for converted_file in converted_files:
                page_num = self.pdf_handler.extract_page_num(converted_file)
                files_page_num.append(page_num)

            # 입력 받기
            title = entry_title.get()
            dept_name = entry_department.get()
            person_name = entry_responsible.get()

            cover_page_docx_path = os.path.join(temp_dir, "cover_page.docx")
            index_page_docx_path = os.path.join(temp_dir, "index_page.docx")

            # 커버, 목차 페이지 생성
            self.word_handler.create_cover_page(title, dept_name, person_name, cover_page_docx_path)
            self.word_handler.create_index_page(file_names_without_ext, files_page_num, index_page_docx_path)

            cover_page_pdf_path = os.path.join(temp_dir, "cover_page.pdf")
            index_page_pdf_path = os.path.join(temp_dir, "index_page.pdf")

            # 커버, 목차 페이지 변환
            self.to_pdf_converter.convert_to_pdf(cover_page_docx_path, cover_page_pdf_path)
            self.to_pdf_converter.convert_to_pdf(index_page_docx_path, index_page_pdf_path)

            # 커버, 목차 페이지 페이지 수 저장
            cover_page_num = self.pdf_handler.extract_page_num(cover_page_pdf_path)
            index_page_num = self.pdf_handler.extract_page_num(index_page_pdf_path)

            # 커버, 목차 페이지를 배열 맨 앞으로 이동
            converted_files.insert(0, index_page_pdf_path)
            converted_files.insert(0, cover_page_pdf_path)

            # 커버, 목차 페이지 수를 배열 맨 앞으로 이동
            files_page_num.insert(0, index_page_num)
            files_page_num.insert(0, cover_page_num)

            merged_file = os.path.join(temp_dir, "merged.pdf")

            # 파일 일괄 병합
            self.pdf_merger.merge_pdf(input_files=converted_files, output_file=merged_file)

            page_file = os.path.join(temp_dir, "page.pdf")

            # 페이지 번호 추가
            self.pdf_handler.insert_page_number(merged_file, page_file, 1)

            output_file = self.file_selector.save_file()

            # 북마크 추가
            file_names_without_ext.insert(0, "index_page")
            file_names_without_ext.insert(0, "cover_page")
            self.pdf_handler.add_bookmark(page_file, file_names_without_ext, files_page_num, output_file)


if __name__ == '__main__':
    app = App()

    root = Tk()
    root.title("N&M")
    root.resizable(width=False, height=False)
    root.geometry("600x400+100+100")

    # 파일 목록 박스
    left_frame = Frame(root)
    left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)
    listbox = Listbox(left_frame, width=40, height=15, font=("Arial", 12))
    listbox.pack(side=LEFT, fill=BOTH, expand=True)

    # 우측 프레임
    right_frame = Frame(root)
    right_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=10, pady=10)

    # 버튼
    btn_frame = Frame(right_frame)
    btn_frame.pack(side=TOP, pady=10)
    btn_load = Button(btn_frame, text='파일 로드', width=9, font=("Arial", 12), command=app.btn_load_click)
    btn_delete = Button(btn_frame, text='파일 삭제', width=9, font=("Arial", 12), command=app.btn_delete_click)
    btn_merge = Button(right_frame, text='파일 병합', width=12, font=("Arial", 12), command=app.btn_merge_click)
    btn_load.pack(side=LEFT, padx=5)
    btn_delete.pack(side=LEFT, padx=5)
    btn_merge.pack(side=BOTTOM, pady=10)

    # 입력 폼
    form_frame = Frame(right_frame)
    form_frame.pack(side=TOP, pady=10)
    lbl_title = Label(form_frame, text="제목:", font=("Arial", 12))
    entry_title = Entry(form_frame, width=30, font=("Arial", 12))
    lbl_department = Label(form_frame, text="부서명:", font=("Arial", 12))
    entry_department = Entry(form_frame, width=30, font=("Arial", 12))
    lbl_responsible = Label(form_frame, text="담당자:", font=("Arial", 12))
    entry_responsible = Entry(form_frame, width=30, font=("Arial", 12))
    lbl_title.grid(row=0, column=0, sticky=W, padx=5, pady=5)
    entry_title.grid(row=0, column=1, padx=5, pady=5)
    lbl_department.grid(row=1, column=0, sticky=W, padx=5, pady=5)
    entry_department.grid(row=1, column=1, padx=5, pady=5)
    lbl_responsible.grid(row=2, column=0, sticky=W, padx=5, pady=5)
    entry_responsible.grid(row=2, column=1, padx=5, pady=5)

    root.mainloop()