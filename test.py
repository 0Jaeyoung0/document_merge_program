
import src.file_io.file_io as file_io
import src.convert_document.convert_to_pdf as convert_to_pdf
import src.handle_document.handle_docx as handle_docx
import src.handle_document.handle_pdf as handle_pdf
import src.merge_document.merge_pdf as merge_pdf

from tkinter import *
from tkinter.messagebox import askokcancel
import tkinter as tk
from tkinter.ttk import Combobox

import tempfile

import os

import xlwings as xw

class App:
    def __init__(self):
        self.input_files = []
        self.order_info = []
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

        self.order_info = [0] * len(self.input_files)


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
            for input_file in self.input_files:
                # 임시 파일 경로 생성
                file_name = os.path.basename(input_file)
                file_name_without_ext = os.path.splitext(file_name)[0]  # 확장명을 제외한 파일명 가져오기
                converted_file = os.path.join(temp_dir, f'{file_name_without_ext}.pdf')

                self.to_pdf_converter.convert_to_pdf(input_file=input_file, output_file=converted_file)

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

    def btn_save_click(self):
        print('save')
        selected_indices = listbox.curselection()
        index = selected_indices[0]

        page_info = input_text.get()
        if page_info:
            self.order_info[index] = page_info

        sheet_info = combo.get()
        if sheet_info:
            self.order_info[index] = sheet_info


    def order_input(self, event):
        selected_indices = listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            self.order_info[index] = str(eval(input_text.get()))

    def radio_button_click(self):
        worksheet_names = []
        selected_item = listbox.get(tk.ACTIVE)
        selected_indices = listbox.curselection()
        index = selected_indices[0]

        # 페이지 설정한 게 있으면 불러오거나 없으면 초기화

        if var1.get() == 1:
            print('0')
        elif var1.get() == 2 or radio2.select():
            input_text.config(state="normal")
            if self.order_info[index] == 0:
                input_text.delete(0, tk.END)
            else:
                input_text.insert(0,self.order_info[index])

        elif var1.get() == 3 or radio3.select():
            combo.config(state="normal")
            app = xw.App(visible=False)
            for input_file in self.input_files:
                if input_file == selected_item:
                    book = xw.Book(input_file)
                    for sheet in book.sheets:
                        worksheet_names.append(sheet.name)
                    combo['values'] = worksheet_names
                    combo.current(0)

                    root.update_idletasks()
                    break

            app.quit()

    def toggle_widgets(self):
        selected_item = listbox.get(tk.ACTIVE)
        # Work if 0 in the array of order_info corresponding to the order of the listbox in which the user has a cursor
        selected_indices = listbox.curselection()
        index = selected_indices[0]


        if selected_item.endswith('.docx') or selected_item.endswith('.doc'):
            radio2.pack(side=tk.LEFT, padx=5)
            input_text.pack(side=tk.LEFT, padx=5)
            label2.pack(side=tk.LEFT, padx=5)
            combo.pack_forget()
            radio3.pack_forget()

            if self.order_info[index] == 0:
                var1.set(1)
                input_text.delete(0, tk.END)

            else:
                print("hi")
                var1.set(2)
                input_text.config(state="normal")
                input_text.delete(0, tk.END)  # 기존 값 삭제
                input_text.insert(0, self.order_info[index])

        elif selected_item.endswith('.xlsx') or selected_item.endswith('.xls'):
            radio3.pack(side=tk.LEFT, padx=5)
            combo.pack(side=tk.LEFT, padx=5)

            radio2.pack_forget()
            input_text.pack_forget()
            label2.pack_forget()
            if self.order_info[index] == 0:
                var1.set(1)

            else:
                var1.set(3)
                combo.config(state="normal")
                combo['values'] = [self.order_info[index]]  # 콤보박스 값 설정
                combo.current(0)  # 첫 번째 값 선택

        print(self.order_info)



if __name__ == '__main__':
    app = App()

    root = tk.Tk()
    root.title("N&M")
    root.resizable(width=False, height=False)
    root.geometry("600x400+100+100")

    # 파일 목록 박스
    left_frame = tk.Frame(root, width=200)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

    listbox = tk.Listbox(left_frame, width=30, height=15, font=("Arial", 12), selectmode=SINGLE)
    listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    var1 = tk.IntVar()

    selected_items = listbox.curselection()

    # 페이지 선택 라벨 프레임
    labelframe = tk.LabelFrame(left_frame, text="페이지 선택")
    labelframe.pack(side=tk.BOTTOM, fill=tk.X, expand=True)

    radio1 = tk.Radiobutton(labelframe, text="ALL", value=1, variable=var1, command= app.radio_button_click)
    radio2 = tk.Radiobutton(labelframe, text="Part", value=2, variable=var1, command=app.radio_button_click)
    radio3 = tk.Radiobutton(labelframe, text="WorkSheet", value=3, variable=var1, command=app.radio_button_click)

    radio1.pack(side=tk.LEFT, padx=5)
    var1.set(1)

    combo = Combobox(labelframe, state='normal', width=15)

    input_text = tk.Entry(labelframe, width=10, state='normal')
    label2 = tk.Label(labelframe, text='예) 1,3,5-7')




    # 우측 프레임
    right_frame = tk.Frame(root, width=300)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 버튼
    btn_frame = tk.Frame(right_frame)
    btn_frame.pack(side=tk.TOP, pady=10)
    btn_load = tk.Button(btn_frame, text='파일 로드', width=9, font=("Arial", 12), command=app.btn_load_click)
    btn_delete = tk.Button(btn_frame, text='파일 삭제', width=9, font=("Arial", 12), command=app.btn_delete_click)
    btn_merge = tk.Button(right_frame, text='파일 병합', width=12, font=("Arial", 12), command=app.btn_merge_click)
    btn_load.pack(side=tk.LEFT, padx=5)
    btn_delete.pack(side=tk.LEFT, padx=5)
    btn_merge.pack(side=tk.BOTTOM, pady=10)

    # 입력 폼
    form_frame = tk.Frame(right_frame)
    form_frame.pack(side=tk.TOP, pady=10)
    lbl_title = tk.Label(form_frame, text="제목:", font=("Arial", 12))
    entry_title = tk.Entry(form_frame, width=30, font=("Arial", 12))
    lbl_department = tk.Label(form_frame, text="부서명:", font=("Arial", 12))
    entry_department = tk.Entry(form_frame, width=30, font=("Arial", 12))
    lbl_responsible = tk.Label(form_frame, text="담당자:", font=("Arial", 12))
    entry_responsible = tk.Entry(form_frame, width=30, font=("Arial", 12))
    lbl_title.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    entry_title.grid(row=0, column=1, padx=5, pady=5)
    lbl_department.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    entry_department.grid(row=1, column=1, padx=5, pady=5)
    lbl_responsible.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    entry_responsible.grid(row=2, column=1, padx=5, pady=5)

    # bind
    listbox.bind('<<ListboxSelect>>', lambda event: app.toggle_widgets())
    input_text.bind("<Return>", app.order_input)

    root.mainloop()