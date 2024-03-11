from merge_document.merge_docx import merge_docx
from merge_document.merge_xlsx import merge_xlsx_multi_sheet, merge_xlsx_single_sheet, merge_xlsx_single_table
import os

if __name__ == "__main__":
    docx_list = ['docx_sample/test1.docx',
                 'docx_sample/test2.docx']
    # merge_docx(docx_list, 'docx_sample/output.docx')

    sheet_list = [
        ("xlsx_sample/test2.xlsx", "참여학사조직"),
        ("xlsx_sample/test1.xlsx", "창원대")
    ]
    # merge_xlsx_single_sheet(sheet_list, 'xlsx_sample/output_single_sheet.xlsx')
    # merge_xlsx_multi_sheet(sheet_list, 'xlsx_sample/output_multi_sheet.xlsx')
    merge_xlsx_single_sheet(sheet_list, 'xlsx_sample/output_multi_sheet.xlsx')


def file_load(file_list):
    # 사용자가 원하는 파일 불러오기
    path = "./"
    dirPath = os.listdir(path)
    print(dirPath)

    while True:
        file_name = input("불러올 파일명(.docx)을 입력하세요[exit 입력 시 나감]: ")
        if not file_name.endswith(".docx"):
            print("올바른 형식이 아닙니다.")
        if file_name == "exit":
            break
        else:
            file_list.append(file_name)

    # 불러온 파일 확인
    for file in file_list:
        print(file)