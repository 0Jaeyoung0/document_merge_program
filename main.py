import src.file_io.file_io as file_io
import src.convert_document.convert_to_pdf as convert_to_pdf
import src.handle_document.handle_docx as handle_docx
import src.handle_document.handle_pdf as handle_pdf
import src.merge_document.merge_pdf as merge_pdf
import tempfile
import os

if __name__ == '__main__':
    # 파일 목록 가져오기
    file_selector = file_io.FileIO()
    input_files = file_selector.open_files()

    to_pdf_converter = convert_to_pdf.ToPdfConverter()
    file_names_without_ext = []
    converted_files = []

    with tempfile.TemporaryDirectory() as temp_dir:
        # 파일 일괄 변환
        for index, input_file in enumerate(input_files, start=1):
            # 임시 파일 경로 생성
            file_name = os.path.basename(input_file)
            file_name_without_ext = os.path.splitext(file_name)[0]  # 확장명을 제외한 파일명 가져오기
            converted_file = os.path.join(temp_dir, f'{file_name_without_ext}.pdf')

            to_pdf_converter.convert_to_pdf(input_file=input_file, output_file=converted_file)

            file_names_without_ext.append(file_name_without_ext)
            converted_files.append(converted_file)

        # 각 파일 당 페이지 수 저장
        pdf_handler = handle_pdf.PdfHandler()
        files_page_num = []

        for converted_file in converted_files:
            page_num = pdf_handler.extract_page_num(converted_file)
            files_page_num.append(page_num)

        # 입력 받기
        title = "제목"
        dept_name = "부서명"
        person_name = "홍길동"

        cover_page_docx_path = os.path.join(temp_dir, "cover_page.docx")
        index_page_docx_path = os.path.join(temp_dir, "index_page.docx")

        # 커버, 목차 페이지 생성
        word_handler = handle_docx.WordHandler()
        word_handler.create_cover_page(title, dept_name, person_name, cover_page_docx_path)
        word_handler.create_index_page(file_names_without_ext, files_page_num, index_page_docx_path)

        cover_page_pdf_path = os.path.join(temp_dir, "cover_page.pdf")
        index_page_pdf_path = os.path.join(temp_dir, "index_page.pdf")

        # 커버, 목차 페이지 변환
        to_pdf_converter.convert_to_pdf(cover_page_docx_path, cover_page_pdf_path)
        to_pdf_converter.convert_to_pdf(index_page_docx_path, index_page_pdf_path)

        # 커버, 목차 페이지 페이지 수 저장
        cover_page_num = pdf_handler.extract_page_num(cover_page_pdf_path)
        index_page_num = pdf_handler.extract_page_num(index_page_pdf_path)

        # 커버, 목차 페이지를 배열 맨 앞으로 이동
        converted_files.insert(0, index_page_pdf_path)
        converted_files.insert(0, cover_page_pdf_path)

        # 커버, 목차 페이지 수를 배열 맨 앞으로 이동
        files_page_num.insert(0, index_page_num)
        files_page_num.insert(0, cover_page_num)

        merged_file = os.path.join(temp_dir, "merged.pdf")

        # 파일 일괄 병합
        pdf_merger = merge_pdf.PdfMerger()
        pdf_merger.merge_pdf(input_files=converted_files, output_file=merged_file)

        page_file = os.path.join(temp_dir, "page.pdf")

        # 페이지 번호 추가
        pdf_handler.insert_page_number(merged_file, page_file, 1)

        output_file = "output.pdf"

        # 북마크 추가
        file_names_without_ext.insert(0, "index_page")
        file_names_without_ext.insert(0, "cover_page")
        pdf_handler.add_bookmark(page_file, file_names_without_ext, files_page_num, output_file)
