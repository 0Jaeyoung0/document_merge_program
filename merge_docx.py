import os
import random
import xml.etree.ElementTree as ET

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.styles import CT_Style

NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# 문서에서 모든 하이퍼링크를 제거합니다. 이는 부동 이미지가 있는 문서에서 오염되지 않은 docx 파일을 생성할 수 있게 합니다.
def handle_numbers(merged_doc, sub):
    # merged_doc의 numbering.xml 파일 가져오기
    try:
        merged_doc_numbering_part = merged_doc.part.numbering_part.numbering_definitions._numbering
    except:
        return
    # merged_doc의 abstractNum 요소 목록 가져오기
    merged_doc_abstract_list = merged_doc_numbering_part.xpath('//w:abstractNum')
    # 템플릿의 num 요소 목록 가져오기
    merged_doc_num_list = merged_doc_numbering_part.xpath('//w:num')

    # sub에 대해 동일한 작업 수행
    try:
        sub_numbering_part = sub.part.numbering_part.numbering_definitions._numbering
    except:
        return
    sub_abstract_list = sub_numbering_part.xpath('//w:abstractNum')
    sub_num_list = sub_numbering_part.xpath('//w:num')

    # merged_doc의 가장 높은 abstractNumId 찾기
    merged_doc_highest_abstract_id = 0
    for elem in merged_doc_abstract_list:
        abstract_id = int(elem.get(NS + 'abstractNumId'))
        if abstract_id > merged_doc_highest_abstract_id:
            merged_doc_highest_abstract_id = abstract_id

    # merged_doc의 가장 높은 abstractNumId 증가시키기. 이를 통해 sub의 abstractNum 요소가 충돌하는 ID를 가지지 않도록 함
    merged_doc_highest_abstract_id += 1

    # sub의 numbering part의 각 abstractNum 요소 반복
    for elem in sub_abstract_list:
        # 현재 요소의 abstractNumId 가져오기
        abstract_id = int(elem.get(NS + 'abstractNumId'))
        # 현재 abstractNumId를 merged_doc의 가장 높은 abstractNumId 값으로 증가시킴. 이는 각 abstractNum 요소가 충돌하는 ID를 가지지 않도록 함
        new_id = abstract_id + merged_doc_highest_abstract_id
        # abstractNum 요소의 새로운 ID 설정
        elem.set(NS + 'abstractNumId', str(new_id))
        # abstractNum 요소를 merged_doc의 numbering part에 추가
        merged_doc_numbering_part.append(elem)

    # merged_doc의 가장 높은 numId 찾기
    merged_doc_highest_num_id = 0
    for elem in merged_doc_num_list:
        num_id = int(elem.get(NS + 'numId'))
        if num_id > merged_doc_highest_num_id:
            merged_doc_highest_num_id = num_id

    # sub의 numbering part의 각 num 요소 반복
    for elem in sub_num_list:
        # num 요소에서 자식 (특히 abstractNumId) 가져오기
        children = elem.getchildren()
        for child in children:
            # num 요소에서 abstractNumId 참조 값 가져오기
            if child.tag == NS + 'abstractNumId':

                abstract_id_val = int(child.get(NS + 'val'))
                if abstract_id_val is not None:
                    # abstractNumId 참조 값에 merged_doc의 가장 높은 abstractNum id 값과 동일한 양만큼 증가시킴.
                    # 이것은 sub의 abstractNum의 id 값을 이전에 증가시켜 sub의 numbering 참조를 유지하기 위한 것임
                    new_abstract_id = abstract_id_val + merged_doc_highest_abstract_id
                    child.set(NS + 'val', str(new_abstract_id))

        # 현재 요소의 numId 가져오기
        num_id = int(elem.get(NS + 'numId'))

        # 현재 numId를 merged_doc의 가장 높은 numId 값과 동일한 양만큼 증가시킴.
        # 이렇게 함으로써 sub의 문서 부분에서 발생한 오프셋을 보정할 필요가 있음
        new_id = num_id + merged_doc_highest_num_id

        # 현재 num 요소의 새로운 ID 설정
        elem.set(NS + 'numId', str(new_id))
        # num 요소를 merged_doc의 numbering part에 추가
        merged_doc_numbering_part.append(elem)

    # sub의 document.xml 파일에서 각 numId 요소 가져오기
    sub_doc_num_id_list = sub.part.element.xpath('//w:numId')

    # sub의 문서 부분의 각 numId 요소 반복
    for elem in sub_doc_num_id_list:
        # 현재 요소의 ID 값 가져오기
        num_id = int(elem.get(NS + 'val'))
        # 현재 numId를 merged_doc의 가장 높은 numId 값과 동일한 양만큼 증가시킴.
        # 이전에 sub의 num 요소의 ID를 증가시켜 생성된 오프셋과 일치하도록 함
        new_id = num_id + merged_doc_highest_num_id
        # 현재 numId 요소의 새로운 ID 값 설정
        elem.set(NS + 'val', str(new_id))

    # 모든 abstractNum 요소는 모든 num 요소보다 앞에 나와야 함
    # 통합된 merged_doc의 numbering.xml 파일 가져오기
    final_numbering_part = merged_doc.part.numbering_part.numbering_definitions._numbering
    # 통합된 merged_doc의 abstractNum 요소 목록 가져오기
    final_abstract_list = merged_doc_numbering_part.xpath('//w:abstractNum')
    # 통합된 템플릿의 num 요소 목록 가져오기
    final_num_list = merged_doc_numbering_part.xpath('//w:num')

    # 모든 abstract 요소를 저장할 목록 유지
    abstract_list = []
    # 통합된 numbering part에서 각 abstract 요소 제거
    for elem in final_abstract_list:
        abstract_list.append(elem)
        parent = elem.getparent()
        parent.remove(elem)

    # 모든 num 요소를 저장할 목록 유지
    num_list = []
    # 통합된 numbering part에서 각 num 요소 제거
    for elem in final_num_list:
        num_list.append(elem)
        parent = elem.getparent()
        parent.remove(elem)

    # 각 abstract 요소 다시 추가
    for elem in abstract_list:
        final_numbering_part.append(elem)
    # 각 num 요소 다시 추가
    for elem in num_list:
        final_numbering_part.append(elem)


# sub_doc 문서의 모든 인라인 이미지를 merged_doc에 추가하는 함수
def handle_inlines(merged_doc, sub_doc):
    x_path = '//a:blip'
    # sub_doc 문서의 모든 a:blip 엘리먼트를 blip_list 변수에 저장
    blip_list = sub_doc.element.xpath(x_path)

    # shapes 변수에는 관심 있는 미디어를 가진 모든 InlineShape 객체가 들어 있음
    shapes = sub_doc.inline_shapes

    # 각 이미지를 고유한 파일 경로로 저장
    for i in range(len(shapes)):

        # rId를 가져오기
        shape = shapes[i]
        if shape._inline.graphic.graphicData.pic is not None:
            rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        else:
            # 해당 shape를 제거
            drawing_element = shape._inline.getparent()
            drawing_parent = drawing_element.getparent()
            drawing_parent.remove(drawing_element)
            continue

        # ImagePart 객체
        image_part = sub_doc.part.related_parts[rId]

        # Image 객체
        actual_image = image_part.image

        # 이미지의 실제 바이너리 데이터
        image_blob = actual_image.blob

        # 고유한 파일 경로 문자열 생성
        random_id = random.randint(10000, 100000)
        image_path = 'image' + str(random_id) + '.' + actual_image.ext

        # 이미지의 바이너리 데이터를 이미지 파일로 저장
        image_file = open(image_path, "wb")
        image_file.write(image_blob)
        image_file.close()

        # merged_doc 파일에 sub_doc_doc의 미디어에 대한 새로운 Relationship 생성
        new_rId, img = merged_doc.part.get_or_add_image(image_path)
        # new_rId는 추가된 이미지의 새로운 rId를 가리킴
        # img는 Image 객체를 가리킴(사용 안 함)

        blip_element = blip_list[i]
        # blip_list의 각 요소에 대해 새 rId 값을 설정
        blip_element.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rId)

        # 이미지 파일 제거
        os.remove(image_path)


# sub_doc 문서의 모든 스타일을 merged_doc의 'styles.xml' 파일에 추가하는 함수
# 이미 존재하는 스타일은 덮어써짐
def handle_styles(merged_doc, sub_doc):
    # 템플릿 파일의 styles 엘리먼트를 저장하는 변수
    merged_doc_styles = merged_doc.styles.element

    # sub_doc 문서의 styles 엘리먼트를 저장하는 변수
    sub_doc_styles = sub_doc.styles.element

    # sub_doc 문서의 각 스타일을 확인하여 CT_Style 객체인 경우 'styles.xml' 파일에 추가
    for s in sub_doc_styles:
        if isinstance(s, CT_Style):
            merged_doc_styles.append(s)

    for merged_section, sub_section in zip(merged_doc.sections, sub_doc.sections):
        # 상하좌우 여백 중 작은 값을 merged_doc에 설정
        merged_section.top_margin = min(merged_section.top_margin, sub_section.top_margin)
        merged_section.bottom_margin = min(merged_section.bottom_margin, sub_section.bottom_margin)
        merged_section.left_margin = min(merged_section.left_margin, sub_section.left_margin)
        merged_section.right_margin = min(merged_section.right_margin, sub_section.right_margin)


def add_page_break(doc):
    # <w:p> element 생성
    paragraph_element = OxmlElement('w:p')

    # <w:r> element 생성
    run_element = OxmlElement('w:r')

    # <w:br w:type="page"/> element 생성
    br_element = OxmlElement('w:br')
    br_element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'page')

    # <w:r>에 <w:br> append
    run_element.append(br_element)

    # <w:p>에 <w:r> append
    paragraph_element.append(run_element)

    # body에 <w:p> append
    doc.element.body.append(paragraph_element)


def check_page_break(xml_text):
    root = ET.fromstring(xml_text)

    # 모든 하위 태그 중에서 (<w:br w:type="page"> 또는) <Renderpagebreak>를 찾은 후 있을 경우 리턴
    for child in root.iter():
        #if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br' and child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == 'page':
        #    return True
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lastRenderedPageBreak':
            return True

    return False



def merge_docx(file_list, file_name):
    merged_doc = Document()
    selected_pages = []

    for file in file_list:
        sub_doc = Document(file)
        page_number = 1
        user_input = input("저장할 페이지 번호를 입력하세요 (여러 페이지일 경우 쉼표로 구분, 전체 페이지를 저장할 경우 엔터키를 누르세요): ")
        if user_input:
            selected_pages = [int(num) for num in user_input.split(',')]
            if len(selected_pages) != len(set(selected_pages)):
                print("중복된 페이지 번호가 있습니다.")
        # 원하는 페이지가 없을 경우
        else:
            print("페이지를 저장하지 않았습니다.")



        print(f"{file} 저장 중..")
        # 인라인 이미지 처리
        handle_inlines(merged_doc, sub_doc)
        # 스타일 처리
        handle_styles(merged_doc, sub_doc)
        # 넘버링 처리
        handle_numbers(merged_doc, sub_doc)

        # sub_doc_doc의 body 엘리먼트를 merged_doc 파일의 body에 추가
        for element in sub_doc.element.body:
            if check_page_break(element.xml):
                page_number = page_number + 1
            if page_number in selected_pages:
                merged_doc.element.body.append(element)
            elif not selected_pages:
                merged_doc.element.body.append(element)

        # 수동으로 페이지 구분선 추가
        add_page_break(merged_doc)

    # 문서 저장
    merged_doc.save(file_name)

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


# file_load(file_list)
file_list = [r'C:\Users\서예은\Desktop\문서 통합\Python\code\신청서.docx',r'C:\Users\서예은\Desktop\문서 통합\Python\code\설문지.docx']
merge_docx(file_list, 'output.docx')

# file_list = ['docx_sample/test1.docx', 'docx_sample/test2.docx']
# merge_docx(file_list, 'docx_sample/output.docx')
