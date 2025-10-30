#!/usr/bin/env python3
"""
Script xử lý file docx:
1. Xóa trang đầu nếu chỉ có "thẻ 1"
2. Xóa nội dung GIỮA các tag START0 và END (cả cùng và khác paragraph)
3. Xóa các table row có ROW0
4. Xóa tất cả các tag còn lại

Cách sử dụng:
    python process_docx.py input.docx output.docx
"""
import sys
import os
import re
import zipfile
import tempfile
import shutil
from defusedxml import minidom

def unpack_docx(docx_path, extract_dir):
    """Giải nén file docx"""
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def pack_docx(source_dir, output_path):
    """Nén lại thành file docx"""
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_dir)
                docx.write(file_path, arcname)

def get_all_text_from_element(element):
    """Lấy tất cả text từ một element"""
    text_elems = element.getElementsByTagName('w:t')
    return ''.join([t.firstChild.nodeValue if t.firstChild else '' for t in text_elems])

def set_text_in_element(element, new_text):
    """
    Cập nhật text trong element
    Xóa tất cả w:t elements cũ và tạo một w:t element mới với text mới
    """
    # Xóa tất cả các w:t elements cũ
    text_elems = list(element.getElementsByTagName('w:t'))
    for text_elem in text_elems:
        parent = text_elem.parentNode
        if parent:
            parent.removeChild(text_elem)
    
    # Nếu text mới rỗng, không cần thêm gì
    if not new_text:
        return
    
    # Tìm hoặc tạo w:r (run) element
    runs = element.getElementsByTagName('w:r')
    if runs:
        run = runs[0]
    else:
        # Tạo w:r mới nếu chưa có
        run = element.ownerDocument.createElement('w:r')
        element.appendChild(run)
    
    # Tạo w:t element mới với text
    text_elem = element.ownerDocument.createElement('w:t')
    text_elem.setAttribute('xml:space', 'preserve')
    text_node = element.ownerDocument.createTextNode(new_text)
    text_elem.appendChild(text_node)
    run.appendChild(text_elem)

def remove_content_between_tags_same_element(element, start_tag_pattern, end_tag_pattern):
    """
    Xóa tất cả các cặp START-END tag và nội dung giữa chúng trong cùng một element
    Trả về True nếu có xóa, False nếu không
    """
    text = get_all_text_from_element(element)
    original_text = text
    
    # Lặp lại cho đến khi không còn cặp START-END nào
    max_iterations = 100  # Tránh vòng lặp vô hạn
    iteration = 0
    
    while iteration < max_iterations:
        # Tìm cặp START-END đầu tiên
        # Pattern: START...END (non-greedy)
        pattern = start_tag_pattern + r'.*?' + end_tag_pattern
        match = re.search(pattern, text, re.DOTALL)
        
        if not match:
            # Không còn cặp START-END nào
            break
        
        # Xóa cặp START-END này (cả tag và content giữa chúng)
        text = text[:match.start()] + text[match.end():]
        iteration += 1
    
    # Cập nhật text vào element nếu có thay đổi
    if text != original_text:
        if text.strip():  # Nếu còn text
            set_text_in_element(element, text)
        else:  # Nếu text rỗng, có thể giữ element rỗng hoặc đánh dấu để xóa
            set_text_in_element(element, '')
        return True
    
    return False

def process_removal_between_tags(body, start_tag_type, end_tag_type, label):
    """
    Xử lý xóa nội dung giữa các tag START và END
    Xử lý cả trường hợp cùng element và khác element
    """
    # Tạo regex patterns cho START và END tags
    if start_tag_type == 'BLOCK_START':
        start_pattern = rf'\[\[BLOCK_START{label}\]\]'
    elif start_tag_type == 'SECTION_START':
        start_pattern = rf'\[\[SECTION_START{label}\]\]'
    else:
        return 0
    
    if end_tag_type == 'BLOCK_END':
        end_pattern = r'\[\[BLOCK_END\]\]'
    elif end_tag_type == 'SECTION_END':
        end_pattern = r'\[\[SECTION_END\]\]'
    else:
        return 0
    
    all_elements = []
    for child in body.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            if child.tagName in ['w:p', 'w:tbl']:
                all_elements.append(child)
    
    elements_to_remove = []
    i = 0
    
    while i < len(all_elements):
        element = all_elements[i]
        text = get_all_text_from_element(element)
        
        # Trước tiên, xử lý tất cả các cặp START-END trong cùng element
        if re.search(start_pattern, text) and re.search(end_pattern, text):
            print(f"  Xử lý cặp {start_tag_type}{label}-{end_tag_type} trong cùng element")
            removed = remove_content_between_tags_same_element(element, start_pattern, end_pattern)
            if removed:
                # Sau khi xóa, check lại text xem element còn gì không
                remaining_text = get_all_text_from_element(element)
                if not remaining_text.strip():
                    elements_to_remove.append(element)
                # Không tăng i, check lại element này xem còn START tag nào không
                continue
        
        # Kiểm tra xem có START tag (nhưng không có END trong cùng element)
        if re.search(start_pattern, text):
            print(f"  Tìm thấy {start_tag_type}{label}, tìm {end_tag_type} ở các element sau")
            
            # Xóa từ START đến cuối element hiện tại
            match = re.search(start_pattern, text)
            if match:
                new_text = text[:match.start()]
                if new_text.strip():
                    set_text_in_element(element, new_text)
                else:
                    elements_to_remove.append(element)
            
            # Tìm element chứa END tag
            j = i + 1
            found_end = False
            
            while j < len(all_elements):
                next_element = all_elements[j]
                next_text = get_all_text_from_element(next_element)
                
                # Kiểm tra END tag
                end_match = re.search(end_pattern, next_text)
                if end_match:
                    # Tìm thấy END tag
                    print(f"    Tìm thấy {end_tag_type} tại element {j}")
                    
                    # Xóa từ đầu element đến hết END tag
                    new_text = next_text[end_match.end():]
                    if new_text.strip():
                        set_text_in_element(next_element, new_text)
                    else:
                        elements_to_remove.append(next_element)
                    
                    # Đánh dấu các element ở giữa để xóa (từ i+1 đến j-1)
                    for k in range(i + 1, j):
                        if all_elements[k] not in elements_to_remove:
                            elements_to_remove.append(all_elements[k])
                    
                    found_end = True
                    i = j + 1
                    break
                else:
                    # Element này nằm giữa START và END, đánh dấu để xóa
                    if next_element not in elements_to_remove:
                        elements_to_remove.append(next_element)
                    j += 1
            
            if not found_end:
                print(f"  CẢNH BÁO: Không tìm thấy {end_tag_type} tương ứng với {start_tag_type}{label}")
                i += 1
            
            continue
        
        i += 1
    
    # Xóa các elements đã đánh dấu
    for element in elements_to_remove:
        if element.parentNode:
            element.parentNode.removeChild(element)
    
    return len(elements_to_remove)

def remove_rows_with_tag(body, label):
    """Xóa các table row có ROW tag với label cụ thể"""
    tables = body.getElementsByTagName('w:tbl')
    rows_removed = 0
    
    for table in tables:
        rows = list(table.getElementsByTagName('w:tr'))
        for row in rows:
            row_text = get_all_text_from_element(row)
            
            # Tìm tag ROW với label
            if re.search(rf'\[\[ROW{label}\]\]', row_text):
                row.parentNode.removeChild(row)
                rows_removed += 1
    
    return rows_removed

def remove_all_remaining_tags(body):
    """Xóa tất cả các tag còn lại trong document"""
    patterns = [
        r'\[\[BLOCK_START\d+\]\]',
        r'\[\[BLOCK_END\]\]',
        r'\[\[SECTION_START\d+\]\]',
        r'\[\[SECTION_END\]\]',
        r'\[\[ROW\d+\]\]',
    ]
    
    # Xử lý tất cả các paragraphs và tables
    all_elements = []
    for child in body.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            if child.tagName in ['w:p', 'w:tbl']:
                all_elements.append(child)
    
    tags_removed = 0
    
    for element in all_elements:
        text = get_all_text_from_element(element)
        original_text = text
        
        # Xóa tất cả các patterns
        for pattern in patterns:
            text = re.sub(pattern, '', text)
        
        if text != original_text:
            tags_removed += 1
            if text.strip():
                set_text_in_element(element, text)
            else:
                # Nếu element chỉ có tags và không còn gì, có thể giữ nguyên hoặc xóa
                # Ở đây ta giữ element rỗng
                set_text_in_element(element, '')
    
    return tags_removed

def get_first_page_elements(body):
    """Lấy tất cả elements của trang đầu tiên"""
    first_page_elements = []
    
    for child in body.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
            
        if child.tagName not in ['w:p', 'w:tbl']:
            continue
        
        has_page_break = False
        if child.tagName == 'w:p':
            runs = child.getElementsByTagName('w:r')
            for run in runs:
                breaks = run.getElementsByTagName('w:br')
                for br in breaks:
                    br_type = br.getAttribute('w:type')
                    if br_type == 'page':
                        has_page_break = True
                        break
                if has_page_break:
                    break
            
            pPr = child.getElementsByTagName('w:pPr')
            if pPr:
                sectPr = pPr[0].getElementsByTagName('w:sectPr')
                if sectPr:
                    has_page_break = True
        
        first_page_elements.append(child)
        
        if has_page_break:
            break
    
    return first_page_elements

def remove_first_page_if_the1(body):
    """Xóa trang đầu tiên nếu chỉ có nội dung là 'thẻ 1'"""
    print("Bước 0: Kiểm tra và xóa trang đầu nếu chỉ có 'thẻ 1'")
    
    first_page_elements = get_first_page_elements(body)
    
    if not first_page_elements:
        print("  Không tìm thấy elements ở trang đầu")
        return
    
    all_text = ''
    for element in first_page_elements:
        all_text += get_all_text_from_element(element)
    
    normalized_text = all_text.strip().lower()
    
    print(f"  Nội dung trang đầu: '{normalized_text}'")
    
    if normalized_text == 'thẻ 1':
        print("  ✓ Phát hiện trang đầu chỉ có 'thẻ 1', đang xóa...")
        for element in first_page_elements:
            body.removeChild(element)
        print(f"  Đã xóa {len(first_page_elements)} elements từ trang đầu")
    else:
        print("  Trang đầu không phải chỉ có 'thẻ 1', giữ nguyên")

def process_document_xml(xml_path):
    """Xử lý file document.xml"""
    print("Bắt đầu xử lý document.xml")
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    dom = minidom.parseString(content)
    body = dom.getElementsByTagName('w:body')[0]

    # Bước 0: Xóa trang đầu nếu chỉ có "thẻ 1"
    remove_first_page_if_the1(body)

    # Bước 1: Xử lý BLOCK_START0 và BLOCK_END
    print("\nBước 1: Xử lý xóa nội dung giữa BLOCK_START0 và BLOCK_END")
    elements_removed = process_removal_between_tags(body, 'BLOCK_START', 'BLOCK_END', '0')
    print(f"  Đã xóa {elements_removed} elements giữa các BLOCK tags")

    # Bước 2: Xử lý SECTION_START0 và SECTION_END
    print("\nBước 2: Xử lý xóa nội dung giữa SECTION_START0 và SECTION_END")
    elements_removed = process_removal_between_tags(body, 'SECTION_START', 'SECTION_END', '0')
    print(f"  Đã xóa {elements_removed} elements giữa các SECTION tags")

    # Bước 3: Xóa các table row có ROW0
    print("\nBước 3: Xóa các table row có ROW0")
    rows_removed = remove_rows_with_tag(body, '0')
    print(f"  Đã xóa {rows_removed} rows có ROW0")

    # Bước 4: Xóa tất cả các tag còn lại
    print("\nBước 4: Xóa tất cả các tag còn lại")
    tags_removed = remove_all_remaining_tags(body)
    print(f"  Đã xóa tags từ {tags_removed} elements")

    # Bước 5: Lưu document.xml
    print("\nBước 5: Lưu document.xml")
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())
    print("Hoàn thành xử lý document.xml")

def main():
    if len(sys.argv) != 3:
        print("Cách sử dụng: python process_docx.py <input.docx> <output.docx>")
        print("Ví dụ: python process_docx.py TEST.docx TEST_processed.docx")
        sys.exit(1)
    
    input_docx = sys.argv[1]
    output_docx = sys.argv[2]
    
    if not os.path.exists(input_docx):
        print(f"Lỗi: Không tìm thấy file {input_docx}")
        sys.exit(1)
    
    print(f"Đang xử lý file: {input_docx}")
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        print("Đang giải nén file docx...")
        unpack_docx(input_docx, temp_dir)
        
        doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        if not os.path.exists(doc_xml_path):
            print("Lỗi: Không tìm thấy word/document.xml trong file docx")
            sys.exit(1)
        
        process_document_xml(doc_xml_path)
        
        print(f"\nĐang tạo file output: {output_docx}")
        pack_docx(temp_dir, output_docx)
        
        print(f"\n✅ Hoàn thành! File đã được lưu tại: {output_docx}")
        
    finally:
        shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()