#!/usr/bin/env python3
"""
Script xử lý file docx:
1. Xóa nội dung trong các phần có tag "0" (BLOCK_START0, SECTION_START0, ROW0)
   - BAO GỒM CẢ CÁC BẢNG (TABLES) NẰM TRONG CÁC KHỐI NÀY
2. Xóa các tag còn lại nhưng giữ nội dung và format

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

def has_tag_pattern(text, pattern_type, label=None):
    """Kiểm tra xem text có chứa tag pattern không"""
    if pattern_type == 'BLOCK_START':
        if label:
            return bool(re.search(rf'\[\[BLOCK_START{label}\]\]', text))
        return bool(re.search(r'\[\[BLOCK_START\d+\]\]', text))
    elif pattern_type == 'BLOCK_END':
        return '[[BLOCK_END]]' in text
    elif pattern_type == 'SECTION_START':
        if label:
            return bool(re.search(rf'\[\[SECTION_START{label}\]\]', text))
        return bool(re.search(r'\[\[SECTION_START\d+\]\]', text))
    elif pattern_type == 'SECTION_END':
        return '[[SECTION_END]]' in text
    elif pattern_type == 'ROW':
        if label:
            return bool(re.search(rf'\[\[ROW{label}\]\]', text))
        return bool(re.search(r'\[\[ROW\d+\]\]', text))
    return False

def remove_all_tags(text):
    """Xóa tất cả các tag khỏi text"""
    patterns = [
        r'\[\[BLOCK_START\d+\]\]',
        r'\[\[BLOCK_END\]\]',
        r'\[\[SECTION_START\d+\]\]',
        r'\[\[SECTION_END\]\]',
        r'\[\[ROW\d+\]\]',
    ]
    
    result = text
    for pattern in patterns:
        result = re.sub(pattern, '', result)
    return result

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
    """Lấy tất cả text từ một element (paragraph hoặc table)"""
    text_elems = element.getElementsByTagName('w:t')
    return ''.join([t.firstChild.nodeValue if t.firstChild else '' for t in text_elems])

def mark_elements_for_removal(body, start_pattern, end_pattern, label):
    """
    Đánh dấu tất cả elements (paragraphs và tables) cần xóa
    Trả về set các element IDs cần xóa
    """
    # Lấy tất cả child elements của body (bao gồm w:p và w:tbl)
    all_elements = []
    for child in body.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            if child.tagName in ['w:p', 'w:tbl']:
                all_elements.append(child)
    
    elements_to_remove = set()
    in_removal = False
    depth = 0
    
    for element in all_elements:
        text = get_all_text_from_element(element)
        
        # Kiểm tra có start tag với label cụ thể không
        if not in_removal and has_tag_pattern(text, start_pattern, label):
            in_removal = True
            depth = 1
            elements_to_remove.add(id(element))
            continue
        
        if in_removal:
            elements_to_remove.add(id(element))
            
            # Đếm nested tags
            if has_tag_pattern(text, start_pattern):
                depth += 1
            
            if has_tag_pattern(text, end_pattern):
                depth -= 1
            
            if depth == 0:
                in_removal = False
    
    return elements_to_remove

def get_first_page_elements(body):
    """
    Lấy tất cả elements của trang đầu tiên (trước page break đầu tiên)
    """
    first_page_elements = []
    
    for child in body.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
            
        if child.tagName not in ['w:p', 'w:tbl']:
            continue
        
        # Kiểm tra xem element này có page break không
        has_page_break = False
        if child.tagName == 'w:p':
            # Tìm page break trong paragraph
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
            
            # Kiểm tra section break (thường ở cuối paragraph)
            pPr = child.getElementsByTagName('w:pPr')
            if pPr:
                sectPr = pPr[0].getElementsByTagName('w:sectPr')
                if sectPr:
                    has_page_break = True
        
        first_page_elements.append(child)
        
        # Nếu gặp page break, dừng lại
        if has_page_break:
            break
    
    return first_page_elements

def remove_first_page_if_the1(body):
    """
    Xóa trang đầu tiên nếu chỉ có nội dung là 'thẻ 1' (case-insensitive)
    """
    print("Bước 0: Kiểm tra và xóa trang đầu nếu chỉ có 'thẻ 1'")
    
    first_page_elements = get_first_page_elements(body)
    
    if not first_page_elements:
        print("  Không tìm thấy elements ở trang đầu")
        return
    
    # Lấy tất cả text từ trang đầu
    all_text = ''
    for element in first_page_elements:
        all_text += get_all_text_from_element(element)
    
    # Normalize text (strip whitespace và lowercase)
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
    # Đọc file XML
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    dom = minidom.parseString(content)
    body = dom.getElementsByTagName('w:body')[0]
    
    # Xóa trang đầu nếu chỉ có "thẻ 1"
    remove_first_page_if_the1(body)
    
    print("Bước 1: Đánh dấu các BLOCK_START0 cần xóa (bao gồm cả tables)")
    block_element_ids = mark_elements_for_removal(body, 'BLOCK_START', 'BLOCK_END', '0')
    print(f"  Tìm thấy {len(block_element_ids)} elements trong BLOCK_START0")
    
    print("Bước 2: Đánh dấu các SECTION_START0 cần xóa (bao gồm cả tables)")
    section_element_ids = mark_elements_for_removal(body, 'SECTION_START', 'SECTION_END', '0')
    print(f"  Tìm thấy {len(section_element_ids)} elements trong SECTION_START0")
    
    print("Bước 3: Xóa các elements đã đánh dấu (paragraphs và tables)")
    all_elements_to_remove = block_element_ids | section_element_ids
    
    # Xóa các paragraphs và tables đã được đánh dấu
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE:
            if child.tagName in ['w:p', 'w:tbl']:
                if id(child) in all_elements_to_remove:
                    body.removeChild(child)
                    print(f"    Đã xóa: {child.tagName}")
    
    print("Bước 4: Xóa các table row có ROW0 (trong các bảng còn lại)")
    tables = body.getElementsByTagName('w:tbl')
    for table in tables:
        rows = list(table.getElementsByTagName('w:tr'))
        for row in rows:
            # Lấy text từ tất cả cells trong row
            row_text = get_all_text_from_element(row)
            
            if has_tag_pattern(row_text, 'ROW', '0'):
                row.parentNode.removeChild(row)
                print(f"    Đã xóa row có ROW0")
    
    print("Bước 5: Xóa tất cả các tag còn lại")
    all_text_elems = body.getElementsByTagName('w:t')
    for text_elem in all_text_elems:
        if text_elem.firstChild and text_elem.firstChild.nodeValue:
            original_text = text_elem.firstChild.nodeValue
            cleaned_text = remove_all_tags(original_text)
            if cleaned_text != original_text:
                if cleaned_text:
                    text_elem.firstChild.nodeValue = cleaned_text
                else:
                    # Nếu text rỗng sau khi xóa tag, xóa luôn text element
                    parent = text_elem.parentNode
                    if parent:
                        parent.removeChild(text_elem)
    
    print("Bước 6: Lưu document.xml")
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())

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
        
        print(f"Đang tạo file output: {output_docx}")
        pack_docx(temp_dir, output_docx)
        
        print(f"✅ Hoàn thành! File đã được lưu tại: {output_docx}")
        
    finally:
        shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()