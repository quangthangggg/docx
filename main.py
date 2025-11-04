#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
process_docx.py

An toàn cho bảng & header:
- Chỉ thao tác trên text node (w:t); không ghi đè w:tbl, không tái tạo run.
- Xoá chuẩn các khối [[BLOCK_START0]]...[[BLOCK_END]] / [[SECTION_START0]]...[[SECTION_END]]
- Xoá chính xác hàng chứa [[ROW0]]
- Giữ ROW1: chỉ gỡ đúng chuỗi [[ROW1]] (và các tag khác)

Cách dùng:
    python process_docx.py input.docx output.docx
"""

import sys
import os
import re
import zipfile
import tempfile
import shutil
from defusedxml import minidom

# ------------------------------
# Zip helpers
# ------------------------------

def unpack_docx(docx_path, extract_dir):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def pack_docx(source_dir, output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
        for root, _, files in os.walk(source_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_dir)
                docx.write(file_path, arcname)

# ------------------------------
# XML utilities
# ------------------------------

def get_all_text_from_element(element):
    """Nối toàn bộ text từ các w:t con (để debug/log)."""
    text_elems = element.getElementsByTagName('w:t')
    return ''.join([t.firstChild.nodeValue if (t.firstChild is not None) else '' for t in text_elems])

def _iter_text_nodes_in(element):
    """Trả về danh sách w:t (text nodes) theo thứ tự xuất hiện trong element."""
    return list(element.getElementsByTagName('w:t'))

def _concat_and_spans(text_nodes):
    """
    Gộp text của danh sách w:t thành một chuỗi & ánh xạ vị trí:
    Trả về (full_text, spans) với spans[i] = (start, end) trong full_text của text_nodes[i].
    """
    full = []
    spans = []
    pos = 0
    for t in text_nodes:
        s = t.firstChild.nodeValue if (t.firstChild is not None) else ''
        start = pos
        pos += len(s)
        end = pos
        full.append(s)
        spans.append((start, end))
    return ''.join(full), spans

def _apply_kept_ranges_to_text_nodes(text_nodes, spans, kept_ranges):
    """
    kept_ranges: list các (start, end) (nửa mở) trong không gian của full_text cần GIỮ LẠI.
    Hàm sẽ cập nhật text của từng w:t cho đúng phần giao với kept_ranges.
    """
    # Hợp nhất kept_ranges đã cho (phòng trường hợp bị chồng)
    merged = []
    for s, e in sorted(kept_ranges):
        if s >= e:
            continue
        if not merged or s > merged[-1][1]:
            merged.append([s, e])
        else:
            merged[-1][1] = max(merged[-1][1], e)

    # Cập nhật từng text node theo phần giao
    for (node, (ns, ne)) in zip(text_nodes, spans):
        pieces = []
        for (ks, ke) in merged:
            # giao giữa [ns, ne) và [ks, ke)
            s = max(ns, ks)
            e = min(ne, ke)
            if s < e:
                # trích đoạn từ node
                inner_s = s - ns
                inner_e = e - ns
                txt = node.firstChild.nodeValue if (node.firstChild is not None) else ''
                pieces.append(txt[inner_s:inner_e])
        new_text = ''.join(pieces)
        if node.firstChild is None:
            node.appendChild(node.ownerDocument.createTextNode(new_text))
        else:
            node.firstChild.nodeValue = new_text

def _remove_pairs_in_same_paragraph(p, start_pat, end_pat):
    """
    Xoá mọi cặp START..END (và phần giữa) nếu chúng nằm trong CÙNG MỘT w:p.
    Chỉ chỉnh sửa w:t; không đụng run/paragraph khác.
    Trả về True nếu có thay đổi.
    """
    ts = _iter_text_nodes_in(p)
    if not ts:
        return False
    full, spans = _concat_and_spans(ts)

    # Tìm mọi cặp theo non-greedy
    pattern = re.compile(start_pat + r'.*?' + end_pat, flags=re.DOTALL)
    removed = []
    pos = 0
    while True:
        m = pattern.search(full, pos)
        if not m:
            break
        removed.append((m.start(), m.end()))
        pos = m.end()

    if not removed:
        return False

    # kept = complement của removed trong [0, len(full))
    kept = []
    cur = 0
    for (s, e) in removed:
        if cur < s:
            kept.append((cur, s))
        cur = e
    if cur < len(full):
        kept.append((cur, len(full)))

    _apply_kept_ranges_to_text_nodes(ts, spans, kept)
    return True

def _has_ancestor_tag(node, tag_names):
    """
    Trả về True nếu node có ancestor với tagName thuộc tag_names (list).
    tag_names ví dụ: ['w:tbl', 'w:tr', 'w:tc']
    """
    cur = node.parentNode
    while cur is not None:
        if getattr(cur, "tagName", None) in tag_names:
            return True
        cur = cur.parentNode
    return False


def _cut_after_start_in_paragraph(p, start_pat):
    """
    Nếu đoạn có START (không có END), cắt từ vị trí START đến hết đoạn.
    """
    ts = _iter_text_nodes_in(p)
    if not ts:
        return False
    full, spans = _concat_and_spans(ts)

    m = re.search(start_pat, full)
    if not m:
        return False

    kept = []
    if m.start() > 0:
        kept.append((0, m.start()))
    _apply_kept_ranges_to_text_nodes(ts, spans, kept)
    return True

def _cut_before_end_in_paragraph(p, end_pat):
    """
    Nếu đoạn có END (không có START), cắt từ đầu đến hết END.
    """
    ts = _iter_text_nodes_in(p)
    if not ts:
        return False
    full, spans = _concat_and_spans(ts)

    m = re.search(end_pat, full)
    if not m:
        return False

    kept = []
    if m.end() < len(full):
        kept.append((m.end(), len(full)))
    _apply_kept_ranges_to_text_nodes(ts, spans, kept)
    return True

# ------------------------------
# First page helpers
# ------------------------------

def get_first_page_elements(body):
    """Thu tất cả w:p, w:tbl của trang đầu dựa trên page/section break."""
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
                for br in run.getElementsByTagName('w:br'):
                    if br.getAttribute('w:type') == 'page':
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
    """Xoá trang đầu nếu chỉ có 'thẻ 1' (không phân biệt hoa/thường)."""
    print("Bước 0: Kiểm tra và xóa trang đầu nếu chỉ có 'thẻ 1'")
    first = get_first_page_elements(body)
    if not first:
        print("  Không tìm thấy elements ở trang đầu")
        return

    txt = ''.join(get_all_text_from_element(e) for e in first).strip().lower()
    print(f"  Nội dung trang đầu: '{txt}'")
    if txt == 'thẻ 1':
        for e in first:
            if e.parentNode:
                e.parentNode.removeChild(e)
        print(f"  ✓ Đã xóa {len(first)} elements từ trang đầu")

# ------------------------------
# Core processors
# ------------------------------

# *** BẮT ĐẦU THAY ĐỔI ***
# Hàm này là hàm mới, kết hợp logic của `remove_block_content_including_tables`
# và `process_removal_between_tags`
def remove_nodes_between_tags(body, start_tag_type, end_tag_type, label):
    """
    Xoá các node (w:p, w:tbl) nằm giữa [[START_TAG{label}]] và [[END_TAG]].
    Hàm này duyệt các childNodes (w:p, w:tbl) của body và xoá mọi thứ ở giữa,
    bao gồm cả bảng.
    Các tag start/end sẽ được xoá khỏi các node chứa chúng.
    Trả về tổng số thay đổi (nodes bị xóa + số cặp được xử lý trong cùng đoạn).
    """
    # Xác định pattern dựa trên type
    if start_tag_type == 'BLOCK_START':
        start_pat = rf'\[\[BLOCK_START{label}\]\]'
    elif start_tag_type == 'SECTION_START':
        start_pat = rf'\[\[SECTION_START{label}\]\]'
    else:
        print(f"Lỗi: Kiểu tag bắt đầu không hợp lệ: {start_tag_type}")
        return 0 # Kiểu tag không hợp lệ

    if end_tag_type == 'BLOCK_END':
        end_pat = r'\[\[BLOCK_END\]\]'
    elif end_tag_type == 'SECTION_END':
        end_pat = r'\[\[SECTION_END\]\]'
    else:
        print(f"Lỗi: Kiểu tag kết thúc không hợp lệ: {end_tag_type}")
        return 0 # Kiểu tag không hợp lệ

    nodes_to_remove = []
    pairs_handled = 0  # Đếm số cặp được xử lý trong cùng đoạn
    in_block = False

    # body.childNodes là một Live NodeList, cần copy ra list để xoá an toàn
    for node in list(body.childNodes):
        if node.nodeType != node.ELEMENT_NODE:
            continue

        # Chỉ xử lý w:p và w:tbl
        if node.tagName not in ['w:p', 'w:tbl']:
            continue

        node_text = get_all_text_from_element(node)
        start_match = re.search(start_pat, node_text)
        end_match = re.search(end_pat, node_text)

        if in_block:
            if end_match:
                in_block = False
                # Check if the node will be empty after removing the tag
                node_text_after_removal = re.sub(end_pat, '', node_text, flags=re.DOTALL)
                if not node_text_after_removal.strip():
                    nodes_to_remove.append(node)
                else:
                    # Xoá tag [[END_TAG]] khỏi node này
                    if node.tagName == 'w:p':
                        _cut_before_end_in_paragraph(node, end_pat)
            else:
                nodes_to_remove.append(node)

        elif start_match:
            # Nếu start và end trong cùng 1 node và theo đúng thứ tự
            if end_match and start_match.start() < end_match.start():
                if node.tagName == 'w:p':
                    if _remove_pairs_in_same_paragraph(node, start_pat, end_pat):
                        pairs_handled += 1
                        # Check if the paragraph is now empty and should be removed
                        node_text_after = get_all_text_from_element(node)
                        if not node_text_after.strip():
                            nodes_to_remove.append(node)
                # Với bảng (w:tbl), nếu START và END trong cùng 1 bảng thì xóa cả bảng
                elif node.tagName == 'w:tbl':
                    nodes_to_remove.append(node)
            else:
                # Bắt đầu một block mới
                in_block = True
                # Check if the node will be empty after removing the tag and content after it
                node_text_after_removal = re.sub(start_pat + r'.*$', '', node_text, flags=re.DOTALL)
                if not node_text_after_removal.strip():
                    nodes_to_remove.append(node)
                else:
                    # Chỉ xoá tag và phần sau nó
                    if node.tagName == 'w:p':
                        _cut_after_start_in_paragraph(node, start_pat)

    for node in nodes_to_remove:
        if node.parentNode:
            node.parentNode.removeChild(node)

    return len(nodes_to_remove) + pairs_handled

# *** KẾT THÚC THAY ĐỔI ***
# (Hàm `remove_block_content_including_tables` và `process_removal_between_tags` cũ đã bị xóa)


def clear_row_content_with_tag(body, label):
    """Xoá nội dung của w:tr có [[ROW{label}]], nhưng giữ lại hàng.
    Nội dung ở đây là các text nodes (w:t).
    """
    rows_cleared = 0
    tag_pattern = rf'\[\[ROW{label}\]\]'
    
    for tr in body.getElementsByTagName('w:tr'):
        # Phải kiểm tra lại parentNode vì có thể hàng đã bị xoá trong bước trước
        if not tr.parentNode:
            continue
            
        row_text = get_all_text_from_element(tr)
        if re.search(tag_pattern, row_text):
            # Tìm thấy hàng chứa tag. Xoá text của tất cả w:t con.
            text_nodes = tr.getElementsByTagName('w:t')
            for t in text_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = ''
            rows_cleared += 1
    return rows_cleared


def remove_rows_with_tag(body, label):
    """Xoá w:tr có [[ROW{label}]] (vd [[ROW0]])."""
    rows_removed = 0
    tag_pattern = rf'\[\[ROW{label}\]\]'
    
    # We need to iterate and remove carefully.
    # It's better to find all rows to be removed first, then remove them.
    rows_to_remove = []
    for tr in body.getElementsByTagName('w:tr'):
        row_text = get_all_text_from_element(tr)
        if re.search(tag_pattern, row_text):
            rows_to_remove.append(tr)

    for tr in rows_to_remove:
        if tr.parentNode:
            print(f"  - Removing a w:tr node containing [[ROW{label}]]")
            tr.parentNode.removeChild(tr)
            rows_removed += 1
            
    return rows_removed

def _replace_tags_in_text_nodes(body, patterns):
    """
    Gỡ tag bằng replace trực tiếp trong w:t để không đụng run/paragraph.
    patterns: list[str] regex
    """
    changed = 0
    for t in body.getElementsByTagName('w:t'):
        if t.firstChild and t.firstChild.nodeType == t.firstChild.TEXT_NODE:
            old = t.firstChild.nodeValue
            new = old
            for pat in patterns:
                new = re.sub(pat, '', new)
            if new != old:
                t.firstChild.nodeValue = new
                changed += 1
    return changed

def remove_all_remaining_tags(body):
    """
    Gỡ sạch các tag còn lại, bao gồm [[ROW_END]], xử lý cả trường hợp tag bị tách.
    """
    patterns = [
        r'\[\[BLOCK_START\d+\]\]',
        r'\[\[BLOCK_END\]\]',
        r'\[\[SECTION_START\d+\]\]',
        r'\[\[SECTION_END\]\]',
        r'\[\[ROW\d+\]\]',   # gỡ mọi [[ROWx]]
        r'\[\[ROW_END\]\]',   # gỡ [[ROW_END]]
    ]
    
    changed_elements_count = 0
    
    # Iterate through all paragraphs and table cells, as these are common containers for w:t
    # Also iterate through w:tr for completeness, though w:tc is usually sufficient for table text
    for container_tag in ['w:p', 'w:tc', 'w:tr']:
        for container_elem in body.getElementsByTagName(container_tag):
            text_nodes = _iter_text_nodes_in(container_elem)
            if not text_nodes:
                continue
            
            full_text, spans = _concat_and_spans(text_nodes)
            
            kept_ranges = [(0, len(full_text))] # Initially, keep everything
            
            for pat in patterns:
                new_kept_ranges = []
                for k_start, k_end in kept_ranges:
                    current_segment = full_text[k_start:k_end]
                    
                    # Find all matches of the pattern within the current segment
                    matches = list(re.finditer(pat, current_segment, flags=re.DOTALL))
                    
                    if not matches:
                        new_kept_ranges.append((k_start, k_end))
                        continue
                    
                    current_pos = 0
                    for m in matches:
                        # Add the part before the match
                        if m.start() > current_pos:
                            new_kept_ranges.append((k_start + current_pos, k_start + m.start()))
                        current_pos = m.end()
                    
                    # Add the part after the last match
                    if current_pos < len(current_segment):
                        new_kept_ranges.append((k_start + current_pos, k_end))
                kept_ranges = new_kept_ranges
            
            # Apply the final kept_ranges to the text nodes
            original_full_text = ''.join([t.firstChild.nodeValue if t.firstChild else '' for t in text_nodes])
            _apply_kept_ranges_to_text_nodes(text_nodes, spans, kept_ranges)
            new_full_text = ''.join([t.firstChild.nodeValue if t.firstChild else '' for t in text_nodes])
            
            if original_full_text != new_full_text:
                changed_elements_count += 1
                
    return changed_elements_count

def classify_node(node):
    if node.nodeType != node.ELEMENT_NODE or node.tagName not in ['w:p', 'w:tbl']:
        return 'other'
    
    if node.tagName == 'w:tbl':
        return 'content'

    # It's a w:p
    text = get_all_text_from_element(node).strip()
    has_text = bool(text)
    has_drawing = bool(node.getElementsByTagName('w:drawing'))

    is_page_break = False
    for br in node.getElementsByTagName('w:br'):
        if br.getAttribute('w:type') == 'page':
            is_page_break = True
            break
    
    if not is_page_break:
        pPrs = node.getElementsByTagName('w:pPr')
        if pPrs:
            if pPrs[0].getElementsByTagName('w:sectPr'):
                is_page_break = True

    if is_page_break:
        if has_text or has_drawing:
            return 'content_and_break'
        else:
            return 'break'
    
    if has_text or has_drawing:
        return 'content'
    
    return 'empty_p' # Empty paragraph

def remove_blank_pages(body):
    nodes = list(body.childNodes)
    classifications = [classify_node(n) for n in nodes]

    nodes_to_remove = []
    for i in range(len(classifications) - 1):
        # Look for a break followed by another break, with only empty paragraphs in between
        if classifications[i] == 'break':
            # Find the next non-empty_p classification
            next_content_idx = -1
            for j in range(i + 1, len(classifications)):
                if classifications[j] != 'empty_p':
                    next_content_idx = j
                    break
            
            if next_content_idx != -1 and classifications[next_content_idx] in ['break', 'content_and_break']:
                # We have a break, followed by empties, followed by another break.
                # Remove the first break.
                nodes_to_remove.append(nodes[i])

    for node in nodes_to_remove:
        if node.parentNode:
            node.parentNode.removeChild(node)
    
    return len(nodes_to_remove)

def remove_all_empty_paragraphs(body):
    """Removes all paragraphs that contain no visible content."""
    nodes_to_remove = []
    for p in body.getElementsByTagName('w:p'):
        text = get_all_text_from_element(p).strip()
        has_drawing = p.getElementsByTagName('w:drawing')
        if not text and not has_drawing:
            # Ensure the paragraph is not a page break before removing
            if classify_node(p) not in ['break', 'content_and_break']:
                nodes_to_remove.append(p)
    
    for node in nodes_to_remove:
        if node.parentNode:
            node.parentNode.removeChild(node)
            
    return len(nodes_to_remove)

# ------------------------------
# Orchestrator
# ------------------------------

def process_document_xml(xml_path):
    print("Bắt đầu xử lý document.xml")
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    dom = minidom.parseString(content)
    body = dom.getElementsByTagName('w:body')[0]

    # 0) Trang đầu nếu chỉ có "thẻ 1"
    remove_first_page_if_the1(body)

    # *** BẮT ĐẦU THAY ĐỔI ***
    # 1) Xoá toàn bộ block [[BLOCK_START0]]...[[BLOCK_END]], bao gồm cả bảng
    # Lặp lại cho đến khi không còn cặp nào
    print("\nBước 1: Xử lý BLOCK_START0..BLOCK_END (bao gồm cả bảng)")
    total_removed_block = 0
    iteration = 0
    while True:
        iteration += 1
        removed_nodes_block = remove_nodes_between_tags(body, 'BLOCK_START', 'BLOCK_END', '0')
        total_removed_block += removed_nodes_block
        print(f"  [Lần {iteration}] Xử lý {removed_nodes_block} thay đổi")
        if removed_nodes_block == 0:
            break
    print(f"  Tổng cộng đã xoá/xử lý {total_removed_block} nodes/cặp")

    # 2) SECTION_START0..SECTION_END (hiện cũng xoá bao gồm cả bảng)
    # Lặp lại cho đến khi không còn cặp nào
    print("\nBước 2: Xử lý SECTION_START0..SECTION_END (bao gồm cả bảng)")
    total_removed_section = 0
    while True:
        removed_nodes_section = remove_nodes_between_tags(body, 'SECTION_START', 'SECTION_END', '0')
        total_removed_section += removed_nodes_section
        if removed_nodes_section == 0:
            break
    print(f"  Đã xoá {total_removed_section} nodes (đoạn, bảng) ở giữa các SECTION tag")
    # *** KẾT THÚC THAY ĐỔI ***

    # 3) Xoá hoàn toàn hàng [[ROW0]]
    print("\nBước 3: Xoá hoàn toàn các hàng có [[ROW0]]")
    rows_removed_0 = remove_rows_with_tag(body, '0')
    print(f"  Đã xoá {rows_removed_0} hàng ROW0")

    # 4) Xoá tag [[ROW1]] (giữ nội dung)
    print("\nBước 4: Xoá tag [[ROW1]] (giữ nội dung)")
    # No specific function call here, remove_all_remaining_tags will handle it.

    # 5) Gỡ tag còn lại (gồm cả [[ROW1]] etc.): replace tại w:t
    print("\nBước 5: Gỡ các tag còn lại")
    tags_changed = remove_all_remaining_tags(body)
    print(f"  Đã sửa {tags_changed} text nodes có tag")

    # 6) Xoá trang trắng
    print("\nBước 6: Xoá các trang trắng")
    pages_removed = remove_blank_pages(body)
    print(f"  Đã xoá {pages_removed} trang trắng")

    # 7) Dọn dẹp các đoạn văn trống
    print("\nBước 7: Dọn dẹp các đoạn văn trống")
    empty_paras_removed = remove_all_empty_paragraphs(body)
    print(f"  Đã xoá {empty_paras_removed} đoạn văn trống")

    # 8) Lưu lại
    print("\nBước 8: Lưu document.xml")
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())
    print("Hoàn thành xử lý document.xml")

# ------------------------------
# CLI
# ------------------------------

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