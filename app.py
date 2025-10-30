#!/usr/bin/env python3
"""
FastAPI application để xử lý file docx
UPDATED: Xử lý đúng nhiều cặp START-END liên tiếp trong cùng paragraph
"""
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from typing import List
import os
import tempfile
import shutil
from datetime import datetime
import zipfile
from defusedxml import minidom
import re
import logging
from logging.handlers import RotatingFileHandler
import asyncio
from concurrent.futures import ThreadPoolExecutor
import io

# Cấu hình logging
LOG_DIR = "logs"
try:
    os.makedirs(LOG_DIR, exist_ok=True)
    test_file = os.path.join(LOG_DIR, ".test")
    with open(test_file, 'w') as f:
        f.write('test')
    os.remove(test_file)
except (PermissionError, OSError):
    import tempfile
    LOG_DIR = tempfile.gettempdir()
    print(f"Warning: Using temp directory for logs: {LOG_DIR}")

# Tạo logger
logger = logging.getLogger("docx_processor")
logger.setLevel(logging.INFO)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

try:
    file_handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "app.log"),
        maxBytes=10*1024*1024,
        backupCount=5
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
except (PermissionError, OSError) as e:
    logger.warning(f"Could not create log file: {e}. Logging to console only.")

app = FastAPI(title="DOCX Processor API", description="API để xử lý file DOCX")

# Tạo thư mục
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
ZIP_DIR = "zips"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ZIP_DIR, exist_ok=True)

executor = ThreadPoolExecutor(max_workers=4)

logger.info("Application started with 4 workers")

# ===== CÁC HÀM XỬ LÝ - UPDATED =====

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
        else:  # Nếu text rỗng
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
            logger.info(f"  Xử lý cặp {start_tag_type}{label}-{end_tag_type} trong cùng element")
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
            logger.info(f"  Tìm thấy {start_tag_type}{label}, tìm {end_tag_type} ở các element sau")
            
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
                    logger.info(f"    Tìm thấy {end_tag_type} tại element {j}")
                    
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
                logger.warning(f"  Không tìm thấy {end_tag_type} tương ứng với {start_tag_type}{label}")
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
                # Nếu element chỉ có tags và không còn gì
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
    logger.info("Bước 0: Kiểm tra và xóa trang đầu nếu chỉ có 'thẻ 1'")
    
    first_page_elements = get_first_page_elements(body)
    
    if not first_page_elements:
        logger.info("  Không tìm thấy elements ở trang đầu")
        return
    
    all_text = ''
    for element in first_page_elements:
        all_text += get_all_text_from_element(element)
    
    normalized_text = all_text.strip().lower()
    
    logger.info(f"  Nội dung trang đầu: '{normalized_text}'")
    
    if normalized_text == 'thẻ 1':
        logger.info("  ✓ Phát hiện trang đầu chỉ có 'thẻ 1', đang xóa...")
        for element in first_page_elements:
            body.removeChild(element)
        logger.info(f"  Đã xóa {len(first_page_elements)} elements từ trang đầu")
    else:
        logger.info("  Trang đầu không phải chỉ có 'thẻ 1', giữ nguyên")

def process_document_xml(xml_path):
    """Xử lý file document.xml"""
    logger.info("Bắt đầu xử lý document.xml")
    with open(xml_path, 'r', encoding='utf-8') as f:
        content = f.read()

    dom = minidom.parseString(content)
    body = dom.getElementsByTagName('w:body')[0]

    # Bước 0: Xóa trang đầu nếu chỉ có "thẻ 1"
    remove_first_page_if_the1(body)

    # Bước 1: Xử lý BLOCK_START0 và BLOCK_END
    logger.info("Bước 1: Xử lý xóa nội dung giữa BLOCK_START0 và BLOCK_END")
    elements_removed = process_removal_between_tags(body, 'BLOCK_START', 'BLOCK_END', '0')
    logger.info(f"  Đã xóa {elements_removed} elements giữa các BLOCK tags")

    # Bước 2: Xử lý SECTION_START0 và SECTION_END
    logger.info("Bước 2: Xử lý xóa nội dung giữa SECTION_START0 và SECTION_END")
    elements_removed = process_removal_between_tags(body, 'SECTION_START', 'SECTION_END', '0')
    logger.info(f"  Đã xóa {elements_removed} elements giữa các SECTION tags")

    # Bước 3: Xóa các table row có ROW0
    logger.info("Bước 3: Xóa các table row có ROW0")
    rows_removed = remove_rows_with_tag(body, '0')
    logger.info(f"  Đã xóa {rows_removed} rows có ROW0")

    # Bước 4: Xóa tất cả các tag còn lại
    logger.info("Bước 4: Xóa tất cả các tag còn lại")
    tags_removed = remove_all_remaining_tags(body)
    logger.info(f"  Đã xóa tags từ {tags_removed} elements")

    # Bước 5: Lưu document.xml
    logger.info("Bước 5: Lưu document.xml")
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())
    logger.info("Hoàn thành xử lý document.xml")

def process_docx_file(input_path, output_path):
    """Xử lý file docx"""
    logger.info(f"Bắt đầu xử lý file: {input_path}")
    temp_dir = tempfile.mkdtemp()

    try:
        logger.info("Đang giải nén file docx...")
        unpack_docx(input_path, temp_dir)

        doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        if not os.path.exists(doc_xml_path):
            logger.error("Không tìm thấy word/document.xml trong file docx")
            raise Exception("Không tìm thấy word/document.xml trong file docx")

        process_document_xml(doc_xml_path)

        logger.info(f"Đang tạo file output: {output_path}")
        pack_docx(temp_dir, output_path)

        logger.info(f"✅ Hoàn thành! File đã được lưu tại: {output_path}")

    finally:
        shutil.rmtree(temp_dir)
        logger.info(f"Đã dọn dẹp thư mục tạm: {temp_dir}")

def cleanup_file(file_path: str):
    """Xóa file sau khi download xong"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Đã xóa file sau khi download: {file_path}")
    except Exception as e:
        logger.error(f"Lỗi khi xóa file {file_path}: {str(e)}")

# ===== API ENDPOINTS =====

@app.get("/", response_class=HTMLResponse)
async def index():
    """Trang chủ với giao diện upload file"""
    with open("index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

@app.post("/process")
async def process_file(file: UploadFile = File(...)):
    """Endpoint để xử lý file docx được upload"""
    logger.info(f"Nhận request xử lý file: {file.filename}")

    if not file.filename.endswith('.docx'):
        logger.warning(f"File không hợp lệ: {file.filename}")
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .docx")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_filename = f"input_{timestamp}.docx"
    output_filename = f"processed_{timestamp}.docx"

    input_path = os.path.join(UPLOAD_DIR, input_filename)
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        logger.info(f"Lưu file upload: {input_filename}")
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        process_docx_file(input_path, output_path)

        os.remove(input_path)
        logger.info(f"Xử lý thành công file: {file.filename} -> {output_filename}")

        return {
            "message": "Xử lý file thành công",
            "output_filename": output_filename
        }

    except Exception as e:
        logger.error(f"Lỗi khi xử lý file {file.filename}: {str(e)}", exc_info=True)
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)

        raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file: {str(e)}")

@app.post("/process-multiple")
async def process_multiple_files(files: List[UploadFile] = File(...)):
    """Endpoint để xử lý nhiều file docx song song"""
    logger.info(f"Nhận request xử lý {len(files)} file(s)")

    for file in files:
        if not file.filename.endswith('.docx'):
            logger.warning(f"File không hợp lệ: {file.filename}")
            raise HTTPException(status_code=400, detail=f"File {file.filename} không phải .docx")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if len(files) == 1:
        file = files[0]
        input_filename = f"input_{timestamp}.docx"
        output_filename = f"processed_{timestamp}.docx"

        input_path = os.path.join(UPLOAD_DIR, input_filename)
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        try:
            logger.info(f"Lưu file upload: {input_filename}")
            with open(input_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            process_docx_file(input_path, output_path)

            os.remove(input_path)
            logger.info(f"Xử lý thành công file: {file.filename} -> {output_filename}")

            return {
                "message": "Xử lý file thành công",
                "output_filename": output_filename,
                "processed_count": 1,
                "is_zip": False
            }

        except Exception as e:
            logger.error(f"Lỗi khi xử lý file {file.filename}: {str(e)}", exc_info=True)
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
            raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file: {str(e)}")

    else:
        try:
            input_paths = []
            output_paths = []

            for idx, file in enumerate(files):
                input_filename = f"input_{timestamp}_{idx}.docx"
                output_filename = f"processed_{timestamp}_{idx}.docx"

                input_path = os.path.join(UPLOAD_DIR, input_filename)
                output_path = os.path.join(OUTPUT_DIR, output_filename)

                logger.info(f"Lưu file upload {idx + 1}/{len(files)}: {input_filename}")
                with open(input_path, "wb") as buffer:
                    shutil.copyfileobj(file.file, buffer)

                input_paths.append(input_path)
                output_paths.append((output_path, file.filename))

            logger.info(f"Bắt đầu xử lý song song {len(files)} files với ThreadPoolExecutor")
            loop = asyncio.get_event_loop()

            async def process_file_async(inp, outp):
                return await loop.run_in_executor(executor, process_docx_file, inp, outp)

            tasks = [process_file_async(inp, outp[0]) for inp, outp in zip(input_paths, output_paths)]
            await asyncio.gather(*tasks)

            for input_path in input_paths:
                if os.path.exists(input_path):
                    os.remove(input_path)

            zip_filename = f"processed_{timestamp}.zip"
            zip_path = os.path.join(ZIP_DIR, zip_filename)

            logger.info(f"Tạo file zip: {zip_filename}")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for output_path, original_name in output_paths:
                    if os.path.exists(output_path):
                        arcname = f"processed_{original_name}"
                        zipf.write(output_path, arcname)
                        os.remove(output_path)

            logger.info(f"Xử lý thành công {len(files)} files -> {zip_filename}")

            return {
                "message": f"Xử lý thành công {len(files)} files",
                "output_filename": zip_filename,
                "processed_count": len(files),
                "is_zip": True
            }

        except Exception as e:
            logger.error(f"Lỗi khi xử lý nhiều files: {str(e)}", exc_info=True)

            for input_path in input_paths:
                if os.path.exists(input_path):
                    os.remove(input_path)
            for output_path, _ in output_paths:
                if os.path.exists(output_path):
                    os.remove(output_path)

            raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý files: {str(e)}")

@app.get("/download/{filename}")
async def download_file(filename: str, background_tasks: BackgroundTasks):
    """Endpoint để tải file đã xử lý"""
    file_path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File không tồn tại")

    background_tasks.add_task(cleanup_file, file_path)
    
    logger.info(f"Đang gửi file để download: {filename}")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/download-zip/{filename}")
async def download_zip(filename: str, background_tasks: BackgroundTasks):
    """Endpoint để tải file zip"""
    file_path = os.path.join(ZIP_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File không tồn tại")

    background_tasks.add_task(cleanup_file, file_path)
    
    logger.info(f"Đang gửi file zip để download: {filename}")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/zip"
    )

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "ok"}

@app.get("/logs")
async def get_logs():
    """Endpoint để xem logs gần đây"""
    log_file = os.path.join(LOG_DIR, "app.log")

    if not os.path.exists(log_file):
        return {"logs": "Không có logs nào."}

    try:
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            recent_logs = ''.join(lines[-50:]) if len(lines) > 50 else ''.join(lines)

        return {"logs": recent_logs}
    except Exception as e:
        logger.error(f"Lỗi khi đọc logs: {str(e)}")
        return {"logs": f"Lỗi khi đọc logs: {str(e)}"}

if __name__ == "__main__":
    import uvicorn
    logger.info("Starting server on http://0.0.0.0:8000")
    logger.info("For production with workers, use: uvicorn app:app --host 0.0.0.0 --port 8000 --workers 4")
    uvicorn.run(app, host="0.0.0.0", port=8000)