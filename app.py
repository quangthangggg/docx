#!/usr/bin/env python3
"""
FastAPI application để xử lý file docx
UPDATED: Xử lý cả tables trong các khối tag "0" và xóa trang đầu nếu chỉ có "thẻ 1"
"""
from fastapi import FastAPI, File, UploadFile, HTTPException
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
    # Test write permission
    test_file = os.path.join(LOG_DIR, ".test")
    with open(test_file, 'w') as f:
        f.write('test')
    os.remove(test_file)
except (PermissionError, OSError):
    # Fallback to temp directory
    import tempfile
    LOG_DIR = tempfile.gettempdir()
    print(f"Warning: Using temp directory for logs: {LOG_DIR}")

# Tạo logger
logger = logging.getLogger("docx_processor")
logger.setLevel(logging.INFO)

# Console handler (always works)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Formatter
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
console_handler.setFormatter(formatter)

# Add console handler
logger.addHandler(console_handler)

# File handler với rotation (optional)
try:
    file_handler = RotatingFileHandler(
        os.path.join(LOG_DIR, "app.log"),
        maxBytes=10*1024*1024,  # 10MB
        backupCount=5
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
except (PermissionError, OSError) as e:
    logger.warning(f"Could not create log file: {e}. Logging to console only.")

app = FastAPI(title="DOCX Processor API", description="API để xử lý file DOCX")

# Tạo thư mục để lưu file tạm
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
ZIP_DIR = "zips"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ZIP_DIR, exist_ok=True)

# Thread pool cho xử lý song song
executor = ThreadPoolExecutor(max_workers=4)

logger.info("Application started with 4 workers")

# ===== CÁC HÀM XỬ LÝ - UPDATED =====

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
    logger.info("Bước 0: Kiểm tra và xóa trang đầu nếu chỉ có 'thẻ 1'")
    
    first_page_elements = get_first_page_elements(body)
    
    if not first_page_elements:
        logger.info("  Không tìm thấy elements ở trang đầu")
        return
    
    # Lấy tất cả text từ trang đầu
    all_text = ''
    for element in first_page_elements:
        all_text += get_all_text_from_element(element)
    
    # Normalize text (strip whitespace và lowercase)
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

    # Xóa trang đầu nếu chỉ có "thẻ 1"
    remove_first_page_if_the1(body)

    logger.info("Bước 1: Đánh dấu các BLOCK_START0 cần xóa (bao gồm cả tables)")
    block_element_ids = mark_elements_for_removal(body, 'BLOCK_START', 'BLOCK_END', '0')
    logger.info(f"  Tìm thấy {len(block_element_ids)} elements trong BLOCK_START0")

    logger.info("Bước 2: Đánh dấu các SECTION_START0 cần xóa (bao gồm cả tables)")
    section_element_ids = mark_elements_for_removal(body, 'SECTION_START', 'SECTION_END', '0')
    logger.info(f"  Tìm thấy {len(section_element_ids)} elements trong SECTION_START0")

    logger.info("Bước 3: Xóa các elements đã đánh dấu (paragraphs và tables)")
    all_elements_to_remove = block_element_ids | section_element_ids

    # Xóa các paragraphs và tables đã được đánh dấu
    elements_removed = 0
    for child in list(body.childNodes):
        if child.nodeType == child.ELEMENT_NODE:
            if child.tagName in ['w:p', 'w:tbl']:
                if id(child) in all_elements_to_remove:
                    body.removeChild(child)
                    elements_removed += 1
    logger.info(f"  Đã xóa {elements_removed} elements (paragraphs và tables)")

    logger.info("Bước 4: Xóa các table row có ROW0 (trong các bảng còn lại)")
    tables = body.getElementsByTagName('w:tbl')
    rows_removed = 0
    for table in tables:
        rows = list(table.getElementsByTagName('w:tr'))
        for row in rows:
            row_text = get_all_text_from_element(row)

            if has_tag_pattern(row_text, 'ROW', '0'):
                row.parentNode.removeChild(row)
                rows_removed += 1
    logger.info(f"  Đã xóa {rows_removed} rows có ROW0")

    logger.info("Bước 5: Xóa tất cả các tag còn lại")
    all_text_elems = body.getElementsByTagName('w:t')
    tags_removed = 0
    for text_elem in all_text_elems:
        if text_elem.firstChild and text_elem.firstChild.nodeValue:
            original_text = text_elem.firstChild.nodeValue
            cleaned_text = remove_all_tags(original_text)
            if cleaned_text != original_text:
                tags_removed += 1
                if cleaned_text:
                    text_elem.firstChild.nodeValue = cleaned_text
                else:
                    parent = text_elem.parentNode
                    if parent:
                        parent.removeChild(text_elem)
    logger.info(f"  Đã xóa tags từ {tags_removed} text elements")

    logger.info("Bước 6: Lưu document.xml")
    with open(xml_path, 'w', encoding='utf-8') as f:
        f.write(dom.toxml())
    logger.info("Hoàn thành xử lý document.xml")

def process_docx_file(input_path, output_path):
    """Xử lý file docx từ input_path và lưu vào output_path"""
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

# ===== API ENDPOINTS =====

@app.get("/", response_class=HTMLResponse)
async def index():
    """Trang chủ với giao diện upload file"""
    with open("index.html", "r", encoding="utf-8") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

@app.post("/process")
async def process_file(file: UploadFile = File(...)):
    """
    Endpoint để xử lý file docx được upload
    """
    logger.info(f"Nhận request xử lý file: {file.filename}")

    # Kiểm tra định dạng file
    if not file.filename.endswith('.docx'):
        logger.warning(f"File không hợp lệ: {file.filename}")
        raise HTTPException(status_code=400, detail="Chỉ chấp nhận file .docx")

    # Tạo tên file unique
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    input_filename = f"input_{timestamp}.docx"
    output_filename = f"processed_{timestamp}.docx"

    input_path = os.path.join(UPLOAD_DIR, input_filename)
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        # Lưu file upload
        logger.info(f"Lưu file upload: {input_filename}")
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Xử lý file
        process_docx_file(input_path, output_path)

        # Xóa file input sau khi xử lý
        os.remove(input_path)
        logger.info(f"Xử lý thành công file: {file.filename} -> {output_filename}")

        return {
            "message": "Xử lý file thành công",
            "output_filename": output_filename
        }

    except Exception as e:
        logger.error(f"Lỗi khi xử lý file {file.filename}: {str(e)}", exc_info=True)
        # Xóa file nếu có lỗi
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)

        raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý file: {str(e)}")

@app.post("/process-multiple")
async def process_multiple_files(files: List[UploadFile] = File(...)):
    """
    Endpoint để xử lý nhiều file docx song song
    """
    logger.info(f"Nhận request xử lý {len(files)} file(s)")

    # Validate files
    for file in files:
        if not file.filename.endswith('.docx'):
            logger.warning(f"File không hợp lệ: {file.filename}")
            raise HTTPException(status_code=400, detail=f"File {file.filename} không phải .docx")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Nếu chỉ có 1 file, xử lý đơn giản
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

            # Xử lý file
            process_docx_file(input_path, output_path)

            # Xóa file input
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

    # Xử lý nhiều file song song
    else:
        try:
            # Lưu tất cả files
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

            # Xử lý song song
            logger.info(f"Bắt đầu xử lý song song {len(files)} files với ThreadPoolExecutor")
            loop = asyncio.get_event_loop()

            async def process_file_async(inp, outp):
                return await loop.run_in_executor(executor, process_docx_file, inp, outp)

            # Chạy song song
            tasks = [process_file_async(inp, outp[0]) for inp, outp in zip(input_paths, output_paths)]
            await asyncio.gather(*tasks)

            # Xóa input files
            for input_path in input_paths:
                if os.path.exists(input_path):
                    os.remove(input_path)

            # Tạo file zip
            zip_filename = f"processed_{timestamp}.zip"
            zip_path = os.path.join(ZIP_DIR, zip_filename)

            logger.info(f"Tạo file zip: {zip_filename}")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for output_path, original_name in output_paths:
                    if os.path.exists(output_path):
                        # Giữ tên gốc trong zip
                        arcname = f"processed_{original_name}"
                        zipf.write(output_path, arcname)
                        # Xóa file sau khi đã add vào zip
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

            # Cleanup
            for input_path in input_paths:
                if os.path.exists(input_path):
                    os.remove(input_path)
            for output_path, _ in output_paths:
                if os.path.exists(output_path):
                    os.remove(output_path)

            raise HTTPException(status_code=500, detail=f"Lỗi khi xử lý files: {str(e)}")

@app.get("/download/{filename}")
async def download_file(filename: str):
    """
    Endpoint để tải file đã xử lý
    """
    file_path = os.path.join(OUTPUT_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File không tồn tại")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.get("/download-zip/{filename}")
async def download_zip(filename: str):
    """
    Endpoint để tải file zip
    """
    file_path = os.path.join(ZIP_DIR, filename)

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File không tồn tại")

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
        # Đọc 50 dòng cuối của log file
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            recent_logs = ''.join(lines[-50:]) if len(lines) > 50 else ''.join(lines)

        return {"logs": recent_logs}
    except Exception as e:
        logger.error(f"Lỗi khi đọc logs: {str(e)}")
        return {"logs": f"Lỗi khi đọc logs: {str(e)}"}

if __name__ == "__main__":
    import uvicorn
    # Note: workers parameter doesn't work with uvicorn.run()
    # Use: uvicorn app:app --host 0.0.0.0 --port 8000 --workers 4
    logger.info("Starting server on http://0.0.0.0:8000")
    logger.info("For production with workers, use: uvicorn app:app --host 0.0.0.0 --port 8000 --workers 4")
    uvicorn.run(app, host="0.0.0.0", port=8000)