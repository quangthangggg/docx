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
import main as docx_main_logic

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

import main as docx_main_logic



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



def process_document_xml(xml_path):

    """Xử lý file document.xml"""

    logger.info("Bắt đầu xử lý document.xml (thông qua main.py)")

    # Call the process_document_xml from main.py

    docx_main_logic.process_document_xml(xml_path)

    logger.info("Hoàn thành xử lý document.xml (thông qua main.py)")



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