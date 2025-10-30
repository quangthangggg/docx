#!/bin/bash

# Script khởi động DOCX Processor API

echo "🚀 Starting DOCX Processor API..."

# Tạo các thư mục cần thiết
mkdir -p uploads outputs zips logs
chmod 755 uploads outputs zips logs 2>/dev/null || true

# Kiểm tra xem đã cài đặt dependencies chưa
echo "📦 Checking dependencies..."
if ! python3 -c "import fastapi" 2>/dev/null; then
    echo "Installing dependencies..."
    pip3 install -r requirements.txt
fi

# Chạy application với 4 workers
echo ""
echo "✅ Starting server with 4 workers at http://0.0.0.0:8000"
echo "📱 Open your browser at: http://localhost:8000"
echo "🔍 Press Ctrl+C to stop"
echo ""

# Chạy với uvicorn và 4 workers
uvicorn app:app --host 0.0.0.0 --port 8000 --workers 4 --log-level info
