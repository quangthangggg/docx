# Sử dụng Python 3.11 slim image
FROM python:3.11-slim

# Thiết lập working directory
WORKDIR /app

# Copy requirements file
COPY requirements.txt .

# Cài đặt dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY app.py .
COPY main.py .
COPY index.html .

# Tạo các thư mục cần thiết
RUN mkdir -p uploads outputs logs zips

# Expose port
EXPOSE 8000

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python -c "import requests; requests.get('http://localhost:8000/health')" || exit 1

# Run application with 4 workers
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]
