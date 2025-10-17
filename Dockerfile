# Dockerfile
FROM python:3.11-slim

# 系統套件：LibreOffice（無紅字轉 PDF）、中文字型、libgdiplus（Aspose 需要）
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice fonts-noto-cjk libgdiplus && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
ENV PORT=8000
# 讓 uvicorn 運行 FastAPI
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
