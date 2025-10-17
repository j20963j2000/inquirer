FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# LibreOffice（轉 PDF 無紅字）、中文字型、libgdiplus（Aspose 需要）、ICU（關鍵）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-noto-cjk \
    libgdiplus \
    libicu72 \
    tzdata \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
# 適應 Render/Railway 的動態埠
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
