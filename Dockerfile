FROM python:3.11-slim

# 先確保 .NET 不依賴 ICU（你目前這樣可以順利跑起來）
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1

# 安裝 LibreOffice（只裝 calc + 共用元件就夠用）、中文字型與時區
#   - libreoffice-calc：需要用 calc 引擎把 xlsx 轉 pdf
#   - libreoffice-common：提供 soffice 啟動器與共用檔
#   - fonts-noto-cjk：避免中文缺字
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-common \
    fonts-noto-cjk \
    tzdata \
 && rm -rf /var/lib/apt/lists/*

# 檢查 soffice 是否存在（避免部署後才發現沒裝好）
RUN which soffice && soffice --headless --version

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render 的動態埠
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
