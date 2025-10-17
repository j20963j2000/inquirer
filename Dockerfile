FROM python:3.11-slim

# 你現在的服務已經 OK，就先保留跳過 ICU 的設定（之後要走正式 ICU 再換）
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1

# 安裝 LibreOffice（整套最保險；體積較大但不用猜元件）
# 也裝中文字型避免缺字
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-noto-cjk \
    tzdata \
 && rm -rf /var/lib/apt/lists/*

# 建置期直接驗證 soffice 存在與版本（這兩行在 build log 會看到）
RUN which soffice
RUN soffice --headless --version

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render 會給 PORT 環境變數
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
