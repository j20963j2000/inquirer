FROM python:3.11-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=0

# 必要系統套件：
# - libreoffice：把單一工作表轉 PDF（無紅字）
# - fonts-noto-cjk：中文字型
# - libgdiplus：Aspose 依賴
# - libicu-dev：ICU（重點，trixie 會裝到 libicu74，bookworm 會裝到 libicu72）
# - tzdata：時區
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-noto-cjk \
    libgdiplus \
    libicu-dev \
    tzdata \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 先複製程式碼
COPY . .

# 🔎 建置期 smoke test：確保 Aspose + ICU 正常，避免跑到執行階段才爆。
RUN python - <<'PY'
import aspose.cells as ac
wb = ac.Workbook()
print("Aspose smoke test OK. worksheets:", wb.worksheets.count)
PY

# 相容雲端平台的動態埠
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
