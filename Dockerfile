FROM python:3.11-slim-bookworm

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=0

# 必要系統套件（含 ICU 72）
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

# 建置期 smoke test（Render 免費版也看得到結果）
RUN python - <<'PY'
import aspose.cells as ac
wb = ac.Workbook()
print("Aspose smoke test OK. worksheets:", wb.worksheets.count)
PY

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
