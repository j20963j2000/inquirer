# app.py
import os
from pathlib import Path
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import PlainTextResponse
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv

from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage

from user_input_parsing import parse_user_text
from make_quote_linux import make_quote
from remove_watermark import remove_watermark

# ---- 環境變數 ----
load_dotenv()
CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
CHANNEL_TOKEN  = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
TEMPLATE_XLSX  = os.getenv("TEMPLATE_XLSX", "維修報價單範本.xlsx")
SHEET_NAME     = os.getenv("SHEET_NAME")  # 例如：貝拉5；不填=第一張
PDF_ENGINE     = os.getenv("PDF_ENGINE", "libreoffice")  # 或 aspose
SOFFICE_PATH   = os.getenv("SOFFICE_PATH")  # 例如 /usr/bin/soffice 或 Windows 的路徑
PUBLIC_BASE_URL= os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")  # 給 LINE 用的可公開網址
OUTPUT_DIR     = os.getenv("OUTPUT_DIR", "public")

if not CHANNEL_SECRET or not CHANNEL_TOKEN:
    raise RuntimeError("請設定 LINE_CHANNEL_SECRET / LINE_CHANNEL_ACCESS_TOKEN")

# ---- 準備目錄與 LINE SDK ----
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
line_bot_api = LineBotApi(CHANNEL_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# ---- FastAPI ----
app = FastAPI(title="QuotationBot")

# 靜態檔案（下載用）
app.mount("/files", StaticFiles(directory=OUTPUT_DIR), name="files")

@app.get("/healthz")
async def health():
    return {"ok": True}

@app.post("/callback")
async def callback(request: Request):
    signature = request.headers.get("X-Line-Signature", "")
    body = (await request.body()).decode("utf-8")
    try:
        handler.handle(body, signature)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    return PlainTextResponse("OK")

# ---- LINE 事件處理 ----
@handler.add(MessageEvent, message=TextMessage)
def on_text(event: MessageEvent):
    text = event.message.text
    sets, items = parse_user_text(text)

    if not items:
        example = (
            "請貼上如下格式：\n"
            "客戶名稱: 凱凱超級公司\n"
            "報價時間: 2030-01-01\n"
            "產品: 嬤嬤啦餐飲配送機器人\n說明: 快拆電池保護蓋組件\n數量: 2\n單價: 3150\n優惠單價: 3000\n"
            "----\n"
            "產品: 嬤嬤啦餐飲配送機器人\n說明: 快跑電池保護蓋組件\n數量: 2\n單價: 500\n優惠單價: 450"
        )
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="我需要至少一筆產品資訊喔～\n\n" + example)
        )
        return

    # 檔名基底（放 public/ 下）
    client = sets.get("ClientName", "客戶")
    base = f"{client}_{event.timestamp}"  # 保證唯一
    base_path = str(Path(OUTPUT_DIR) / base)

    try:
        xlsx_out, pdf_out = make_quote(
            xlsx_in=TEMPLATE_XLSX,
            name=base_path,               # 讓輸出直接存到 public/ 下
            sheet=SHEET_NAME,
            sets=sets,
            items=items,
            pdf_engine=PDF_ENGINE,
            soffice_path=SOFFICE_PATH,
        )
    except Exception as e:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=f"產生報價單失敗：{e}")
        )
        return

    # 轉為可下載 URL
    xlsx_url = f"{PUBLIC_BASE_URL}/files/{Path(xlsx_out).name}"
    pdf_url  = f"{PUBLIC_BASE_URL}/files/{Path(pdf_out).name}"

    p = Path(pdf_out)  # 例如 /app/public/凱凱超級公司_....pdf
    output_to_user_pdf = str(p.with_name(p.stem + "_clean.pdf"))

    output_to_user_pdf = remove_watermark(input_pdf = str(p), output_to_user_pdf = output_to_user_pdf)
    pdf_url_to_user  = f"{PUBLIC_BASE_URL}/files/{Path(output_to_user_pdf).name}"

    msg = (
        "✅ 報價單已完成！\n"
        f"Excel：{xlsx_url}\n"
        f"PDF：{pdf_url_to_user}\n"
        "(連結有效取決於你伺服器是否持續運作)"
    )
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
