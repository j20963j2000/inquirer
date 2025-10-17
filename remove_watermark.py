# remove_mark.py
from pathlib import Path
import fitz  # PyMuPDF

def remove_watermark(input_pdf: str | Path,
                     watermark_text: str = "Confidential",
                     out_dir: str | Path | None = None) -> str:
    """
    只移除 PDF 中等於 watermark_text 的文字水印。
    - 支援旋轉文字（用 quads）
    - 每頁先收集再一次套用遮蔽
    - 產生唯一輸出檔名：<原檔名>_clean.pdf
    回傳：輸出檔案的絕對路徑字串
    """
    input_pdf = Path(input_pdf)
    out_dir = Path(out_dir) if out_dir else input_pdf.parent
    out_path = out_dir / f"{input_pdf.stem}_clean.pdf"

    doc = fitz.open(input_pdf)
    total_hits = 0

    for page in doc:
        # 用 quads 抓旋轉/斜體的文字
        quads = page.search_for(watermark_text, quads=True)
        # 放大一點點，避免邊緣殘留
        for q in quads:
            rect = q.rect + (-1, -1, 1, 1)
            page.add_redact_annot(rect, fill=(1, 1, 1))
        if quads:
            page.apply_redactions()
            total_hits += len(quads)

    # 垃圾回收與壓縮，讓檔案變小
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()
    print(f"[watermark] removed {total_hits} hits -> {out_path}")
    return str(out_path)


if __name__ == "__main__":
    input_pdf = r"C:\Users\j2096\OneDrive\Desktop\QuotationBot\linebot\報價單_凱凱_20251017-145756.pdf"
    output_pdf = "output.pdf"
    watermark_text = "Confidential"  # 这里替换为你的水印文本

    remove_watermark(input_pdf, output_pdf, watermark_text)
