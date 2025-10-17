# user_input_parsing.py
import re
from typing import Dict, List, Tuple

# 支援全形/半形冒號
_COLON = r"[:：]"

# 中文鍵 ↔ 欄位名稱對應
HEADER_MAP = {
    "客戶名稱": "ClientName",
    "客戶": "ClientName",
    "客戶名": "ClientName",
    "公司": "ClientName",
    "to": "ClientName",
    "報價時間": "QuoteDate",
    "報價日期": "QuoteDate",
    "日期": "QuoteDate",
}

ITEM_MAP = {
    "產品": "Product",
    "品名": "Product",
    "商品": "Product",
    "說明": "Desc",
    "描述": "Desc",
    "備註": "Desc",
    "數量": "Count",
    "qty": "Count",
    "單價": "Price",
    "價格": "Price",
    "優惠單價": "ProvidePrice",
    "報價": "ProvidePrice",
    "特價": "ProvidePrice",
    "成交價": "ProvidePrice",
}

def _norm_key(k: str) -> str:
    return k.strip().lower()

def parse_user_text(text: str) -> Tuple[Dict[str, str], List[Dict[str, str]]]:
    """
    解析使用者貼在 LINE 的文字：
    - 頭部欄位：客戶名稱/報價時間 ...
    - 多個商品區塊以 '----' 或空白行分隔
    回傳： (sets_dict, items_list)
    """
    # 先切成行，清掉 BOM 與多餘空白
    lines = [re.sub(r"\ufeff", "", l).strip() for l in text.splitlines()]
    blocks: List[List[str]] = []
    buf: List[str] = []
    for ln in lines:
        if not ln or re.fullmatch(r"-{3,}", ln):
            if buf:
                blocks.append(buf); buf = []
            continue
        buf.append(ln)
    if buf: blocks.append(buf)

    sets: Dict[str, str] = {}
    items: List[Dict[str, str]] = []

    # 第一個區塊可能同時含 header 與第一個 item
    current_item: Dict[str, str] = {}

    def flush_item():
        nonlocal current_item
        if any(v for v in current_item.values()):
            items.append(current_item)
        current_item = {}

    for blk in blocks:
        for ln in blk:
            m = re.split(_COLON, ln, maxsplit=1)
            if len(m) != 2:
                continue
            k_raw, v_raw = m[0].strip(), m[1].strip()
            k = _norm_key(k_raw)

            # 先試 header
            key = HEADER_MAP.get(k)
            if key:
                sets[key] = v_raw
                continue

            # 再試 item
            key = ITEM_MAP.get(k)
            if key:
                # 如果看到「產品」且當前已有產品，視為新商品開始
                if key == "Product" and current_item.get("Product"):
                    flush_item()
                current_item[key] = v_raw
                continue

        # 一個區塊結束就 flush 一次
        flush_item()

    # 型別微調
    for it in items:
        if "Count" in it:
            try: it["Count"] = int(float(str(it["Count"]).replace(",", "")))
            except: pass
        for price_key in ("Price", "ProvidePrice"):
            if price_key in it:
                try: it[price_key] = float(str(it[price_key]).replace(",", ""))
                except: pass

    return sets, items
