# -*- coding: utf-8 -*-
# make_quote_linux.py
# 1) 用 Aspose.Cells 生成 .xlsx（插入列 = Insert Copied Cells 效果，圖片/圖形跟著移動縮放）
# 2) 預設用 LibreOffice (soffice --headless) 把「單一指定工作表」轉成 PDF（無浮水印）
# 3) 找不到 soffice 時，回退用 Aspose 匯出（會出紅字），並印出警告
#
# 依賴：
#   pip install aspose-cells-python
#   （非 pip）LibreOffice（若要無紅字 PDF）：Windows 用 winget/choco，Linux 用 apt/dnf
#
# 函式入口：
#   make_quote(xlsx_in, name=None, xlsx_out=None, pdf_out=None,
#              sheet=None, sets=None, items=None,
#              template_row=11, first_insert_row=12,
#              pdf_engine="libreoffice", soffice_path=None) -> (xlsx_out, pdf_out)

import argparse, os, sys, platform, subprocess, shutil, tempfile
from typing import Dict, List, Tuple
from datetime import datetime
from pathlib import Path

import aspose.cells as ac
from aspose.cells.drawing import PlacementType
from aspose.cells.rendering import SheetSet
from aspose.cells import FontConfigs

# ---------------- CLI 參數（仍保留相容） ----------------
def parse_set_args(sets: List[str]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for s in sets or []:
        if "=" in s:
            k, v = s.split("=", 1)
            out[k.strip()] = v.strip()
        else:
            print(f"[WARN] 忽略無效 --set: {s}", file=sys.stderr)
    return out

def parse_item_args(items: List[str]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for s in items or []:
        row: Dict[str, str] = {}
        for p in [p for p in s.split(",") if p.strip()]:
            if "=" in p:
                k, v = p.split("=", 1)
                row[k.strip()] = v.strip()
        if row:
            rows.append(row)
    return rows

# ---------------- 輔助：決定輸出路徑 ----------------
def decide_outputs(xlsx_in: str, name_arg: str | None, out_arg: str | None, pdf_arg: str | None) -> Tuple[str, str]:
    if name_arg:
        base = Path(name_arg)
        if base.suffix.lower() in (".xlsx", ".pdf", ".xlsm", ".xlsb", ".xls"):
            base = base.with_suffix("")
        xlsx_out = out_arg or str(base.with_suffix(".xlsx"))
        pdf_base = pdf_arg or str(base.with_suffix(".pdf"))
    else:
        xlsx_out = out_arg or f"{os.path.splitext(xlsx_in)[0]}_modified.xlsx"
        pdf_base = pdf_arg or f"{os.path.splitext(xlsx_out)[0]}.pdf"
    return xlsx_out, pdf_base

# ---------------- 輔助：字型設定（避免 Aspose 匯出 PDF 中文亂碼） ----------------
def setup_fonts_for_pdf() -> str | None:
    sysname = platform.system()
    try:
        if sysname == "Windows":
            FontConfigs.set_font_folder(r"C:\Windows\Fonts", True)
            return "Microsoft JhengHei"
        elif sysname == "Darwin":
            FontConfigs.set_font_folders(
                ["/System/Library/Fonts", "/Library/Fonts", os.path.expanduser("~/Library/Fonts")], True
            )
            return "PingFang TC"
        else:
            FontConfigs.set_font_folders(
                ["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.local/share/fonts")], True
            )
            return "Noto Sans CJK TC"
    except Exception:
        pass
    return None

# ---------------- 工作表取得（名稱/索引都通吃） ----------------
def _get_ws(wb: ac.Workbook, sheet_name: str | None) -> ac.Worksheet:
    if sheet_name:
        ws = wb.worksheets.get(sheet_name)
        if ws is None:
            raise ValueError(f"找不到工作表：{sheet_name}")
        return ws
    return wb.worksheets[0]

# ---------------- 基本操作（Aspose） ----------------
def open_book(path: str) -> ac.Workbook:
    return ac.Workbook(path)

def ensure_shapes_move_and_size(ws: ac.Worksheet):
    for shp in ws.shapes:
        try:
            shp.placement = PlacementType.MOVE_AND_SIZE
        except Exception:
            pass

def clear_row_contents(ws: ac.Worksheet, row_1based: int):
    r0 = row_1based - 1
    last_col = max(ws.cells.max_column, ws.cells.max_data_column)
    if last_col < 0:
        last_col = 0
    ws.cells.clear_contents(r0, 0, r0, last_col)  # 只清內容，不動格式/合併/列高

def insert_like_copied_cells(ws: ac.Worksheet, template_row_1based: int, first_insert_row_1based: int, extra_rows: int):
    if extra_rows <= 0:
        return
    t0 = template_row_1based - 1
    i0 = first_insert_row_1based - 1
    ws.cells.insert_rows(i0, extra_rows)  # 自動位移與更新參照
    for k in range(extra_rows):
        ws.cells.copy_row(ws.cells, t0, i0 + k)  # 複製範本列樣式到新列

def write_named_values(wb: ac.Workbook, updates: Dict[str, str]):
    for k, v in updates.items():
        rng = wb.worksheets.get_range_by_name(k)
        if rng is None:
            print(f"[WARN] 找不到 Named Range（或不是範圍）: {k}", file=sys.stderr)
            continue
        try:
            rng.value = v
            print(f"[WRITE] {k} -> R{rng.first_row+1}C{rng.first_column+1} = {v}")
        except Exception as e:
            print(f"[WARN] 無法寫入 {k}: {e}", file=sys.stderr)

def write_items_and_total(
    wb: ac.Workbook,
    sheet_name: str | None,
    items: List[Dict[str, str]],
    template_row: int = 11,
    first_insert_row: int = 12,
):
    ws = _get_ws(wb, sheet_name)
    ensure_shapes_move_and_size(ws)
    clear_row_contents(ws, template_row)

    extra = max(0, len(items) - 1)
    insert_like_copied_cells(ws, template_row, first_insert_row, extra)

    for i, it in enumerate(items):
        r0 = (template_row - 1) + i
        cells = ws.cells
        cells.get(r0, 0).put_value(i + 1)                                # 項次
        cells.get(r0, 1).put_value(it.get("Product", ""))                # 產品
        cells.get(r0, 2).put_value(it.get("Desc", ""))                   # 說明
        cnt = it.get("Count", None)
        try:
            cells.get(r0, 3).put_value(int(float(cnt)))
        except Exception:
            cells.get(r0, 3).put_value(cnt)
        pri = it.get("Price", None)
        try:
            cells.get(r0, 4).put_value(float(pri))
        except Exception:
            cells.get(r0, 4).put_value(pri)
        ppr = it.get("ProvidePrice", None)
        try:
            cells.get(r0, 5).put_value(float(ppr))
        except Exception:
            cells.get(r0, 5).put_value(ppr)

    rng_cnt  = wb.worksheets.get_range_by_name("Count")
    rng_prov = wb.worksheets.get_range_by_name("ProvidePrice")
    rng_fp   = wb.worksheets.get_range_by_name("FinalPrice")

    if len(items) > 0 and rng_fp is not None:
        start_row_1 = template_row
        end_row_1   = template_row + len(items) - 1
        cnt_col_1   = (rng_cnt.first_column + 1)  if rng_cnt  else 4
        prov_col_1  = (rng_prov.first_column + 1) if rng_prov else 6
        c = ws.cells.get(rng_fp.first_row, rng_fp.first_column)
        c.r1c1_formula = (
            f"=SUMPRODUCT(R{start_row_1}C{cnt_col_1}:R{end_row_1}C{cnt_col_1},"
            f"R{start_row_1}C{prov_col_1}:R{end_row_1}C{prov_col_1})"
        )
        wb.calculate_formula()
        print(f"[WRITE-TOTAL-FORMULA] FinalPrice = {c.r1c1_formula}")
    elif rng_fp is not None and len(items) == 0:
        ws.cells.get(rng_fp.first_row, rng_fp.first_column).put_value(0)
    else:
        print("[WARN] 找不到 Named Range: FinalPrice（略過公式寫入）")

# ---------------- PDF 匯出：A) Aspose（可能有紅字） ----------------
def export_sheet_to_pdf_aspose(wb: ac.Workbook, sheet_name: str | None, pdf_base_path: str) -> str:
    base = Path(pdf_base_path).resolve()
    base.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_pdf = base.with_name(f"{base.stem}_{ts}{base.suffix}")

    if sheet_name:
        ws = wb.worksheets.get(sheet_name)
        if ws is None:
            raise ValueError(f"找不到工作表：{sheet_name}")
        idx = ws.index
    else:
        idx = 0

    default_font = setup_fonts_for_pdf()
    wb.worksheets.active_sheet_index = idx
    opt = ac.PdfSaveOptions()
    opt.sheet_set = SheetSet([idx])
    if default_font:
        opt.default_font = default_font

    wb.calculate_formula()
    wb.save(str(out_pdf), opt)
    print(f"[PDF/Aspose] 已輸出：{out_pdf}（注意：若未授權，PDF 上方會有紅字）")
    return str(out_pdf)

# ---------------- PDF 匯出：B) LibreOffice（無紅字） ----------------
def find_soffice(explicit_path: str | None = None) -> str | None:
    if explicit_path:
        p = Path(explicit_path)
        return str(p) if p.exists() else None
    p = shutil.which("soffice")
    if p:
        return p
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice", "/snap/bin/libreoffice"
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None

def save_single_sheet_temp_xlsx(wb: ac.Workbook, sheet_name: str | None, tmp_dir: Path) -> Path:
    if not sheet_name:
        raise RuntimeError("未指定 sheet_name，請在上層決定來源 xlsx。")
    src_ws = _get_ws(wb, sheet_name)
    tmp_wb = ac.Workbook()
    tmp_ws = tmp_wb.worksheets[0]
    tmp_ws.copy(src_ws)
    tmp_ws.name = src_ws.name
    tmp_xlsx = tmp_dir / f"__single_sheet_{src_ws.name}_{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
    tmp_wb.save(str(tmp_xlsx))
    return tmp_xlsx

def export_sheet_to_pdf_libreoffice(wb: ac.Workbook, xlsx_out: str, sheet_name: str | None,
                                    pdf_base_path: str, soffice_path: str | None) -> str:
    base = Path(pdf_base_path).resolve()
    base.parent.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    final_pdf = base.with_name(f"{base.stem}_{ts}{base.suffix}")

    soffice = find_soffice(soffice_path)
    if not soffice:
        print("[WARN] 找不到 LibreOffice (soffice)。改用 Aspose 匯出（會有紅字）。", file=sys.stderr)
        return export_sheet_to_pdf_aspose(wb, sheet_name, pdf_base_path)

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        if sheet_name:
            src_xlsx = save_single_sheet_temp_xlsx(wb, sheet_name, td_path)
        else:
            src_xlsx = Path(xlsx_out)

        wb.calculate_formula()
        wb.save(xlsx_out)

        cmd = [
            soffice, "--headless", "--nologo", "--nodefault",
            "--nolockcheck", "--nofirststartwizard",
            "--convert-to", "pdf:calc_pdf_Export",
            "--outdir", str(td_path),
            str(src_xlsx)
        ]
        print(f"[PDF/LibreOffice] 執行：{' '.join(cmd)}")
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError as e:
            print(f"[ERROR] soffice 轉檔失敗：{e.stderr.decode(errors='ignore')}", file=sys.stderr)
            print("[WARN] 回退用 Aspose 匯出（會有紅字）。", file=sys.stderr)
            return export_sheet_to_pdf_aspose(wb, sheet_name, pdf_base_path)

        produced = td_path / (Path(src_xlsx).stem + ".pdf")
        if not produced.exists():
            print("[ERROR] 未找到 LibreOffice 產生的 PDF。回退 Aspose。", file=sys.stderr)
            return export_sheet_to_pdf_aspose(wb, sheet_name, pdf_base_path)

        shutil.move(str(produced), str(final_pdf))
        print(f"[PDF/LibreOffice] 已輸出：{final_pdf}")
        return str(final_pdf)

# ---------------- 核心：可呼叫的函式 ----------------
def make_quote(
    xlsx_in: str,
    name: str | None = None,
    xlsx_out: str | None = None,
    pdf_out: str | None = None,
    *,
    sheet: str | None = None,
    sets: Dict[str, str] | None = None,
    items: List[Dict[str, str]] | None = None,
    template_row: int = 11,
    first_insert_row: int = 12,
    pdf_engine: str = "libreoffice",
    soffice_path: str | None = None,
) -> Tuple[str, str]:
    """
    產生報價單：寫入命名儲存格、插入 item 列（等同 Insert Copied Cells），並輸出單一分頁 PDF（無紅字：libreoffice）。

    參數：
      xlsx_in        : 範本路徑
      name           : 輸出基本檔名（不含副檔名，可含路徑）。若提供，會同時決定 xlsx_out 與 pdf_out 的基底
      xlsx_out       : 指定 Excel 輸出路徑（可選）
      pdf_out        : 指定 PDF 基底路徑（可選；會自動加時間戳）
      sheet          : 目標分頁名稱（None = 第一張）
      sets           : 命名儲存格寫入，如 {"ClientName":"...", "QuoteDate":"..."}
      items          : 明細列 list[dict]，鍵包含 Product/Desc/Count/Price/ProvidePrice
      template_row   : 樣板列（預設 11）
      first_insert_row: 首筆插入列（預設 12）
      pdf_engine     : "libreoffice"（無紅字，預設）或 "aspose"
      soffice_path   : 指定 soffice 路徑（找不到 PATH 時可用）

    回傳：
      (xlsx_out_path, pdf_out_path)
    """
    xlsx_out_final, pdf_base = decide_outputs(xlsx_in, name, xlsx_out, pdf_out)
    updates = sets or {}
    items_list = items or []

    # 讀原檔
    wb = open_book(xlsx_in)

    # 存一份工作副本（若同名被占用就加時間戳）
    dst = Path(xlsx_out_final).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    candidate = dst
    while True:
        try:
            wb.save(str(candidate))
            break
        except Exception:
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            candidate = dst.with_name(f"{dst.stem}_{ts}{dst.suffix}")
    xlsx_out_final = str(candidate)

    # 對副本操作
    wb = ac.Workbook(xlsx_out_final)

    # 抬頭命名範圍
    if updates:
        write_named_values(wb, updates)

    # 明細
    if items_list:
        target_sheet_name = sheet if sheet else wb.worksheets[0].name
        write_items_and_total(
            wb,
            sheet_name=target_sheet_name,
            items=items_list,
            template_row=template_row,
            first_insert_row=first_insert_row,
        )

    # 存檔
    wb.save(xlsx_out_final)
    print(f"[DONE] Excel 已完成：{xlsx_out_final}")

    # 匯出 PDF
    if pdf_engine.lower() == "libreoffice":
        pdf_out_final = export_sheet_to_pdf_libreoffice(
            wb, xlsx_out_final, sheet if sheet else None, pdf_base, soffice_path
        )
    else:
        pdf_out_final = export_sheet_to_pdf_aspose(wb, sheet if sheet else None, pdf_base)

    print(f"[DONE] PDF 已完成：{pdf_out_final}")
    return xlsx_out_final, pdf_out_final

# ---------------- CLI 包裝（可選，用於相容原用法） ----------------
def main():
    ap = argparse.ArgumentParser(description="Excel 報價單產生器（Aspose 生成 + LibreOffice 無紅字轉 PDF）")
    ap.add_argument("--in", dest="xlsx_in", required=True)
    ap.add_argument("--out", dest="xlsx_out", default=None)
    ap.add_argument("--pdf-out", dest="pdf_out", default=None, help="PDF 基底路徑（會自動加時間戳）")
    ap.add_argument("--name", dest="name", default=None,
                    help="輸出檔名（可含路徑；可不含副檔名或含 .xlsx/.pdf）。若提供，將同時決定 Excel 與 PDF 的基本檔名；--out/--pdf-out 若也提供，會覆蓋此設定。")
    ap.add_argument("--sheet", dest="sheet", default=None, help="報價單所在工作表（預設第一張）")
    ap.add_argument("--template-row", type=int, default=11, help="樣板列（預設 11）")
    ap.add_argument("--first-insert-row", type=int, default=12, help="首筆插入列（預設 12）")
    ap.add_argument("--set", dest="sets", action="append", default=[])
    ap.add_argument("--item", dest="items", action="append", default=[])
    ap.add_argument("--pdf-engine", choices=["libreoffice", "aspose"], default="libreoffice",
                   help="PDF 轉檔引擎：libreoffice（無紅字，預設）或 aspose（可能有紅字）")
    ap.add_argument("--soffice", dest="soffice_path", default=None,
                   help="soffice 的路徑（找不到時可手動指定，例如 C:\\Program Files\\LibreOffice\\program\\soffice.exe）")
    args = ap.parse_args()

    # 轉換 CLI 的 --set/--item 為函式所需型別
    sets_dict = parse_set_args(args.sets)
    items_list = parse_item_args(args.items)

    make_quote(
        xlsx_in=args.xlsx_in,
        name=args.name,
        xlsx_out=args.xlsx_out,
        pdf_out=args.pdf_out,
        sheet=args.sheet,
        sets=sets_dict,
        items=items_list,
        template_row=args.template_row,
        first_insert_row=args.first_insert_row,
        pdf_engine=args.pdf_engine,
        soffice_path=args.soffice_path,
    )

if __name__ == "__main__":
    main()
