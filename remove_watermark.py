import fitz  # PyMuPDF

def remove_watermark(input_pdf, output_to_user_pdf, watermark_text="Confidential"):
    # 打开 PDF 文件
    doc = fitz.open(input_pdf)
    
    # 遍历每一页
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(watermark_text)
        
        # 遍历找到的水印实例
        for inst in text_instances:
            page.add_redact_annot(inst, fill=(1, 1, 1))  # 用白色填充覆盖水印
            page.apply_redactions()
    
    # 保存修改后的 PDF
    doc.save(output_to_user_pdf)
    print(f"水印已成功移除，保存为 {output_to_user_pdf}")
    return output_to_user_pdf

if __name__ == "__main__":
    input_pdf = r"C:\Users\j2096\OneDrive\Desktop\QuotationBot\凱凱超級公司.pdf"
    output_pdf = "output.pdf"
    watermark_text = "Confidential"  # 这里替换为你的水印文本

    remove_watermark(input_pdf, output_pdf, watermark_text)
