"""
將簡體中文 PDF 轉換為繁體中文（最終版 v2）。
策略：
1. 提取每頁所有文字 span 的完整資訊
2. 合併同行 span 做 OpenCC 轉換
3. 用 redact 清除整頁所有文字
4. 用 TextWriter 重新插入所有文字，利用原始 bbox 精確定位
   避免手動計算 x 偏移導致的文字重疊問題
"""

import sys
import os
import fitz
from opencc import OpenCC


def convert_pdf_s2t(input_path: str) -> str:
    cc = OpenCC('s2twp')

    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_繁體{ext}"

    doc = fitz.open(input_path)

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        # 第一步：收集所有 span，合併同行做轉換
        all_spans = []

        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                line_spans = line["spans"]
                if not line_spans:
                    continue

                # 合併整行文字做轉換
                full_text = "".join(s.get("text", "") for s in line_spans)
                full_converted = cc.convert(full_text)

                # 按 span 長度拆分
                pos = 0
                for span in line_spans:
                    original = span.get("text", "")
                    span_len = len(original)
                    converted_chunk = full_converted[pos:pos + span_len]
                    pos += span_len

                    all_spans.append({
                        "bbox": fitz.Rect(span["bbox"]),
                        "origin": fitz.Point(span["origin"]),
                        "size": span["size"],
                        "color": span["color"],
                        "text": converted_chunk,
                        "original": original,
                    })

        # 第二步：用 redact 清除所有文字 span
        for sp in all_spans:
            if sp["original"].strip():
                page.add_redact_annot(sp["bbox"])

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # 第三步：重新插入所有文字
        # 逐 span 插入，使用原始 bbox 寬度來決定字型大小縮放
        for sp in all_spans:
            text = sp["text"]
            if not text.strip():
                continue

            color_int = sp["color"]
            rgb = (((color_int >> 16) & 0xFF) / 255,
                   ((color_int >> 8) & 0xFF) / 255,
                   (color_int & 0xFF) / 255)

            font_size = sp["size"]
            origin = sp["origin"]
            bbox = sp["bbox"]
            available_width = bbox.width

            # 判斷文字是否包含 CJK 字元
            has_cjk = any(ord(ch) >= 0x2E80 for ch in text)

            if has_cjk:
                # 對包含 CJK 的文字，使用 china-t 字型
                # 先測量文字寬度，如果超出原始 bbox 則縮小字型
                fontname = "china-t"
                measured_width = fitz.get_text_length(
                    text, fontname=fontname, fontsize=font_size
                )

                if measured_width > 0 and available_width > 0:
                    # 如果測量寬度超出可用寬度，按比例縮放
                    ratio = available_width / measured_width
                    if ratio < 1.0:
                        # 文字太寬，需要縮小
                        adjusted_size = font_size * ratio
                    elif ratio > 1.5:
                        # 測量值偏小太多（china-t 常見問題），不調整
                        adjusted_size = font_size
                    else:
                        adjusted_size = font_size
                else:
                    adjusted_size = font_size

                # 使用 insert_textbox 讓 PyMuPDF 自動處理文字排列
                # 這比手動計算 x 偏移更可靠
                text_rect = fitz.Rect(
                    bbox.x0,
                    bbox.y0,
                    bbox.x0 + max(available_width, measured_width) * 1.05,
                    bbox.y1
                )

                page.insert_textbox(
                    text_rect,
                    text,
                    fontsize=adjusted_size,
                    fontname=fontname,
                    color=rgb,
                    align=fitz.TEXT_ALIGN_LEFT,
                )
            else:
                # 純 ASCII 文字，使用 helv 字型，直接在 origin 插入
                page.insert_text(
                    origin,
                    text,
                    fontsize=font_size,
                    fontname="helv",
                    color=rgb,
                )

        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 ({len(all_spans)} 個文字區塊)")

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert_pdf_s2t_final.py <PDF檔案路徑>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成！輸出檔案: {output}")
