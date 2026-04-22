"""
將簡體中文 PDF 轉換為繁體中文 (v6)。
改進：先將同一行的所有 span 文字合併後再做 OpenCC 轉換，
然後按原始 span 長度拆分回去，確保跨 span 的詞彙也能正確轉換。
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
        replaced = 0
        total = 0

        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans:
                    continue

                # 合併同一行所有 span 的文字
                full_text = "".join(s.get("text", "") for s in spans)
                full_converted = cc.convert(full_text)

                # 如果整行都沒變化，跳過
                if full_converted == full_text:
                    total += len([s for s in spans if s.get("text", "").strip()])
                    continue

                # 按原始 span 的字元長度拆分轉換後的文字
                # 注意：簡繁轉換通常是 1:1 字元對應
                pos = 0
                for span in spans:
                    original = span.get("text", "")
                    if not original:
                        continue

                    span_len = len(original)
                    converted_chunk = full_converted[pos:pos + span_len]
                    pos += span_len

                    if not original.strip():
                        continue
                    total += 1

                    if converted_chunk == original:
                        continue

                    rect = fitz.Rect(span["bbox"])
                    font_size = span["size"]
                    color_int = span["color"]
                    rgb = (((color_int >> 16) & 0xFF) / 255,
                           ((color_int >> 8) & 0xFF) / 255,
                           (color_int & 0xFF) / 255)

                    page.add_redact_annot(
                        rect,
                        text=converted_chunk,
                        fontsize=0,
                        fontname="china-t",
                        text_color=rgb,
                        align=fitz.TEXT_ALIGN_LEFT,
                    )
                    replaced += 1

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 "
              f"(轉換 {replaced}/{total} 個文字區塊)")

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert_pdf_s2t_v6.py <PDF檔案路徑>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成！輸出檔案: {output}")
