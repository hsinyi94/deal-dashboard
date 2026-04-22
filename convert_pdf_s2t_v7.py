"""
將簡體中文 PDF 轉換為繁體中文 (v7)。
核心改進：以「整行」為單位做 redact，避免重疊 bbox 導致相鄰 span 被清除。
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
        replaced_lines = 0
        total_lines = 0

        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                spans = line["spans"]
                if not spans:
                    continue

                # 合併整行文字
                full_text = "".join(s.get("text", "") for s in spans)
                if not full_text.strip():
                    continue
                total_lines += 1

                full_converted = cc.convert(full_text)
                if full_converted == full_text:
                    continue  # 整行都不需要轉換

                # 計算整行的 bounding box（所有 span 的聯集）
                line_rect = fitz.Rect(spans[0]["bbox"])
                for s in spans[1:]:
                    line_rect |= fitz.Rect(s["bbox"])

                # 取第一個有文字的 span 的顏色和大小作為預設
                color_int = spans[0]["color"]
                font_size = spans[0]["size"]
                for s in spans:
                    if s.get("text", "").strip():
                        color_int = s["color"]
                        font_size = s["size"]
                        break

                rgb = (((color_int >> 16) & 0xFF) / 255,
                       ((color_int >> 8) & 0xFF) / 255,
                       (color_int & 0xFF) / 255)

                page.add_redact_annot(
                    line_rect,
                    text=full_converted,
                    fontsize=0,  # auto-fit
                    fontname="china-t",
                    text_color=rgb,
                    align=fitz.TEXT_ALIGN_LEFT,
                )
                replaced_lines += 1

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 "
              f"(轉換 {replaced_lines}/{total_lines} 行)")

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert_pdf_s2t_v7.py <PDF檔案路徑>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成！輸出檔案: {output}")
