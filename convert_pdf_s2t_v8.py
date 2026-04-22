"""
將簡體中文 PDF 轉換為繁體中文 (v8)。
策略：
1. 收集所有需要轉換的 span 資訊（文字、位置、樣式）
2. 用 redact 清除需要轉換的 span 的原始文字（不填入替代文字）
3. apply_redactions 後，用 insert_text 在原始 origin 位置逐 span 插入繁體文字
關鍵：redact 只清除「需要轉換」的 span，不影響不需轉換的 span。
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

        # 第一步：收集所有行，合併文字做轉換，然後拆回 span 級別
        spans_to_replace = []

        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                line_spans = line["spans"]
                if not line_spans:
                    continue

                # 合併整行文字
                full_text = "".join(s.get("text", "") for s in line_spans)
                full_converted = cc.convert(full_text)

                if full_converted == full_text:
                    continue  # 整行不需轉換

                # 按 span 長度拆分轉換後的文字
                pos = 0
                for span in line_spans:
                    original = span.get("text", "")
                    span_len = len(original)
                    converted_chunk = full_converted[pos:pos + span_len]
                    pos += span_len

                    if not original.strip():
                        continue

                    # 即使這個 span 的文字沒變，如果整行需要轉換，
                    # 我們也需要處理它（因為 redact 可能影響到它）
                    if converted_chunk != original:
                        spans_to_replace.append({
                            "bbox": fitz.Rect(span["bbox"]),
                            "origin": fitz.Point(span["origin"]),
                            "size": span["size"],
                            "color": span["color"],
                            "text": converted_chunk,
                        })

        # 第二步：對每個需要替換的 span，加 redact annotation（不填文字）
        for sp in spans_to_replace:
            page.add_redact_annot(sp["bbox"])

        # 第三步：套用 redaction（清除原文，不影響圖片）
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # 第四步：在原始位置插入繁體文字
        for sp in spans_to_replace:
            color_int = sp["color"]
            rgb = (((color_int >> 16) & 0xFF) / 255,
                   ((color_int >> 8) & 0xFF) / 255,
                   (color_int & 0xFF) / 255)

            page.insert_text(
                sp["origin"],
                sp["text"],
                fontsize=sp["size"],
                fontname="china-t",
                color=rgb,
            )

        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 "
              f"(替換 {len(spans_to_replace)} 個文字區塊)")

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert_pdf_s2t_v8.py <PDF檔案路徑>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成！輸出檔案: {output}")
