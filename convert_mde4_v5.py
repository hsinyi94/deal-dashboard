# -*- coding: utf-8 -*-
"""
將 MDE4/2026_MDE4_3P_CN.pdf 轉換為繁體中文 v5。

核心策略（根本解法）：
- 以「整行」為單位處理
- 清除整行文字 (redact)
- 從行的左邊界 (bbox.x0) 開始，用統一的方式重新渲染：
  - CJK 字元用 china-t，每字寬 = fontsize
  - ASCII 字元用 helv，按實際測量寬度
- 用原始行 bbox 寬度來檢查是否需要縮放
- 這樣同一行字體大小一致、字間距均勻、無重疊
"""

import sys
import os
import io
import shutil

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import fitz
from opencc import OpenCC


def split_ascii_cjk(text):
    """將文字分成 (segment, is_ascii) 段落"""
    if not text:
        return []
    segments = []
    current = ""
    current_is_ascii = None
    for ch in text:
        is_ascii = ord(ch) < 128
        if current_is_ascii is None:
            current_is_ascii = is_ascii
        if is_ascii == current_is_ascii:
            current += ch
        else:
            segments.append((current, current_is_ascii))
            current = ch
            current_is_ascii = is_ascii
    if current:
        segments.append((current, current_is_ascii))
    return segments


def measure_text_width(text, font_size):
    """測量混合文字的總寬度：ASCII 用 helv，CJK 用 china-t"""
    total = 0.0
    for seg_text, is_ascii in split_ascii_cjk(text):
        fontname = "helv" if is_ascii else "china-t"
        total += fitz.get_text_length(seg_text, fontname=fontname, fontsize=font_size)
    return total


def insert_mixed_text(page, origin, text, font_size, rgb):
    """插入混合文字：ASCII 用 helv，CJK 用 china-t"""
    x = origin.x
    for seg_text, is_ascii in split_ascii_cjk(text):
        fontname = "helv" if is_ascii else "china-t"
        page.insert_text(
            fitz.Point(x, origin.y),
            seg_text,
            fontsize=font_size,
            fontname=fontname,
            color=rgb,
        )
        x += fitz.get_text_length(seg_text, fontname=fontname, fontsize=font_size)


def convert_pdf_s2t(input_path: str) -> str:
    cc = OpenCC('s2twp')

    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_繁體{ext}"

    doc = fitz.open(input_path)

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

        # 收集行資訊 — 按「相同 fontsize 的連續 span」分組
        # 這樣同字體大小的文字可以作為一組處理
        line_groups = []  # list of groups, each group is a list of spans with same size

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

                # 按相同 fontsize 分組
                groups = []  # [(spans, text, origin, size, color, bbox)]
                pos = 0
                current_group_spans = []
                current_text = ""
                current_size = None
                current_color = None

                for span in line_spans:
                    original = span.get("text", "")
                    span_len = len(original)
                    converted_chunk = full_converted[pos:pos + span_len]
                    pos += span_len

                    size = span["size"]
                    color = span["color"]

                    if current_size is None:
                        current_size = size
                        current_color = color

                    # 如果 fontsize 改變，開啟新群組
                    if abs(size - current_size) > 0.1:
                        if current_group_spans:
                            groups.append({
                                "spans": current_group_spans,
                                "text": current_text,
                                "size": current_size,
                                "color": current_color,
                            })
                        current_group_spans = []
                        current_text = ""
                        current_size = size
                        current_color = color

                    current_group_spans.append(span)
                    current_text += converted_chunk

                if current_group_spans:
                    groups.append({
                        "spans": current_group_spans,
                        "text": current_text,
                        "size": current_size,
                        "color": current_color,
                    })

                # 為每個群組計算合併 bbox 和 origin
                for g in groups:
                    spans = g["spans"]
                    bbox = fitz.Rect(spans[0]["bbox"])
                    for s in spans[1:]:
                        bbox |= fitz.Rect(s["bbox"])
                    g["bbox"] = bbox
                    g["origin"] = fitz.Point(spans[0]["origin"])
                    g["originals"] = [s.get("text", "") for s in spans]

                line_groups.extend(groups)

        # Redact 所有群組的 bbox
        for g in line_groups:
            if any(t.strip() for t in g["originals"]):
                page.add_redact_annot(g["bbox"])

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # 重新插入每個群組的文字
        for g in line_groups:
            text = g["text"]
            if not text.strip():
                continue

            color_int = g["color"]
            rgb = (((color_int >> 16) & 0xFF) / 255,
                   ((color_int >> 8) & 0xFF) / 255,
                   (color_int & 0xFF) / 255)

            font_size = g["size"]
            origin = g["origin"]
            bbox = g["bbox"]

            # 測量新字型下的寬度
            measured = measure_text_width(text, font_size)
            available = bbox.width

            # 如果寬度超出範圍太多，縮小字體
            if measured > available * 1.02 and available > 0:
                ratio = available / measured
                adjusted_size = font_size * max(ratio, 0.5)
            else:
                adjusted_size = font_size

            insert_mixed_text(page, origin, text, adjusted_size, rgb)

        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 ({len(line_groups)} 個群組)")

    temp_path = output_path + ".tmp"
    doc.save(temp_path, garbage=4, deflate=True)
    doc.close()

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
    except PermissionError:
        output_path = f"{base}_繁體_v5{ext}"
    shutil.move(temp_path, output_path)
    return output_path


def copy_links(source_path: str, target_path: str):
    src_doc = fitz.open(source_path)
    tgt_doc = fitz.open(target_path)
    total_links = 0
    for page_idx in range(min(len(src_doc), len(tgt_doc))):
        src_page = src_doc[page_idx]
        tgt_page = tgt_doc[page_idx]
        for link in src_page.get_links():
            tgt_page.insert_link(link)
            total_links += 1
    tgt_doc.save(target_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    tgt_doc.close()
    src_doc.close()
    print(f"\n已複製 {total_links} 個超連結")


if __name__ == "__main__":
    input_file = "MDE4/2026_MDE4_3P_CN.pdf"
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成！輸出檔案: {output}")

    print(f"\n開始複製超連結...")
    copy_links(input_file, output)
    print(f"全部完成！")
