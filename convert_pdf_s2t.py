# -*- coding: utf-8 -*-
"""
PDF 簡體中文轉繁體中文（通用版）。

遵循 .kiro/steering/pdf-s2t-conversion.md 的標準流程：
- OpenCC 's2twp' 配置（含台灣慣用詞）
- 按「相同 fontsize 的連續 span」分組處理
- ASCII 用 helv、CJK 用 china-t 分段插入
- 必要時整組按比例縮小 fontsize
- 保留原版所有超連結
- 內建品質檢查（簡體殘留、文字重疊、連結數比對）

用法：
    python convert_pdf_s2t.py <PDF檔案路徑>
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

        line_groups = []

        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                line_spans = line["spans"]
                if not line_spans:
                    continue

                full_text = "".join(s.get("text", "") for s in line_spans)
                full_converted = cc.convert(full_text)

                # 按相同 fontsize 分組
                groups = []
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

                for g in groups:
                    spans = g["spans"]
                    bbox = fitz.Rect(spans[0]["bbox"])
                    for s in spans[1:]:
                        bbox |= fitz.Rect(s["bbox"])
                    g["bbox"] = bbox
                    g["origin"] = fitz.Point(spans[0]["origin"])
                    g["originals"] = [s.get("text", "") for s in spans]

                line_groups.extend(groups)

        # Redact 所有群組
        for g in line_groups:
            if any(t.strip() for t in g["originals"]):
                page.add_redact_annot(g["bbox"])

        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # 重新插入
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

            measured = measure_text_width(text, font_size)
            available = bbox.width

            if measured > available * 1.02 and available > 0:
                ratio = available / measured
                adjusted_size = font_size * max(ratio, 0.5)
            else:
                adjusted_size = font_size

            insert_mixed_text(page, origin, text, adjusted_size, rgb)

        print(f"  第 {page_idx + 1}/{len(doc)} 頁完成 ({len(line_groups)} 個群組)")

    # 先存 tmp 再 move，避免檔案鎖定
    temp_path = output_path + ".tmp"
    doc.save(temp_path, garbage=4, deflate=True)
    doc.close()

    try:
        if os.path.exists(output_path):
            os.remove(output_path)
    except PermissionError:
        i = 2
        while os.path.exists(f"{base}_繁體_v{i}{ext}"):
            i += 1
        output_path = f"{base}_繁體_v{i}{ext}"
    shutil.move(temp_path, output_path)
    return output_path


def copy_links(source_path: str, target_path: str) -> int:
    src_doc = fitz.open(source_path)
    tgt_doc = fitz.open(target_path)
    total_links = 0
    for page_idx in range(min(len(src_doc), len(tgt_doc))):
        for link in src_doc[page_idx].get_links():
            tgt_doc[page_idx].insert_link(link)
            total_links += 1
    tgt_doc.save(target_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
    tgt_doc.close()
    src_doc.close()
    return total_links


def quality_check(original_path, converted_path):
    """標準品質檢查，回傳問題清單"""
    issues = []
    src = fitz.open(original_path)
    tgt = fitz.open(converted_path)

    # 1. 簡體殘留
    simp_chars = "页与为从关发对应当时来没这进还门间们让说请问给过"
    simp_issues = set()
    for pn in range(len(tgt)):
        text = tgt[pn].get_text()
        for ch in simp_chars:
            if ch in text:
                simp_issues.add((pn + 1, ch))
    for pn, ch in sorted(simp_issues):
        issues.append(f"[簡體殘留] 第 {pn} 頁殘留 '{ch}'")

    # 2. 文字重疊（> 2px）
    overlap_count = 0
    for pn in range(len(tgt)):
        text_dict = tgt[pn].get_text("dict")
        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                spans = sorted(line["spans"], key=lambda s: s["bbox"][0])
                for i in range(len(spans) - 1):
                    r1 = fitz.Rect(spans[i]["bbox"])
                    r2 = fitz.Rect(spans[i + 1]["bbox"])
                    if r1.y0 > r2.y1 or r2.y0 > r1.y1:
                        continue
                    t1 = spans[i].get("text", "").strip()
                    t2 = spans[i + 1].get("text", "").strip()
                    if not t1 or not t2:
                        continue
                    overlap = r1.x1 - r2.x0
                    if overlap > 2.0:
                        overlap_count += 1
                        if overlap_count <= 10:
                            issues.append(
                                f"[文字重疊] 第 {pn+1} 頁: '{t1}' / '{t2}' 重疊 {overlap:.1f}px"
                            )
    if overlap_count > 10:
        issues.append(f"[文字重疊] 另有 {overlap_count - 10} 處未列出")

    # 3. 連結數比對
    src_links = sum(len(src[i].get_links()) for i in range(len(src)))
    tgt_links = sum(len(tgt[i].get_links()) for i in range(len(tgt)))
    if src_links != tgt_links:
        issues.append(f"[連結遺失] 原版 {src_links} 個，繁體版 {tgt_links} 個")

    src.close()
    tgt.close()
    return issues, {"src_links": src_links, "tgt_links": tgt_links,
                    "overlap_count": overlap_count, "simp_issues": len(simp_issues)}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python convert_pdf_s2t.py <PDF檔案路徑>")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        sys.exit(1)

    print(f"開始轉換: {input_file}")
    output = convert_pdf_s2t(input_file)
    print(f"\n轉換完成: {output}")

    print(f"\n複製超連結...")
    link_count = copy_links(input_file, output)
    print(f"已複製 {link_count} 個超連結")

    print(f"\n執行品質檢查...")
    issues, stats = quality_check(input_file, output)
    print(f"  簡體殘留: {stats['simp_issues']} 處")
    print(f"  文字重疊: {stats['overlap_count']} 處")
    print(f"  連結數: 原版 {stats['src_links']} / 繁體版 {stats['tgt_links']}")
    if issues:
        print(f"\n⚠️ 發現 {len(issues)} 個問題:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("\n✅ 所有品質檢查通過")
