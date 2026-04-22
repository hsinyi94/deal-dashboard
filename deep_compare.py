"""
深入比對原始與轉換後 PDF 的每個 span，找出消失的文字。
"""
import fitz
from opencc import OpenCC

cc = OpenCC('s2twp')

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

for page_idx in range(len(orig_doc)):
    orig_page = orig_doc[page_idx]
    conv_page = conv_doc[page_idx]

    # 取得所有 span
    orig_dict = orig_page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
    conv_dict = conv_page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

    orig_spans = []
    for block in orig_dict["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                t = span.get("text", "")
                if t.strip():
                    orig_spans.append({
                        "text": t,
                        "bbox": tuple(round(x, 1) for x in span["bbox"]),
                        "font": span["font"],
                        "size": span["size"],
                    })

    conv_spans = []
    for block in conv_dict["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                t = span.get("text", "")
                if t.strip():
                    conv_spans.append({
                        "text": t,
                        "bbox": tuple(round(x, 1) for x in span["bbox"]),
                        "font": span["font"],
                        "size": span["size"],
                    })

    # 找出原始有但轉換後消失的文字
    orig_texts = set()
    for s in orig_spans:
        orig_texts.add(s["text"].strip())

    conv_texts = set()
    for s in conv_spans:
        conv_texts.add(s["text"].strip())

    # 比較：原始文字（或其繁體版本）是否出現在轉換後
    missing = []
    for s in orig_spans:
        ot = s["text"].strip()
        expected = cc.convert(ot)
        # 檢查轉換後是否有這段文字（原文或繁體版）
        found = False
        for cs in conv_spans:
            ct = cs["text"].strip()
            if ct == expected or ct == ot:
                found = True
                break
            # 也檢查是否包含
            if expected in ct or ot in ct:
                found = True
                break
        if not found:
            missing.append({
                "original": ot,
                "expected": expected,
                "bbox": s["bbox"],
                "font": s["font"],
            })

    if missing:
        print(f"\n===== 第 {page_idx + 1} 頁: {len(missing)} 個文字區塊消失 =====")
        for m in missing[:15]:
            print(f"  原始: '{m['original']}' → 預期: '{m['expected']}'")
            print(f"    位置: {m['bbox']}, 字型: {m['font']}")
    else:
        print(f"第 {page_idx + 1} 頁: OK (無消失文字)")

    # 也統計 span 數量差異
    if len(conv_spans) < len(orig_spans) * 0.8:
        print(f"  ⚠ Span 數量: 原始 {len(orig_spans)} → 轉換 {len(conv_spans)}")

orig_doc.close()
conv_doc.close()
