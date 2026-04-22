"""檢查消失的 span 是否和其他行的 redact 區域重疊。"""
import fitz
from opencc import OpenCC

cc = OpenCC('s2twp')
doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")

# 已知消失的 span（第 5 頁）
missing_bboxes = [
    ("系", (634.4, 274.6, 650.4, 293.4)),
    ("击此处", (803.4, 278.1, 833.3, 290.0)),
    ("击此处", (803.4, 312.3, 833.4, 324.2)),
    ("击此处", (803.4, 349.0, 833.3, 360.9)),
    ("击此处", (803.4, 385.7, 833.3, 397.5)),
]

page = doc[4]
text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

# 收集所有會被 redact 的 span
redact_rects = []
for block in text_dict["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        line_spans = line["spans"]
        full_text = "".join(s.get("text", "") for s in line_spans)
        full_converted = cc.convert(full_text)
        if full_converted == full_text:
            continue
        pos = 0
        for s in line_spans:
            original = s.get("text", "")
            span_len = len(original)
            converted_chunk = full_converted[pos:pos + span_len]
            pos += span_len
            if not original.strip():
                continue
            if converted_chunk != original:
                redact_rects.append({
                    "text": original,
                    "bbox": fitz.Rect(s["bbox"]),
                })

print(f"第 5 頁: {len(redact_rects)} 個 redact 區域")
print()

for text, bbox_tuple in missing_bboxes:
    missing_rect = fitz.Rect(bbox_tuple)
    print(f"消失的 span: '{text}' bbox={missing_rect}")
    found_overlap = False
    for rr in redact_rects:
        overlap = missing_rect & rr["bbox"]
        if not overlap.is_empty:
            print(f"  ⚠ 和 redact '{rr['text']}' bbox={rr['bbox']} 重疊")
            print(f"    重疊區域: {overlap}")
            found_overlap = True
    if not found_overlap:
        # 檢查這個 span 是否在需要轉換的行裡
        print(f"  ❌ 沒有和任何 redact 區域重疊！")
        # 找出這個 span 所在的行
        for block in text_dict["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for s in line["spans"]:
                    if s.get("text", "") == text:
                        sr = fitz.Rect(s["bbox"])
                        if sr.intersects(missing_rect):
                            full = "".join(sp.get("text", "") for sp in line["spans"])
                            conv = cc.convert(full)
                            print(f"  所在行: '{full}' → '{conv}'")
                            print(f"  行需轉換: {full != conv}")
    print()

doc.close()
