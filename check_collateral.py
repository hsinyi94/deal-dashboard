"""檢查消失的 span 是否被 redact 的 bbox 覆蓋（附帶損害）。"""
import fitz
from opencc import OpenCC

cc = OpenCC('s2twp')
doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")

# 第 5 頁
page = doc[4]
text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

# 收集所有需要 redact 的 span bbox
redact_rects = []
unchanged_spans = []

for block in text_dict["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        line_spans = line["spans"]
        full_text = "".join(s.get("text", "") for s in line_spans)
        full_converted = cc.convert(full_text)

        if full_converted == full_text:
            # 這行不需轉換，記錄其 span
            for s in line_spans:
                if s.get("text", "").strip():
                    unchanged_spans.append({
                        "text": s["text"],
                        "bbox": fitz.Rect(s["bbox"]),
                        "font": s["font"],
                    })
            continue

        # 需要轉換的行，記錄被 redact 的 span
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
                    "converted": converted_chunk,
                    "bbox": fitz.Rect(s["bbox"]),
                })
            else:
                # 在需轉換的行裡，但這個 span 本身沒變
                unchanged_spans.append({
                    "text": original,
                    "bbox": fitz.Rect(s["bbox"]),
                    "font": s["font"],
                })

print(f"第 5 頁: {len(redact_rects)} 個 redact 區域, {len(unchanged_spans)} 個未變更 span")
print()

# 檢查未變更 span 是否和 redact 區域重疊
for us in unchanged_spans:
    for rr in redact_rects:
        overlap = us["bbox"] & rr["bbox"]  # intersection
        if not overlap.is_empty:
            print(f"⚠ 重疊！未變更 span '{us['text']}' ({us['font']})")
            print(f"  span bbox: {us['bbox']}")
            print(f"  redact bbox: {rr['bbox']} ('{rr['text']}' → '{rr['converted']}')")
            print(f"  重疊區域: {overlap}")
            print()

doc.close()
