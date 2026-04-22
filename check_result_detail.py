"""檢查轉換後 PDF 第 5 頁在消失位置附近的文字。"""
import fitz

doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")
page = doc[4]

# 檢查 y 座標在 270-300 範圍的所有 span（消失的「系」和「击此处」在這裡）
text_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)

print("第 5 頁所有 span（按 y 座標排序）:")
all_spans = []
for block in text_dict["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        for span in line["spans"]:
            t = span.get("text", "")
            if t.strip():
                all_spans.append({
                    "text": t,
                    "bbox": span["bbox"],
                    "font": span["font"],
                    "size": span["size"],
                })

all_spans.sort(key=lambda s: (s["bbox"][1], s["bbox"][0]))

for s in all_spans:
    y0 = s["bbox"][1]
    # 只看 y 在 270-400 的（消失文字的區域）
    if 270 <= y0 <= 400:
        print(f"  '{s['text']}' font={s['font']} size={s['size']:.1f} "
              f"bbox=({s['bbox'][0]:.1f}, {s['bbox'][1]:.1f}, {s['bbox'][2]:.1f}, {s['bbox'][3]:.1f})")

# 也用 get_text() 看整頁文字
print("\n整頁文字:")
print(page.get_text()[:1500])

doc.close()
