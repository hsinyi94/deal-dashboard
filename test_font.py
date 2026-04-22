"""測試 china-t 字型是否能渲染所有需要的繁體字元。"""
import fitz
from opencc import OpenCC

cc = OpenCC('s2twp')

# 建立測試 PDF
doc = fitz.open()
page = doc.new_page(width=800, height=1200)

# 從原始 PDF 收集所有需要轉換的文字
orig = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
all_converted = set()
for p in orig:
    text = p.get_text()
    converted = cc.convert(text)
    for ch in converted:
        if ord(ch) > 127:  # 非 ASCII
            all_converted.add(ch)
orig.close()

# 嘗試用 china-t 渲染每個字元
test_text = "".join(sorted(all_converted))
print(f"需要渲染的繁體字元數: {len(test_text)}")

# 寫入測試 PDF
y = 50
line = ""
for ch in test_text:
    line += ch
    if len(line) >= 40:
        page.insert_text((50, y), line, fontsize=14, fontname="china-t")
        y += 20
        line = ""
        if y > 1150:
            page = doc.new_page(width=800, height=1200)
            y = 50

if line:
    page.insert_text((50, y), line, fontsize=14, fontname="china-t")

doc.save("test_font_output.pdf")
doc.close()

# 重新讀取測試 PDF，看哪些字元沒被渲染
test_doc = fitz.open("test_font_output.pdf")
rendered_text = ""
for p in test_doc:
    rendered_text += p.get_text()
test_doc.close()

rendered_chars = set(ch for ch in rendered_text if ord(ch) > 127)
missing_chars = all_converted - rendered_chars

if missing_chars:
    print(f"\n⚠ {len(missing_chars)} 個字元無法用 china-t 渲染:")
    for ch in sorted(missing_chars):
        print(f"  U+{ord(ch):04X} '{ch}'")
else:
    print("\n✓ 所有繁體字元都能用 china-t 正確渲染")
