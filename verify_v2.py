"""驗證 v2 轉換結果：比對原始與轉換後的文字。"""
import fitz
from opencc import OpenCC

cc = OpenCC('s2twp')

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

for i in range(min(3, len(orig_doc))):  # 檢查前 3 頁
    orig_text = orig_doc[i].get_text().strip()
    conv_text = conv_doc[i].get_text().strip()
    expected = cc.convert(orig_text)

    print(f"===== 第 {i+1} 頁 =====")
    print(f"原始 (前200字): {orig_text[:200]}")
    print(f"轉換 (前200字): {conv_text[:200]}")
    print(f"預期 (前200字): {expected[:200]}")
    print(f"轉換字數: {len(conv_text)}, 原始字數: {len(orig_text)}")
    print()

orig_doc.close()
conv_doc.close()
