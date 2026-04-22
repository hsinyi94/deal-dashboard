"""最終驗證：比對每頁的完整文字內容。"""
import fitz
from opencc import OpenCC
import re

cc = OpenCC('s2twp')

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

all_ok = True

for i in range(len(orig_doc)):
    orig_text = orig_doc[i].get_text()
    conv_text = conv_doc[i].get_text()

    # 去掉空白和換行做比較
    orig_clean = re.sub(r'\s+', '', orig_text)
    conv_clean = re.sub(r'\s+', '', conv_text)

    # 原始文字的預期繁體版本
    expected_clean = re.sub(r'\s+', '', cc.convert(orig_text))

    # 檢查轉換後的文字是否包含預期的所有字元
    # （順序可能不同，因為 insert_text 的 z-order 不同）
    orig_chars = sorted(orig_clean)
    conv_chars = sorted(conv_clean)
    expected_chars = sorted(expected_clean)

    # 計算字元覆蓋率
    expected_set = {}
    for ch in expected_clean:
        expected_set[ch] = expected_set.get(ch, 0) + 1

    conv_set = {}
    for ch in conv_clean:
        conv_set[ch] = conv_set.get(ch, 0) + 1

    # 找出預期有但轉換後缺少的字元
    missing = {}
    for ch, count in expected_set.items():
        conv_count = conv_set.get(ch, 0)
        if conv_count < count:
            missing[ch] = count - conv_count

    # 找出轉換後多出的字元
    extra = {}
    for ch, count in conv_set.items():
        exp_count = expected_set.get(ch, 0)
        if count > exp_count:
            extra[ch] = count - exp_count

    total_missing = sum(missing.values())
    total_extra = sum(extra.values())

    status = "OK" if total_missing == 0 else "MISSING"
    print(f"第 {i+1:2d} 頁: 原始 {len(orig_clean)} 字, "
          f"轉換 {len(conv_clean)} 字, "
          f"預期 {len(expected_clean)} 字 "
          f"[{status}]", end="")

    if total_missing > 0:
        all_ok = False
        missing_str = ", ".join(f"'{ch}'x{n}" for ch, n in sorted(missing.items())[:10])
        print(f" 缺少 {total_missing} 字: {missing_str}", end="")
    if total_extra > 0:
        extra_str = ", ".join(f"'{ch}'x{n}" for ch, n in sorted(extra.items())[:5])
        print(f" 多出 {total_extra} 字: {extra_str}", end="")
    print()

if all_ok:
    print("\n✓ 所有頁面的文字內容完整，轉換正確！")
else:
    print("\n⚠ 部分頁面有缺少的字元")

orig_doc.close()
conv_doc.close()
