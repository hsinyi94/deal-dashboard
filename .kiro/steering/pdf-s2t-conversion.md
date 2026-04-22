---
inclusion: fileMatch
fileMatchPattern: 'convert_pdf_s2t*|convert_mde*|*_s2t*.py'
---

# PDF 簡繁轉換品質標準

本專案經常需要將簡體中文 PDF 轉換為繁體中文（台灣慣用詞彙），以下是所有轉換工作 **必須** 遵循的檢查清單與已知的技術陷阱。

## 核心工具

- **PDF 讀寫**：PyMuPDF (`fitz`)
- **簡繁轉換**：`opencc`，配置固定用 `s2twp`（含台灣慣用詞）
- **CJK 內建字型**：`china-t`（繁體中文）
- **ASCII 字型**：`helv`（Helvetica，半形寬度）
- **輸出命名規則**：原檔名加 `_繁體` 後綴

## 必做的品質檢查（每次轉換後都要驗證）

### 1. 簡體殘留檢查
- 原版 PDF 常用多種字型拼接一行文字（如 MeiryoUI + MicrosoftJhengHei + Calibri）
- 某些簡體字（如「页」「发」「们」）常被 Calibri 等拉丁字型以 fallback 方式渲染
- 轉換後必須做全文搜尋，確認下列常見簡體字已完全替換：
  `页 与 为 从 关 发 对 应 当 时 来 没 这 进 还 门 间 们 让 说 请 问 给 过`

### 2. 文字重疊檢查
- 逐行檢查相鄰 span 的 bbox：`r1.x1 - r2.x0 > 2px` 視為異常重疊
- 常見重疊發生在：
  - CJK 字元接 ASCII（例：`B2B瀏覽` 的 B 和 瀏 重疊）
  - 原版用 Calibri 渲染的 CJK 字（例：「页」）和相鄰字重疊
  - 同一行多 span 切換字型處

### 3. 字體大小一致性
- 同一視覺上的文字塊，字體大小必須一致
- **禁止**逐 span 按測量寬度縮放 fontsize（會造成同行大小不均）
- 只在整組（同 fontsize 的連續 span）超出 bbox > 2% 時才整組縮放

### 4. 文字溢出框架檢查
- 文字不能超出原版的彩色背景框、表格儲存格邊界
- 若轉換後寬度超出 bbox，整組按比例縮小，而非任意位置截斷

### 5. 殘留覆蓋文字檢查
- **最容易被忽略**：apply_redactions 後可能仍殘留底層的簡體文字
- 插入新文字前必須確認原文字已完全清除
- 用 `page.add_redact_annot(bbox)` + `page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)`
- Redact bbox 必須涵蓋整個 span 的完整範圍，不可只取文字寬度

### 6. 超連結保留
- 轉換會清除原版文字，但也會連帶遺失超連結
- **必做**：從原版 PDF 讀取所有 `page.get_links()`，用 `tgt_page.insert_link(link)` 複製到新檔
- 用 `tgt_doc.save(target_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)` 增量保存
- 驗證：比對原版和繁體版的總連結數應相同

### 7. 圖形與圖片保留
- `apply_redactions` 時必須用 `images=fitz.PDF_REDACT_IMAGE_NONE` 避免破壞圖片
- 不要對圖片區域加 redact annotation

## 已知的技術陷阱

### 陷阱 1：ASCII 字元寬度差異
- 原版日文字型（MeiryoUI）中的 ASCII 是半形（char_w ≈ 0.5 × fontsize）
- `china-t` 中的 ASCII 是全形（char_w = fontsize），會撐爆原始 bbox
- **解法**：每個 span 內必須拆成 ASCII 段和 CJK 段分別插入，ASCII 用 `helv`、CJK 用 `china-t`

### 陷阱 2：逐 span origin 插入的累積誤差
- 原版每個 span 的 origin 基於原始字型的字元寬度計算
- 若相鄰 span 用不同字型，其 origin 間距不等於 `china-t` 下的寬度
- **解法**：以「相同 fontsize 的連續 span 群組」為單位，從群組 bbox.x0 開始，位置由累加計算

### 陷阱 3：Calibri 渲染的簡體字
- Calibri 不支援中文，但 PDF 可能強制使用，導致「页」等字出現異常 bbox
- 簡繁轉換後這些字會變成「頁」，但位置和寬度可能異常
- **解法**：按「文字內容」而非「原始字型」決定使用的目標字型

### 陷阱 4：檔案鎖定
- 若 PDF 被 Adobe Reader 等開啟，`doc.save()` 會失敗
- **解法**：先存到 `output_path + ".tmp"`，再用 `shutil.move` 覆蓋，失敗時退回到 `_v{N}` 後綴

### 陷阱 5：Windows 控制台 Unicode
- Python 在 Windows CMD/PowerShell 列印中文可能報 `UnicodeEncodeError`
- **解法**：腳本開頭加 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')`

## 標準轉換流程

```python
# 1. 讀取原版 PDF
doc = fitz.open(input_path)

# 2. 對每頁：提取 text_dict，按「line → 相同 fontsize 的連續 span」分組

# 3. 對每個群組：
#    - 合併文字，用 OpenCC('s2twp') 轉換
#    - 加 redact annotation（用群組的合併 bbox）

# 4. apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

# 5. 重新插入文字：
#    - 從群組 bbox.x0 開始
#    - 分 ASCII/CJK 段，ASCII 用 helv、CJK 用 china-t
#    - fontsize 保持原始，僅在寬度超出 bbox > 2% 時整組縮放
#    - 位置由各段實測寬度累加

# 6. 存檔（先 tmp 再 move，避免鎖定）

# 7. 複製原版所有超連結到新檔

# 8. 執行品質檢查：簡體殘留、文字重疊、連結數比對
```

## 品質驗證腳本

每次轉換完成後，必須執行以下檢查並報告結果：

```python
def quality_check(original_path, converted_path):
    """標準品質檢查，回傳問題清單"""
    issues = []
    src = fitz.open(original_path)
    tgt = fitz.open(converted_path)

    # 1. 簡體殘留
    simp_chars = "页与为从关发对应当时来没这进还门间们让说请问给过"
    for pn in range(len(tgt)):
        text = tgt[pn].get_text()
        for ch in simp_chars:
            if ch in text:
                issues.append(f"第 {pn+1} 頁殘留簡體字 '{ch}'")
                break

    # 2. 文字重疊（> 2px）
    for pn in range(len(tgt)):
        text_dict = tgt[pn].get_text("dict")
        for block in text_dict["blocks"]:
            if block["type"] != 0: continue
            for line in block["lines"]:
                spans = sorted(line["spans"], key=lambda s: s["bbox"][0])
                for i in range(len(spans)-1):
                    r1 = fitz.Rect(spans[i]["bbox"])
                    r2 = fitz.Rect(spans[i+1]["bbox"])
                    if r1.y0 > r2.y1 or r2.y0 > r1.y1: continue
                    if r1.x1 - r2.x0 > 2.0:
                        issues.append(f"第 {pn+1} 頁文字重疊: '{spans[i]['text']}' / '{spans[i+1]['text']}'")

    # 3. 連結數比對
    src_links = sum(len(src[i].get_links()) for i in range(len(src)))
    tgt_links = sum(len(tgt[i].get_links()) for i in range(len(tgt)))
    if src_links != tgt_links:
        issues.append(f"連結數不符：原版 {src_links}，繁體版 {tgt_links}")

    src.close(); tgt.close()
    return issues
```
