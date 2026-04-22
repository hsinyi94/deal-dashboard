---
inclusion: manual
---

# 促銷活動 Email + LINE Blurb 產出 SOP

當使用者提供促銷活動資訊並要求產出 email 和 LINE blurb 時，按以下流程執行。

## 使用者需提供的資訊

以下為必要欄位，若缺少請主動詢問：

1. **活動名稱**（例：4月月促、積分狂歡節 #5）
2. **活動期間**（開始日期時間 ～ 結束日期時間，可多場）
3. **提報截止日期**
4. **提報連結**（Qualtrics 連結，若不需要提報可省略）
5. **賣家平台連結**（如適用）
6. **特殊注意事項**（如有）

以下為選填：
- Email 標題（若未提供則自動生成）
- 是否包含 Points DEAL 說明
- 額外的報名方式說明

## 產出流程

### Step 1：準備 Amazon logo（關鍵！務必先做）

**⚠️ 每個新活動都必須先執行 `build_logo.py` 產出正確的 logo。**

正確的 logo 顯示途徑（已驗證可用，未來不要再嘗試其他方法）：

1. **來源**：使用專案根目錄的 `amazon_logo.svg`（Wikimedia Commons 官方 Amazon wordmark，含 path8/10/12/14/16/18/30 + `<use>` 指向 path30）
2. **白化**：將所有 `fill` 屬性和 inline `fill:` style 改為 `#FFFFFF`，存成 `amazon_logo_white.svg`
3. **光柵化**：用 **Chrome headless 在 SVG 的 native intrinsic size (603×182) 渲染**，4× supersampling 取得高解析 PNG
   - ⚠️ **絕對不要** 把 SVG 包在 `display:flex` 容器 + 百分比寬度裡——Chrome 有 flex + percent-sized SVG 佈局 bug，會讓字母 m 被裁掉
   - ⚠️ **絕對不要** 嘗試 cairosvg / svglib / reportlab——Windows 上通常沒裝 libcairo runtime
4. **合成**：用 PIL 把 PNG 縮放到 `0.55 × 1024 = 563 px` 寬，居中貼到 1024×320 的 `#161D26` 深色背景 canvas 上
5. **輸出**：`amazon_logo_header.png`（根目錄 + 複製一份到活動資料夾供 HTML 相對引用）

**不要做的事：**
- ❌ 不要自己用 PIL 繪圖畫 amazon 字樣（無法複製 Ember 字體）
- ❌ 不要從 Wikimedia 線上直接抓 SVG（HTTP 403）
- ❌ 不要從 freebiesupply / pngmart 線上 PNG（可能是 smile icon 不是 wordmark）
- ❌ 不要用 WordEditor / InlineShapes.AddPicture 路線（舊版 Outlook 存檔後會 re-encode 成 image001.png 但可能失效）
- ❌ 不要用 base64 data URI 嵌在 `<img src="data:...">`（Outlook 2016+ 預設擋掉 data URI）

`build_logo.py` 的核心邏輯（已驗證可用）：

```python
# 1. 白化 SVG
svg = re.sub(r'fill="[^"]*"', 'fill="#FFFFFF"', svg)
svg = re.sub(r'fill:#[0-9A-Fa-f]{3,6}', 'fill:#FFFFFF', svg)

# 2. Chrome headless 用 native size 渲染（key insight: 不要用 flex）
cmd = [CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
       f"--window-size={SVG_W*4},{SVG_H*4}",
       f"--screenshot={png_path}",
       "--default-background-color=00000000",
       f"file:///{html_path}"]  # html 裡直接放 svg、不要 flex 容器

# 3. PIL 合成到 1024x320 canvas 上
canvas = Image.new("RGB", (1024, 320), (0x16, 0x1D, 0x26))
canvas.paste(logo, (x, y), logo)
```

### Step 2：產出 email_template.html
- 依照 `email-template-guide.md` 的品牌規範和 HTML 結構
- 使用 CSS class + `<style>` 區塊（瀏覽器預覽用）
- **字體：`'Microsoft JhengHei', '微軟正黑體', sans-serif`**（不要用 Noto Sans TC + Google Fonts，預設會失敗）
- Logo 用相對路徑 `<img src="amazon_logo_header.png">`，並複製一份 PNG 到活動資料夾

### Step 3：產出 .msg 檔案
- 將 HTML 轉為 table-based layout + 全 inline style
- **字體 stack：`'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif`**（JhengHei 擺最前）
- Amazon logo 用 **CID 嵌入**：`Attachments.Add(logo_path, 1, 0)` + `PR_ATTACH_CONTENT_ID (0x3712001F) = "amazonlogo"` + `PR_ATTACHMENT_HIDDEN (0x7FFE000B) = True`
- **Subject 務必用 Unicode property tag 強制**：`PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0037001F", SUBJECT)`，否則中文被丟成 `?`
- **PDF 附件**：先 `shutil.copyfile` 到 ASCII 暫存路徑再 `Attachments.Add`，然後用 `attachment.DisplayName = "中文檔名.pdf"` 設顯示名稱
- **SaveAs 用 `9` (olMSGUnicode)**，不是 `3` (olMSG ANSI)
- 透過 Outlook COM 產出 .msg
- 檔名格式：`{年}_{活動簡稱}.msg`，放在活動資料夾

### Step 4：產出 LINE_blurb.txt
- 依照 LINE blurb 格式規範
- 精簡聚焦：活動檔期、截止日、提報管道（含連結）
- 純文字，適合直接複製貼上

### Step 5：組織到資料夾
活動資料夾 `{年}_{活動簡稱}/` 裡放：
```
email_template.html
amazon_logo_header.png    ← HTML 預覽用（相對引用）
{年}_{活動簡稱}.msg       ← Outlook 可直接開啟
LINE_blurb.txt            ← 複製貼到 LINE
{活動介紹}.pdf            ← 若有附件，複製一份進來
```

### Step 6：驗收提示
- 列出已產出的檔案清單
- **提醒使用者務必用 Classic Outlook 開啟 .msg 校稿**
- 告知新版 Outlook（紫色圖示，OWA desktop 殼）預覽本地 .msg 時會剝離 CID attachment，導致 logo 不顯示——這是新版 Outlook 架構限制，不是 .msg 問題
- 實際寄送給收件者時，收件端 Outlook（含網頁版）都會正常顯示 logo

## Outlook 新舊版差異（收件 / 預覽行為）

| 情境 | Classic Outlook | 新版 Outlook |
|------|----------------|-------------|
| 打開本地 .msg 預覽 | ✅ CID logo 正常 | ❌ CID 被剝離 logo 破圖 |
| 收到寄出的郵件 | ✅ 正常 | ✅ 正常（因為實際是 MIME email，不是 .msg） |

**結論**：.msg 只用來預覽/校稿，**一律以 Classic Outlook 為準**。

## 範例輸入

```
活動：6月月促
期間：2026年6月15日（一）9:00 ～ 6月18日（四）23:59
截止：2026年6月5日（五）23:59
提報連結：https://amazonexteu.qualtrics.com/jfe/form/SV_xxxxx
```

## 範例輸出檔案

```
2026_6月月促/
├── email_template.html
├── amazon_logo_header.png
├── 2026_6月月促.msg
└── LINE_blurb.txt
```
