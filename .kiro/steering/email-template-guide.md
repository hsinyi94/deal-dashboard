---
inclusion: auto
---

# TWGS 促銷 Email & LINE Blurb 品牌規範

當使用者要求產出促銷活動的 email 或 LINE blurb 時，請嚴格遵守以下規範。

## 品牌色彩（TWGS Visual Identity）

| 名稱 | 色碼 | 用途 |
|------|------|------|
| Squid Ink | `#161D26` | Header/Footer 背景、主文字色 |
| Off-White | `#F5F3EF` | 頁面背景、Notes 區塊背景 |
| White | `#FFFFFF` | 內容區背景、Header 文字 |
| Smile Orange | `#FF6200` | CTA 按鈕、重點提示、accent bar、連結色（僅作點綴，不作大面積背景） |
| Deep Orange | `#F55600` | 備用強調色 |

## 字型

- 中文：Noto Sans TC（HTML）/ Microsoft JhengHei（Outlook fallback）/ Calibri
- 英文：Ember Modern Display Bold（標題）、Amazon Ember Regular（內文），fallback 用 Calibri, Arial
- 最小字級：14px

## Email HTML 結構

```
1. Header（深色背景 #161D26）
   - Amazon logo（白色版本）
   - 活動主標題（白色文字）
2. Orange accent bar（4px #FF6200）
3. Content 區
   - 問候語
   - 適用對象 / 參加費用 標籤
   - 活動卡片（圓角邊框、淺灰背景 #FAFAF8）
   - 報名截止（左側橘色邊框提示框）
   - 提報注意事項（Off-White 背景區塊）
   - Points DEAL 說明（如適用）
   - 報名方式（方式一：賣家平台 / 方式二：Qualtrics）
   - CTA 按鈕（橘色圓角）
   - 結尾問候
4. Footer（深色背景 #161D26）
   - Amazon Global Selling Taiwan
   - 地址
   - 五個連結：Website | Facebook | LINE | YouTube | Seller University
```

## Footer 固定連結

- Website: https://gs.amazon.com.tw/
- Facebook: https://www.facebook.com/AmazonGlobalSellingTaiwan
- LINE: https://page.line.me/920bogra?openQrModal=true
- YouTube: https://www.youtube.com/@amazonglobalsellingtw
- Seller University: https://gs.amazon.com.tw/learn

## .msg 檔案產出規則

- 使用 Outlook COM（pywin32）產出
- HTML 必須使用 table-based layout + 全 inline style（Outlook 不支援 `<style>` 區塊和 CSS class）
- Logo 使用 CID 嵌入方式（Attachment + PropertyAccessor 設定 Content-ID）
- 儲存格式：olMSG (3)

## LINE Blurb 格式

```
📢【亞馬遜日本站】{活動主題}提醒

━━━━━━━━━━━━━━━━━━

📌 活動檔期
{每個活動用 ① ② 編號，含日期時間}

━━━━━━━━━━━━━━━━━━

⏰ 提報截止
🔴 {截止日期時間}

━━━━━━━━━━━━━━━━━━

📝 提報管道
▸ 管道一：賣家平台 + 連結
▸ 管道二：Qualtrics 表單 + 連結

━━━━━━━━━━━━━━━━━━

💡 {注意事項摘要}

有任何問題歡迎隨時聯繫 🙌
```

## 語言規範

- 所有面向賣家的內容一律使用**繁體中文**
- 專有名詞保留英文：ASIN, SKU, Points DEAL, Qualtrics, Amazon Global Selling Taiwan
- 日文促銷名稱保留原文：如「ポイント DEAL」
