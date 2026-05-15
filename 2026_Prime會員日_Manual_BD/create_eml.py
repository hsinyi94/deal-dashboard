"""
Generate .eml file for 2026 Prime Day Manual BD promotion.
"""
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formatdate
from email.generator import Generator

CAMPAIGN_DIR = os.path.join(os.getcwd(), "2026_Prime會員日_Manual_BD")
EML_OUT = os.path.join(CAMPAIGN_DIR, "2026_Prime會員日_Manual_BD.eml")
LOGO_PATH = os.path.join(os.getcwd(), "amazon_logo_header.png")

SUBJECT = "【亞馬遜日本站】2026 Prime 會員日 Manual BD 招募 — 第一輪 5/23 截止"

MANUAL_BD_LINK = "https://suv.ma.globalsellingcommunity.cn/t/5RR5fa"
DEAL_DL_LINK = "https://sellercentral-japan.amazon.com/gc/jp-deal/deal-download-cn"
SC_BD_LINK = "https://sellercentral.amazon.co.jp/merchandising-new/landing"
FEE_HELP_LINK = "https://sellercentral.amazon.co.jp/help/hub/reference/external/G202111590?locale=en-US"

html_body = f'''<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="margin:0;padding:0;background-color:#F5F3EF;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;color:#161D26;font-size:15px;line-height:1.7;">
<table width="640" cellpadding="0" cellspacing="0" border="0" align="center" style="background-color:#FFFFFF;">

<!-- HEADER -->
<tr><td style="background-color:#161D26;padding:32px 40px;text-align:center;">
  <img src="cid:amazonlogo" alt="amazon" width="240" style="display:block;margin:0 auto;">
  <p style="color:#FFFFFF;font-size:20px;font-weight:700;margin:20px 0 0;letter-spacing:0.5px;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;">
    【亞馬遜日本站】2026 Prime 會員日<br>Manual BD 招募通知
  </p>
</td></tr>
<!-- ORANGE BAR -->
<tr><td style="height:4px;background-color:#FF6200;font-size:0;line-height:0;">&nbsp;</td></tr>

<!-- BODY -->
<tr><td style="padding:36px 40px;">

  <p style="font-size:16px;font-weight:700;margin:0 0 16px;">親愛的賣家您好，</p>
  <p style="margin:0 0 16px;">感謝您一直以來的支持。現正式啟動 <b>2026 年 Amazon Prime 會員日促銷活動手動申報 BD（Manual BD）</b>的商品招募，請把握第一輪報名的費用折扣！</p>

  <!-- IMPORTANT NOTICE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF4EC;border-left:4px solid #FF6200;margin:16px 0;">
  <tr><td style="padding:16px 20px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 8px;color:#FF6200;">&#9888;&#65039; 重要：手動申報 BD 遷移至新系統</p>
    <p style="margin:0;font-size:14px;line-height:1.6;">隨著手動申報 BD 和賣家平台自主申報的 Z 劃算（BD）系統的整合，兩種促銷活動將擁有<b>統一的功能和費用結構</b>。新系統支援透過賣家平台 → 秒殺頁面直接監控促銷報錯和調整價格，協助您最大限度地減少潛在的銷售損失。手動申報的 BD 提交將繼續由您的客戶經理提供支援。</p>
  </td></tr>
  </table>

  <!-- Applicable / Fee -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0;">
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#161D26;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">適用對象</span>
    &nbsp;亞馬遜日本站賣家（Manual BD 手動申報）
  </td></tr>
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">參加條件</span>
    &nbsp;設定至少 <b>1% 積分</b>
  </td></tr>
  </table>

  <!-- EVENT -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FAFAF8;">
  <tr><td style="padding:20px 24px;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;margin-bottom:10px;">活動期間</span>
    <p style="font-size:16px;font-weight:700;margin:8px 0;color:#161D26;">2026 Amazon Prime 會員日</p>
    <p style="font-size:14px;color:#161D26;margin:0;">&#128197; 活動時間：請關注後續來自亞馬遜官方的通知</p>
  </td></tr>
  </table>

  <!-- DEADLINE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF4EC;border-left:4px solid #FF6200;margin:24px 0;">
  <tr><td style="padding:14px 20px;font-size:15px;font-weight:700;">
    <p style="margin:0 0 8px;">&#9200; 第一輪報名截止：<span style="color:#FF6200;">2026 年 5 月 23 日（週五）18:00</span>（有費用折扣！）</p>
    <p style="margin:0 0 8px;">&#9200; 第二輪報名截止：2026 年 5 月 30 日（週五）18:00</p>
    <p style="margin:0;">&#9200; 第三輪報名截止：2026 年 6 月 6 日（週五）18:00</p>
  </td></tr>
  </table>

  <!-- POINTS -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:24px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#127919; 積分設定期間</p>
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">請等待後續通知。</li>
      <li style="margin-bottom:8px;font-size:14px;">務必於指定期間內完成至少 1% 的積分設定，否則無法參加活動。</li>
    </ul>
  </td></tr>
  </table>

  <!-- FEE STRUCTURE -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128176; 費用結構</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:12px 0 10px;font-size:14px;">
    <tr>
      <th style="width:45%;padding:10px 14px;border:1px solid #E8E6E1;background:#FAFAF8;font-weight:700;text-align:left;">項目</th>
      <th style="padding:10px 14px;border:1px solid #E8E6E1;background:#FAFAF8;font-weight:700;text-align:left;">金額</th>
    </tr>
    <tr>
      <td style="padding:10px 14px;border:1px solid #E8E6E1;">每個促銷預付費用</td>
      <td style="padding:10px 14px;border:1px solid #E8E6E1;color:#FF6200;font-weight:700;">&yen;600</td>
    </tr>
    <tr>
      <td style="padding:10px 14px;border:1px solid #E8E6E1;">每個促銷可變費用</td>
      <td style="padding:10px 14px;border:1px solid #E8E6E1;color:#FF6200;font-weight:700;">銷售額 1.0%（上限 &yen;20,000）</td>
    </tr>
  </table>
  <p style="font-size:13px;color:#6B7280;margin:0 0 20px;">參考：<a href="{FEE_HELP_LINK}" style="color:#FF6200;text-decoration:none;">Best Deal 費用詳情（賣家平台 Help）</a></p>

  <!-- FEE DISCOUNT -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#128161; 費用說明與折扣</p>
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;"><b>費用上限</b>：手動申報 BD 無論提交多少 ASIN，整個活動期間最多只收取 <b>10 個促銷活動</b>的費用。</li>
      <li style="margin-bottom:8px;font-size:14px;"><b>第一輪折扣</b>：若於第一輪（5/23 18:00 前）提交，無論 ASIN 數量，整個活動期間最多僅收取 <b>6 個促銷活動</b>的費用。</li>
      <li style="margin-bottom:8px;font-size:14px;"><b>第二輪之後</b>：最多收取 10 個促銷活動費用。</li>
      <li style="margin-bottom:8px;font-size:14px;">若第一輪提交後，在第二輪或之後追加提交，費用折扣仍然適用。</li>
      <li style="margin-bottom:8px;font-size:14px;">若同時在第一輪和第二輪（或之後）都有提交，但第一輪建立的所有 ASIN 都被取消了，則費用折扣將不適用。</li>
    </ul>
  </td></tr>
  </table>

  <!-- OTHER NOTES -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#128203; 其他注意事項</p>
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">原則上手動申報 BD 按<b>每個父 ASIN</b> 收費；系統遷移過渡期間，可能出現多個父 ASIN 合併為一個促銷的情況。</li>
      <li style="margin-bottom:8px;font-size:14px;">申報結果查詢：<br>1）透過賣家平台 → 廣告 → 秒殺 查看由亞馬遜建立的 Z 劃算；或者<br>2）透過賣家平台訪問 <a href="{DEAL_DL_LINK}" style="color:#FF6200;text-decoration:none;">Deal Download 頁面</a>（通常申報窗口關閉後一至兩週內更新）。</li>
    </ul>
  </td></tr>
  </table>

  <!-- CTA -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:28px 0 16px;">
  <tr><td align="center" style="text-align:center;">
    <a href="{MANUAL_BD_LINK}" style="display:inline-block;background-color:#FF6200;color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 36px;border-radius:6px;">點此前往 Manual BD 提報</a>
    <p style="margin:10px 0 0;font-size:13px;color:#6B7280;">（第一輪截止：5/23 18:00，把握最多 6 個促銷費用折扣）</p>
  </td></tr>
  </table>

  <!-- Self Registration -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128722; 想自行於賣家平台報名 Z 劃算（BD）？</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;"><b>參加費用</b>：每次 &yen;600 + 促銷銷售額 1%（上限 &yen;20,000）／每個父 ASIN</li>
      <li style="margin-bottom:8px;font-size:14px;"><b>報名截止時間</b>：2026 年 7 月 5 日</li>
      <li style="margin-bottom:8px;font-size:14px;"><b>報名入口</b>：<a href="{SC_BD_LINK}" style="color:#FF6200;text-decoration:none;">點擊此處報名 Z 劃算（BD）</a></li>
    </ul>
  </td></tr>
  </table>

  <p style="margin:28px 0 16px;">如有任何問題，歡迎隨時與我們聯繫。期待您的積極參與！</p>

  <!-- DASHBOARD CTA -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:28px 0 8px;">
  <tr><td align="center" style="text-align:center;background-color:#FFF4EC;border:1px dashed #FF6200;border-radius:10px;padding:22px 24px;">
    <p style="font-size:15px;font-weight:700;color:#161D26;margin:0 0 6px;">&#128640; 想一次掌握所有開放中與即將開放的促銷提報？</p>
    <p style="font-size:13px;color:#6B7280;margin:0 0 14px;">我們整理了完整的 Deal 提報儀表板，活動狀態、截止倒數、提報連結一目了然</p>
    <a href="https://hsinyi94.github.io/deal-dashboard/" style="display:inline-block;background-color:#FFFFFF;color:#FF6200;text-decoration:none;font-weight:700;font-size:14px;padding:10px 24px;border:2px solid #FF6200;border-radius:6px;">前往 Deal 提報儀表板 &rarr;</a>
  </td></tr>
  </table>

  <p style="margin-top:24px;">祝 生意興隆，<br>Amazon Global Selling Taiwan 團隊</p>

</td></tr>

<!-- FOOTER -->
<tr><td style="background-color:#161D26;padding:24px 40px;text-align:center;color:#FFFFFF;font-size:12px;line-height:1.8;">
  <p style="margin:0 0 4px;">Amazon Global Selling Taiwan</p>
  <p style="margin:0 0 4px;">23F, No. 100, Songren Rd., Xinyi Dist., Taipei City 110016, Taiwan (R.O.C.)</p>
  <p style="margin:0;">
    <a href="https://gs.amazon.com.tw/" style="color:#FF6200;text-decoration:none;">Website</a> &#65372;
    <a href="https://www.facebook.com/AmazonGlobalSellingTaiwan" style="color:#FF6200;text-decoration:none;">Facebook</a> &#65372;
    <a href="https://page.line.me/920bogra?openQrModal=true" style="color:#FF6200;text-decoration:none;">LINE</a> &#65372;
    <a href="https://www.youtube.com/@amazonglobalsellingtw" style="color:#FF6200;text-decoration:none;">YouTube</a> &#65372;
    <a href="https://gs.amazon.com.tw/learn" style="color:#FF6200;text-decoration:none;">Seller University</a>
  </p>
</td></tr>

</table>
</body></html>'''

# Build the .eml
msg = MIMEMultipart("related")
msg["Subject"] = SUBJECT
msg["From"] = "Amazon Global Selling Taiwan <noreply@amazon.com>"
msg["To"] = ""
msg["Date"] = formatdate(localtime=True)
msg["MIME-Version"] = "1.0"

# HTML part
html_part = MIMEText(html_body, "html", "utf-8")
msg.attach(html_part)

# Inline logo
with open(LOGO_PATH, "rb") as f:
    img = MIMEImage(f.read(), _subtype="png")
    img.add_header("Content-ID", "<amazonlogo>")
    img.add_header("Content-Disposition", "inline", filename="amazon_logo_header.png")
    msg.attach(img)

# Write .eml
os.makedirs(CAMPAIGN_DIR, exist_ok=True)
with open(EML_OUT, "w", encoding="utf-8") as f:
    gen = Generator(f)
    gen.flatten(msg)

print(f"Done: {EML_OUT}")
print(f"Size: {os.path.getsize(EML_OUT)} bytes")
