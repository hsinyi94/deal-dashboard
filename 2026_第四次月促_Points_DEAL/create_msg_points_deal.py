"""
Generate .msg for 2026 第四次月促（MDE#4）Points DEAL 積分促銷 招募通知

Structure:
- Pure table layout + inline CSS (no <style> block, no CSS classes)
- Microsoft JhengHei / Calibri / Arial font stack
- Logo via CID: Attachments.Add(path, 1, 0) + PR_ATTACH_CONTENT_ID
- Unicode subject via PR_SUBJECT_W
- SaveAs olMSGUnicode (9) for CJK safety
"""
import os
import win32com.client

logo_path = os.path.join(os.getcwd(), "amazon_logo_header.png")

CAMPAIGN_DIR = os.path.join(os.getcwd(), "2026_第四次月促_Points_DEAL")
MSG_OUT = os.path.join(CAMPAIGN_DIR, "2026_第四次月促_Points_DEAL.msg")

SUBJECT = "【亞馬遜日本站】2026 第四次月促（MDE#4）積分促銷 Points DEAL 招募 — 截止 5/19（二）23:59"

QUALTRICS_LINK = "https://amazonexteu.qualtrics.com/jfe/form/SV_e8OXHh3llHfCPno"
POINTS_DEAL_PORTAL = "https://sellercentral.amazon.co.jp/gc/3p-points/points-deal-cn"

PR_SUBJECT_W = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"

html = f'''<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="margin:0;padding:0;background-color:#F5F3EF;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;color:#161D26;font-size:15px;line-height:1.7;">
<table width="640" cellpadding="0" cellspacing="0" border="0" align="center" style="background-color:#FFFFFF;">

<!-- HEADER -->
<tr><td style="background-color:#161D26;padding:32px 40px;text-align:center;">
  <img src="cid:amazonlogo" alt="amazon" width="240" style="display:block;margin:0 auto;">
  <p style="color:#FFFFFF;font-size:20px;font-weight:700;margin:20px 0 0;letter-spacing:0.5px;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;">
    【亞馬遜日本站】2026 第四次月促（MDE#4）<br>積分促銷 Points DEAL 招募通知
  </p>
</td></tr>
<!-- ORANGE BAR -->
<tr><td style="height:4px;background-color:#FF6200;font-size:0;line-height:0;">&nbsp;</td></tr>

<!-- BODY -->
<tr><td style="padding:36px 40px;">

  <p style="font-size:16px;font-weight:700;margin:0 0 16px;">親愛的賣家您好，</p>
  <p style="margin:0 0 16px;">感謝您一直以來的支持。現正式啟動 <b>2026 年亞馬遜日本站第四次月促活動積分促銷（Points DEAL）</b>的商品招募，歡迎踴躍提報！</p>

  <!-- EVENT -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FAFAF8;">
  <tr><td style="padding:20px 24px;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;margin-bottom:10px;">活動期間</span>
    <p style="font-size:16px;font-weight:700;margin:8px 0;color:#161D26;">2026 第四次月促活動積分促銷（Points DEAL）</p>
    <p style="font-size:14px;color:#161D26;margin:0;">&#128197; 2026/5/27（週三）9:00 ～ 6/2（週二）23:59（JST）</p>
  </td></tr>
  </table>

  <!-- IMPORTANT NOTICE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF4EC;border-left:4px solid #FF6200;margin:16px 0;">
  <tr><td style="padding:16px 20px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 8px;color:#FF6200;">&#9888;&#65039; 注意事項</p>
    <p style="margin:0;font-size:14px;line-height:1.6;">與 Z 劃算相同，從第四次月促開始，活動頁面從 5/27 9:00（JST）開始，但<b>積分促銷徽章和積分從 5/27 0:00（JST）即開始生效</b>。</p>
  </td></tr>
  </table>

  <!-- DEADLINE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF4EC;border-left:4px solid #FF6200;margin:24px 0;">
  <tr><td style="padding:14px 20px;font-size:15px;font-weight:700;">
    <p style="margin:0 0 6px;">&#9200; 提報截止日期：<span style="color:#FF6200;">2026 年 5 月 19 日（週二）23:59</span></p>
    <p style="margin:0;font-size:13px;font-weight:400;color:#161D26;">錯過申報窗口不接受任何形式的補報。</p>
  </td></tr>
  </table>

  <!-- Applicable / Fee -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:20px 0;">
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#161D26;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">適用對象</span>
    &nbsp;亞馬遜日本站賣家（非時尚品類產品）
  </td></tr>
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">參加費用</span>
    &nbsp;免費
  </td></tr>
  </table>

  <!-- CONDITIONS -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128203; ASIN「Points Deal 積分促銷」提報條件</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">為 ASIN 設定 <b>5%～50%</b> 的積分</li>
      <li style="margin-bottom:8px;font-size:14px;">銷售價格不能比歷史價格（Was Price）高出一定範圍</li>
      <li style="margin-bottom:8px;font-size:14px;">活動期間價格 &minus; 活動積分的價格 低於 歷史價格（Was Price）&minus; 積分的價格</li>
      <li style="margin-bottom:8px;font-size:14px;">無歷史價格（Was Price）的 ASIN 也可參與。參與 Points Deal <b>不影響 Was Price</b></li>
      <li style="margin-bottom:8px;font-size:14px;">賣家評分為 0 或 &#9733;3.0 以上。評分數量少於 5 個的賣家，需保持最低 &#9733;2.0 的評分</li>
      <li style="margin-bottom:8px;font-size:14px;">商品評論為 0 或 &#9733;3.0 以上。評論數量少於 5 個的產品，需保持最低 &#9733;2.5 的評分</li>
      <li style="margin-bottom:8px;font-size:14px;">申請時有可售庫存</li>
      <li style="margin-bottom:8px;font-size:14px;">如為賣家自配送的商品，需向消費者免收配送費</li>
      <li style="margin-bottom:8px;font-size:14px;">以下產品<b>無法參與</b>本次活動：電子菸、成人用品、藥品、二手商品、收藏品、翻新品</li>
    </ul>
  </td></tr>
  </table>

  <!-- SUBMISSION NOTES -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128221; 提報須知</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">本連結僅供<b>非時尚品類</b>產品提報使用。</li>
      <li style="margin-bottom:8px;font-size:14px;">此申報連結不能保證貴司的申請一定會被通過，未通過條件審查的商品將透過郵件通知。</li>
      <li style="margin-bottom:8px;font-size:14px;">亞馬遜保留最終解釋權，可在不通知您的情況下對活動進行變更或取消，且預計開跑時間可能因狀況更改。</li>
      <li style="margin-bottom:8px;font-size:14px;">請賣家對自己的提報負責，每次重複提交的 ASIN 只會選取<b>最後一次提交</b>的 ASIN 為準。</li>
      <li style="margin-bottom:8px;font-size:14px;">即使提報後得到通知通過審查，仍然存在活動開始後不符合條件，從而喪失活動參與資格或 Points Deal 標誌顯示資格的情況。敬請注意。</li>
    </ul>
  </td></tr>
  </table>

  <!-- MERCHANT TOKEN -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF4EC;border-left:4px solid #FF6200;margin:16px 0;">
  <tr><td style="padding:16px 20px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 8px;color:#FF6200;">&#128273; 商家記號填寫提醒</p>
    <p style="margin:0;font-size:14px;line-height:1.6;">您的商家記號為（AXXXXX），請不要留空格或漏填、錯填。商家記號為字母及數字組合的 <b>14 位記號</b>，非店鋪名稱，請勿錯填。</p>
  </td></tr>
  </table>

  <!-- CONFLICT WARNING -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#9888;&#65039; 促銷衝突提醒</p>
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">對於已經申報 Points Deal 的 ASIN，請勿在月促或積分狂歡節（Points Festival）活動期間設定 DOTD、秒殺（LD）、Z 劃算（BD）和優惠券。</li>
    </ul>
  </td></tr>
  </table>

  <!-- TEMPLATE NOTES -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128196; 模板填報注意事項</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">每次填報<b>必須下載該連結內新的模板</b>進行填報。使用舊的模板可能會導致本次申報直接失敗。</li>
      <li style="margin-bottom:8px;font-size:14px;">不要修改模板格式，否則將無法讀取資料。</li>
      <li style="margin-bottom:8px;font-size:14px;">確保上傳的 Excel 表格內包含所有需要提報的 ASIN list。</li>
      <li style="margin-bottom:8px;font-size:14px;">如果需要修改資訊請重新上傳 Excel 表格，並確認表格內包含所有需要提報的 ASIN list。</li>
      <li style="margin-bottom:8px;font-size:14px;">上傳前請注意儲存為 Excel 格式，否則將無法讀取資料。</li>
      <li style="margin-bottom:8px;font-size:14px;">建議您優先為<b>推薦 ASIN</b> 提報 Points DEAL。推薦 ASIN 清單是基於亞馬遜的歷史資料生成的，為這些 ASIN 提交 Points Deal 活動既可以確保活動期間的銷售，又可以降低活動後價格風險（Points Deal 不影響 ASIN 的 Was Price）。</li>
    </ul>
  </td></tr>
  </table>

  <!-- SELF-SERVICE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#128161; 靈活提報推薦</p>
    <p style="margin:0;font-size:14px;">需靈活提報的賣家，推薦使用 Points DEAL 自助服務。詳情請查看 <a href="{POINTS_DEAL_PORTAL}" style="color:#FF6200;text-decoration:none;font-weight:700;">Points DEAL 入口網站</a></p>
  </td></tr>
  </table>

  <!-- CTA -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:28px 0 16px;">
  <tr><td align="center" style="text-align:center;">
    <a href="{QUALTRICS_LINK}" style="display:inline-block;background-color:#FF6200;color:#FFFFFF;text-decoration:none;font-weight:700;font-size:15px;padding:14px 36px;border-radius:6px;">點此前往 Points DEAL 提報</a>
    <p style="margin:10px 0 0;font-size:13px;color:#6B7280;">（提報截止：5/19（週二）23:59，逾期不受理補報）</p>
  </td></tr>
  </table>

  <!-- PORTAL LINK -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:12px 0 24px;">
  <tr><td align="center" style="text-align:center;">
    <a href="{POINTS_DEAL_PORTAL}" style="color:#FF6200;font-size:14px;font-weight:700;text-decoration:none;">&#128279; Points DEAL 入口網站（更多活動資訊 &amp; 自助服務）</a>
  </td></tr>
  </table>

  <p style="margin:0 0 16px;">活動細節請參照邀請郵件附件活動介紹。如有任何問題，歡迎隨時與我們聯繫。期待您的積極參與！</p>

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

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = SUBJECT
try:
    mail.PropertyAccessor.SetProperty(PR_SUBJECT_W, SUBJECT)
except Exception as e:
    print(f"[warn] PR_SUBJECT_W: {e}")

mail.HTMLBody = html

# CID logo
att = mail.Attachments.Add(logo_path, 1, 0)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "amazonlogo"
)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
    True
)

import tempfile

os.makedirs(CAMPAIGN_DIR, exist_ok=True)

# Save to TEMP first (ASCII path) to avoid CJK path issues, then copy
temp_msg = os.path.join(tempfile.gettempdir(), "points_deal_mde4.msg")
mail.SaveAs(temp_msg, 3)  # olMsg format for max compatibility

import shutil
shutil.copy2(temp_msg, MSG_OUT)
os.remove(temp_msg)
print(f"Done: {MSG_OUT}")
