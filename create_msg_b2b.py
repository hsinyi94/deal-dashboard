"""
Generate .msg for 2026 日亞企業購品牌專場活動 (普通展位)

Mirrors the exact structure of create_msg_final2.py (which is known to work):
- Pure table layout + inline CSS (no <style> block, no CSS classes)
- Calibri / Microsoft JhengHei / Arial font stack
- Logo via CID: Attachments.Add(path, 1, 0) + PR_ATTACH_CONTENT_ID
- Unicode subject via PR_SUBJECT_W
- SaveAs olMSGUnicode (9) for CJK safety
- PDF attachment staged to ASCII path, renamed via .DisplayName
"""
import os
import shutil
import tempfile
import win32com.client

logo_path = os.path.join(os.getcwd(), "amazon_logo_header.png")

CAMPAIGN_DIR = os.path.join(os.getcwd(), "2026_日亞企業購品牌專場_普通展位")
PDF_SRC = os.path.join(CAMPAIGN_DIR, "2026日亞企業購品牌專場活動介紹-普通展位.pdf")
PDF_DISPLAY_NAME = "2026日亞企業購品牌專場活動介紹-普通展位.pdf"
MSG_OUT = os.path.join(CAMPAIGN_DIR, "2026_日亞企業購品牌專場_普通展位.msg")

SUBJECT = "【亞馬遜日本站】2026 企業購品牌專場活動（普通展位）開跑通知 - 即日起至 2026/12/31"

SU_SINGLE = "https://sellercentral-japan.amazon.com/learn/courses?moduleId=53dd5575-3ed8-4f18-8d30-94ee94494d01&courseId=c62c045f-a7ed-4051-98a1-d1c407edfc89&refTag=su_course_accordion&modLanguage=Chinese"
SU_BULK = "https://sellercentral-japan.amazon.com/learn/courses?ref_=su_course_accordion&moduleId=e9e07f8a-031f-4de3-9831-46c31c3df372&courseId=c62c045f-a7ed-4051-98a1-d1c407edfc89&modLanguage=Chinese"

PR_SUBJECT_W = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"

html = f'''<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="margin:0;padding:0;background-color:#F5F3EF;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;color:#161D26;font-size:15px;line-height:1.7;">
<table width="640" cellpadding="0" cellspacing="0" border="0" align="center" style="background-color:#FFFFFF;">

<!-- HEADER -->
<tr><td style="background-color:#161D26;padding:32px 40px;text-align:center;">
  <img src="cid:amazonlogo" alt="amazon" width="240" style="display:block;margin:0 auto;">
  <p style="color:#FFFFFF;font-size:20px;font-weight:700;margin:20px 0 0;letter-spacing:0.5px;font-family:'Microsoft JhengHei','微軟正黑體',Calibri,Arial,sans-serif;">
    【亞馬遜日本站】2026 企業購品牌專場活動<br>普通展位開跑通知
  </p>
</td></tr>
<!-- ORANGE BAR -->
<tr><td style="height:4px;background-color:#FF6200;font-size:0;line-height:0;">&nbsp;</td></tr>

<!-- BODY -->
<tr><td style="padding:36px 40px;">

  <p style="font-size:16px;font-weight:700;margin:0 0 16px;">親愛的賣家您好，</p>
  <p style="margin:0 0 16px;">眾所期待的亞馬遜日本站<b>企業購品牌專場活動</b>即將啟動！我們為品牌賣家打造專屬活動頁面，並透過企業購首頁廣告位的形式為品牌賣家吸引更多企業買家流量，促進 B2B 端銷量。期待您的積極參與！</p>

  <!-- Applicable / Fee -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:20px;">
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#161D26;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">適用對象</span>
    &nbsp;亞馬遜日本站品牌賣家（需已開通企業購）
  </td></tr>
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">參加費用</span>
    &nbsp;免費
  </td></tr>
  </table>

  <!-- EVENT -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FAFAF8;">
  <tr><td style="padding:20px 24px;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;margin-bottom:10px;">活動期間</span>
    <p style="font-size:16px;font-weight:700;margin:8px 0;color:#161D26;">2026 日亞企業購品牌專場活動（普通展位）</p>
    <p style="font-size:14px;color:#161D26;margin:0;">&#128197; 活動期間：即日起 ～ 2026年12月31日（週四）23:59</p>
  </td></tr>
  </table>

  <!-- KEY NOTICE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:24px 0;">
  <tr><td style="background-color:#FFF4EC;border-left:4px solid #FF6200;padding:14px 20px;font-weight:700;font-size:15px;">
    &#9989; 無須另行提報：活動期間普通展位每日刷新，自動抓取符合條件的 ASIN。建議儘早於後台完成企業折扣設定，爭取儘早投放！
  </td></tr>
  </table>

  <!-- CONDITIONS -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128203; 普通展位參與條件</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:16px 0;">
  <tr><td style="padding:20px 24px;">
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;"><b>企業折扣</b>（單件企業價 或 數量折扣最大檔）須在「B2C 價格扣除積分」的基礎上<b>再 5% OFF 以上</b>。</li>
      <li style="margin-bottom:8px;font-size:14px;">活動期間普通展位<b>每日刷新</b>，系統自動抓取符合條件的 ASIN，無需另外提交表單。</li>
      <li style="margin-bottom:8px;font-size:14px;">建議儘早設定 5% 以上企業折扣，爭取最大曝光時間。</li>
      <li style="margin-bottom:8px;font-size:14px;">更多活動介紹請參考附件 PDF：<i>2026 日亞企業購品牌專場活動介紹－普通展位</i>。</li>
    </ul>
  </td></tr>
  </table>

  <!-- GUIDES -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128736;&#65039; 後台設定教學（賣家大學）</p>

  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FFFFFF;">
  <tr><td style="padding:20px 24px;font-size:14px;line-height:1.7;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#FF6200;">教學一：設定單個 ASIN 的企業價格／數量折扣</p>
    <p style="margin:0 0 10px;">適用於少量商品逐一設定，步驟清晰、適合初次上手。</p>
    <p style="margin:0;">&#128073; <a href="{SU_SINGLE}" style="color:#FF6200;text-decoration:none;font-weight:700;">點擊前往賣家大學：單個 ASIN 設定教學</a></p>
  </td></tr>
  </table>

  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FFFFFF;">
  <tr><td style="padding:20px 24px;font-size:14px;line-height:1.7;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#FF6200;">教學二：批量設定多個 ASIN 的企業價格／數量折扣</p>
    <p style="margin:0 0 10px;">適用於大量商品一次設定，可透過上傳模板提升效率。</p>
    <p style="margin:0;">&#128073; <a href="{SU_BULK}" style="color:#FF6200;text-decoration:none;font-weight:700;">點擊前往賣家大學：批量設定教學</a></p>
  </td></tr>
  </table>

  <p style="margin:28px 0 16px;">如有任何問題，歡迎隨時與我們聯繫。期待您的積極參與！</p>
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

# Stage PDF to ASCII path so Outlook COM Attachments.Add doesn't fail on CJK
tmpdir = tempfile.mkdtemp(prefix="msg_stage_")
pdf_stage = os.path.join(tmpdir, "b2b_campaign_guide.pdf")
if os.path.exists(PDF_SRC):
    shutil.copyfile(PDF_SRC, pdf_stage)
    pdf_ok = True
else:
    print(f"[warn] PDF not found: {PDF_SRC}")
    pdf_ok = False

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = SUBJECT
try:
    mail.PropertyAccessor.SetProperty(PR_SUBJECT_W, SUBJECT)
except Exception as e:
    print(f"[warn] PR_SUBJECT_W: {e}")

mail.HTMLBody = html

# CID logo — same call as create_msg_final2.py (known working)
att = mail.Attachments.Add(logo_path, 1, 0)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "amazonlogo"
)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
    True
)

# PDF with CJK display name
if pdf_ok:
    att_pdf = mail.Attachments.Add(pdf_stage)
    try:
        att_pdf.DisplayName = PDF_DISPLAY_NAME
    except Exception as e:
        print(f"[warn] DisplayName: {e}")

os.makedirs(os.path.dirname(MSG_OUT), exist_ok=True)
mail.SaveAs(MSG_OUT, 9)
print(f"Done: {MSG_OUT}")

shutil.rmtree(tmpdir, ignore_errors=True)
