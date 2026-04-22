import win32com.client
import os

logo_path = os.path.join(os.getcwd(), "amazon_logo_white.png")
SC_LINK = "https://sellercentral-japan.amazon.com/ap/signin?clientContext=356-3263692-7875216&openid.pape.preferred_auth_policies=Policy15&openid.pape.max_auth_age=300&openid.return_to=https%3A%2F%2Fsellercentral-japan.amazon.com%2Fgc%2F3p-points%2Fpoints-deal-cn&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=sc_jp_amazon_com_v2&openid.mode=checkid_setup&intercept=false&language=zh_TW&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&pageId=sc_amazon_v3_unified&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0&ssoResponse=eyJ6aXAiOiJERUYiLCJlbmMiOiJBMjU2R0NNIiwiYWxnIjoiQTI1NktXIn0.IjzzSL5zQBglUnx4bzb3nZzp5EgPEL1xSI8ycxG5ReFYytOrdnjIVA.6YtFVeKUczsvdOFY.wsNOL3Y9KPZV6iBX1NvVdYrlkh3MM96nV4cA4P_22ydtNW8NxAHV8vo-D4c_eSQCbZqWDsJG2aBd9DmvXbuAzrGLwjjiSJcsNzfSG9H64mzVOZNzfGhCrrfUS0WVkZyUXbdLoiVHR2yZTlyG2HZiflEc-egBnoruC1njSMy8dwN-SkGKOwyw225zMFvtydCoK2w8Tjf8FQ.T5ZSGHXyrfxEINKYxAtJvg"

QUALTRICS = "https://amazonexteu.qualtrics.com/jfe/form/SV_e8OXHh3llHfCPno"

html = f'''<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="margin:0;padding:0;background-color:#F5F3EF;font-family:Calibri,'Microsoft JhengHei',Arial,sans-serif;color:#161D26;font-size:15px;line-height:1.7;">
<table width="640" cellpadding="0" cellspacing="0" border="0" align="center" style="background-color:#FFFFFF;">

<!-- HEADER -->
<tr><td style="background-color:#161D26;padding:32px 40px;text-align:center;">
  <img src="cid:amazonlogo" alt="amazon" width="200" style="display:block;margin:0 auto;">
  <p style="color:#FFFFFF;font-size:20px;font-weight:700;margin:20px 0 0;letter-spacing:0.5px;font-family:Calibri,'Microsoft JhengHei',Arial,sans-serif;">
    &#12304;&#20126;&#39340;&#36956;&#26085;&#26412;&#31449;&#12305;2026&#24180;4&#26376;&#26376;&#20419;&#31309;&#20998;&#20419;&#37559;<br>+5&#26376;&#31309;&#20998;&#29378;&#27489;&#31680; &#20419;&#37559;&#27963;&#21205;&#22577;&#21517;&#36890;&#30693;
  </p>
</td></tr>
<!-- ORANGE BAR -->
<tr><td style="height:4px;background-color:#FF6200;font-size:0;line-height:0;">&nbsp;</td></tr>

<!-- BODY -->
<tr><td style="padding:36px 40px;">

  <p style="font-size:16px;font-weight:700;margin:0 0 16px;">親愛的賣家您好，</p>
  <p style="margin:0 0 16px;">感謝您持續在亞馬遜日本站經營！我們誠摯邀請您參加即將到來的兩場促銷活動，把握機會提升商品曝光與銷售表現。</p>

  <!-- Applicable / Fee -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:20px;">
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#161D26;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">適用對象</span>
    &nbsp;亞馬遜賣家（符合積分促銷條件的 ASIN 將獲得 Points DEAL 紅色徽章）
  </td></tr>
  <tr><td style="padding:6px 0;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:13px;font-weight:700;padding:4px 12px;border-radius:4px;">參加費用</span>
    &nbsp;免費
  </td></tr>
  </table>

  <!-- EVENT 1 -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FAFAF8;">
  <tr><td style="padding:20px 24px;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;margin-bottom:10px;">活動 ①</span>
    <p style="font-size:16px;font-weight:700;margin:8px 0;color:#161D26;">2026年亞馬遜 4月月促</p>
    <p style="font-size:14px;color:#161D26;margin:0;">&#128197; 活動期間：2026年4月30日（週四）9:00 ～ 2026年5月3日（週日）23:59</p>
  </td></tr>
  </table>

  <!-- EVENT 2 -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FAFAF8;">
  <tr><td style="padding:20px 24px;">
    <span style="display:inline-block;background-color:#FF6200;color:#FFFFFF;font-size:12px;font-weight:700;padding:3px 10px;border-radius:4px;margin-bottom:10px;">活動 ②</span>
    <p style="font-size:16px;font-weight:700;margin:8px 0;color:#161D26;">2026年亞馬遜積分狂歡節 #5</p>
    <p style="font-size:14px;color:#161D26;margin:0;">&#128197; 活動期間：2026年5月8日（週五）9:00 ～ 2026年5月17日（週日）23:59</p>
  </td></tr>
  </table>

  <!-- DEADLINE -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:24px 0;">
  <tr><td style="background-color:#FFF4EC;border-left:4px solid #FF6200;padding:14px 20px;font-weight:700;font-size:15px;">
    ⏰ 報名截止：2026年4月19日（週日）23:59
  </td></tr>
  </table>

  <!-- NOTES -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#F5F3EF;border-radius:8px;margin:24px 0;">
  <tr><td style="padding:20px 24px;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#161D26;">&#128203; 提報注意事項</p>
    <ul style="margin:0;padding-left:20px;">
      <li style="margin-bottom:8px;font-size:14px;">賣家可選擇以下任一方案參加：<br>① 僅參加 4月月促<br>② 僅參加 5月積分狂歡節<br>③ 兩場活動皆參加</li>
      <li style="margin-bottom:8px;font-size:14px;">請為所有 ASIN 提交<b>同一個時間段</b>。若出現時間段混合的情況，提報將不予通過。</li>
      <li style="margin-bottom:8px;font-size:14px;">若需要為不同商品設定不同的參與時間，請透過<b>賣家平台</b>提交。</li>
    </ul>
  </td></tr>
  </table>

  <!-- WHAT IS POINTS DEAL -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#127991;&#65039; 什麼是亞馬遜「Points DEAL」積分促銷？</p>
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;background-color:#FAFAF8;margin:16px 0;">
  <tr><td style="padding:20px 24px;font-size:14px;line-height:1.7;">
    <p style="margin:0 0 12px;">・亞馬遜「Points DEAL」積分促銷（以下簡稱「Points DEAL」）是一種新型的促銷模式，賣家可透過設定亞馬遜積分來參與其中。</p>
    <p style="margin:0 0 12px;">・在 Points DEAL 出現之前，只有打折的商品才能參與亞馬遜的促銷活動。因此，對於多通路銷售（如實體店或其他網站）的賣家來說，參與促銷活動存在一定困難。</p>
    <p style="margin:0;">・如今，亞馬遜允許賣家在不打折的情況下也可以參與銷售活動。賣家只需提供更高的積分，就可以在搜尋結果頁面和商品詳情頁上顯示「ポイント DEAL」標識。</p>
  </td></tr>
  </table>

  <!-- REGISTRATION METHODS -->
  <p style="font-size:16px;font-weight:700;color:#161D26;margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid #FF6200;display:inline-block;">&#128221; 報名方式</p>

  <!-- Method 1 -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FFFFFF;">
  <tr><td style="padding:20px 24px;font-size:14px;line-height:1.7;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#FF6200;">方式一：透過賣家平台提交</p>
    <p style="margin:0 0 10px;">現在可以透過賣家平台批量建立促銷（同時最多 20,000 個 SKU）。透過賣家平台提交，您可以查看積分促銷徽章審核結果、對已設定的促銷進行修改和取消，根據銷售情況靈活調整。</p>
    <p style="margin:0 0 10px;">請進入 <a href="{SC_LINK}" style="color:#FF6200;text-decoration:none;font-weight:700;">亞馬遜積分促銷 Portal</a> 了解更多詳情。</p>
    <p style="font-size:13px;color:#666;margin:0 0 10px;">＊可以設定活動期間以外的時間<br>＊提交期間：可隨時在賣家平台上提交。建議在活動開始前一天完成提交</p>
    <p style="margin:0;">
      &#128073; <a href="{SC_LINK}" style="color:#FF6200;text-decoration:none;font-weight:700;">點擊此處從賣家平台設定 Points DEAL</a><br>
      &#128073; <a href="{SC_LINK}" style="color:#FF6200;text-decoration:none;font-weight:700;">點擊此處了解如何設定積分促銷</a><br>
      &#128073; <a href="{SC_LINK}" style="color:#FF6200;text-decoration:none;font-weight:700;">點擊此處查看設定教學影片</a>
    </p>
  </td></tr>
  </table>

  <!-- Method 2 -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #E8E6E1;border-radius:8px;margin-bottom:16px;background-color:#FFFFFF;">
  <tr><td style="padding:20px 24px;font-size:14px;line-height:1.7;">
    <p style="font-size:15px;font-weight:700;margin:0 0 12px;color:#FF6200;">方式二：透過亞馬遜 Qualtrics 系統提交（提交後無法修改／取消）</p>
    <p style="margin:0 0 10px;"><b>【申請方法】</b><br>點擊 Qualtrics 連結，閱讀活動說明後，下載提報表格，填寫後提交。</p>
  </td></tr>
  </table>

  <!-- CTA -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:28px 0;">
  <tr><td align="center">
    <a href="{QUALTRICS}" style="display:inline-block;background-color:#FF6200;color:#FFFFFF;text-decoration:none;font-size:15px;font-weight:700;padding:12px 36px;border-radius:6px;">立即提報（Qualtrics）→</a>
  </td></tr>
  </table>

  <p style="margin:0 0 16px;">如有任何問題，歡迎隨時與我們聯繫。期待您的參與！</p>
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
mail.Subject = "【亞馬遜日本站】2026年4月月促積分促銷+5月積分狂歡節 促銷活動報名通知- 截止日期 4/19（日）23:59"
mail.HTMLBody = html

att = mail.Attachments.Add(logo_path, 1, 0)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "amazonlogo"
)
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
    True
)

msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)
print(f"Done: {msg_path}")
