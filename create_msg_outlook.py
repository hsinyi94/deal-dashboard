import win32com.client
import os

logo_path = os.path.join(os.getcwd(), "amazon_logo_white.png")

# Build fully inline-styled HTML (Outlook ignores <style> blocks in many cases)
html = '''<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body style="margin:0; padding:0; background-color:#F5F3EF; font-family:Calibri, Arial, sans-serif; color:#161D26; font-size:15px; line-height:1.7;">

<div style="max-width:640px; margin:0 auto; background-color:#FFFFFF;">

  <!-- Header -->
  <div style="background-color:#161D26; padding:32px 40px; text-align:center;">
    <img src="cid:amazonlogo" alt="amazon" width="200" style="display:block; margin:0 auto;">
    <h1 style="color:#FFFFFF; font-size:22px; font-weight:700; margin:20px 0 0; letter-spacing:0.5px; font-family:Calibri, Arial, sans-serif;">2026&#24180;&#20126;&#39340;&#36956; 4&#26376;&#26376;&#20419; + 5&#26376;&#31309;&#20998;&#29378;&#27489;&#31680;</h1>
  </div>
  <div style="height:4px; background-color:#FF6200;"></div>

  <!-- Content -->
  <div style="padding:36px 40px;">

    <p style="font-size:16px; font-weight:700; margin:0 0 16px;">&#35242;&#24859;&#30340;&#36067;&#23478;&#24744;&#22909;&#65292;</p>
    <p style="margin:0 0 16px;">&#24863;&#35613;&#24744;&#25345;&#32396;&#22312;&#20126;&#39340;&#36956;&#26085;&#26412;&#31449;&#32147;&#29151;&#65281;&#25105;&#20497;&#35488;&#25711;&#36992;&#35531;&#24744;&#21443;&#21152;&#21363;&#23559;&#21040;&#20358;&#30340;&#20841;&#22580;&#20419;&#37559;&#27963;&#21205;&#65292;&#25226;&#25569;&#27231;&#26371;&#25552;&#21319;&#21830;&#21697;&#26333;&#20809;&#33287;&#37559;&#21806;&#34920;&#29694;&#12290;</p>

    <!-- Event 1 -->
    <div style="border:1px solid #E8E6E1; border-radius:8px; padding:20px 24px; margin-bottom:16px; background-color:#FAFAF8;">
      <span style="display:inline-block; background-color:#FF6200; color:#FFFFFF; font-size:12px; font-weight:700; padding:3px 10px; border-radius:4px; margin-bottom:10px;">&#27963;&#21205; &#9312;</span>
      <h2 style="font-size:16px; font-weight:700; margin:0 0 8px; color:#161D26; font-family:Calibri, Arial, sans-serif;">2026&#24180;&#20126;&#39340;&#36956; 4&#26376;&#26376;&#20419;</h2>
      <p style="font-size:14px; color:#161D26; margin:0;">&#128197; &#27963;&#21205;&#26399;&#38291;&#65306;2026&#24180;4&#26376;30&#26085;&#65288;&#36913;&#22235;&#65289;9:00 &#65374; 2026&#24180;5&#26376;3&#26085;&#65288;&#36913;&#26085;&#65289;23:59</p>
    </div>

    <!-- Event 2 -->
    <div style="border:1px solid #E8E6E1; border-radius:8px; padding:20px 24px; margin-bottom:16px; background-color:#FAFAF8;">
      <span style="display:inline-block; background-color:#FF6200; color:#FFFFFF; font-size:12px; font-weight:700; padding:3px 10px; border-radius:4px; margin-bottom:10px;">&#27963;&#21205; &#9313;</span>
      <h2 style="font-size:16px; font-weight:700; margin:0 0 8px; color:#161D26; font-family:Calibri, Arial, sans-serif;">2026&#24180;&#20126;&#39340;&#36956;&#31309;&#20998;&#29378;&#27489;&#31680; #5</h2>
      <p style="font-size:14px; color:#161D26; margin:0;">&#128197; &#27963;&#21205;&#26399;&#38291;&#65306;2026&#24180;5&#26376;8&#26085;&#65288;&#36913;&#20116;&#65289;9:00 &#65374; 2026&#24180;5&#26376;17&#26085;&#65288;&#36913;&#26085;&#65289;23:59</p>
    </div>

    <!-- CTA -->
    <div style="text-align:center; margin:28px 0;">
      <a href="https://amazonexteu.qualtrics.com/jfe/form/SV_e8OXHh3llHfCPno" style="display:inline-block; background-color:#FF6200; color:#FFFFFF; text-decoration:none; font-size:15px; font-weight:700; padding:12px 36px; border-radius:6px;">&#31435;&#21363;&#25552;&#22577; &#8594;</a>
    </div>

    <!-- Deadline -->
    <div style="background-color:#FFF4EC; border-left:4px solid #FF6200; padding:14px 20px; margin:24px 0; font-weight:700; font-size:15px;">
      &#9200; &#22577;&#21517;&#25130;&#27490;&#65306;2026&#24180;4&#26376;19&#26085;&#65288;&#36913;&#26085;&#65289;23:59
    </div>

    <!-- Notes -->
    <div style="background-color:#F5F3EF; border-radius:8px; padding:20px 24px; margin:24px 0;">
      <h3 style="font-size:15px; font-weight:700; margin:0 0 12px; color:#161D26; font-family:Calibri, Arial, sans-serif;">&#128203; &#25552;&#22577;&#27880;&#24847;&#20107;&#38917;</h3>
      <ul style="margin:0; padding-left:20px;">
        <li style="margin-bottom:8px; font-size:14px;">&#36067;&#23478;&#21487;&#36984;&#25799;&#20197;&#19979;&#20219;&#19968;&#26041;&#26696;&#21443;&#21152;&#65306;<br>&#9312; &#20677;&#21443;&#21152; 4&#26376;&#26376;&#20419;<br>&#9313; &#20677;&#21443;&#21152; 5&#26376;&#31309;&#20998;&#29378;&#27489;&#31680;<br>&#9314; &#20841;&#22580;&#27963;&#21205;&#30342;&#21443;&#21152;</li>
        <li style="margin-bottom:8px; font-size:14px;">&#35531;&#28858;&#25152;&#26377; ASIN &#25552;&#20132;<b>&#21516;&#19968;&#20491;&#26178;&#38291;&#27573;</b>&#12290;&#33509;&#20986;&#29694;&#26178;&#38291;&#27573;&#28151;&#21512;&#30340;&#24773;&#27841;&#65292;&#25552;&#22577;&#23559;&#19981;&#20104;&#36890;&#36942;&#12290;</li>
        <li style="margin-bottom:8px; font-size:14px;">&#33509;&#38656;&#35201;&#28858;&#19981;&#21516;&#21830;&#21697;&#35373;&#23450;&#19981;&#21516;&#30340;&#21443;&#33287;&#26178;&#38291;&#65292;&#35531;&#36879;&#36942;<b>&#36067;&#23478;&#24179;&#21488;</b>&#25552;&#20132;&#12290;</li>
      </ul>
    </div>

    <p style="margin:0 0 16px;">&#22914;&#26377;&#20219;&#20309;&#21839;&#38988;&#65292;&#27489;&#36814;&#38568;&#26178;&#33287;&#25105;&#20497;&#32879;&#32363;&#12290;&#26399;&#24453;&#24744;&#30340;&#21443;&#33287;&#65281;</p>
    <p style="margin-top:24px;">
      &#31069; &#29983;&#24847;&#33288;&#38534;&#65292;<br>
      Amazon Global Selling Taiwan &#22296;&#38538;
    </p>

  </div>

  <!-- Footer -->
  <div style="background-color:#161D26; padding:24px 40px; text-align:center; color:#FFFFFF; font-size:12px; line-height:1.8;">
    <p style="margin:0 0 4px;">Amazon Global Selling Taiwan</p>
    <p style="margin:0 0 4px;">23F, No. 100, Songren Rd., Xinyi Dist., Taipei City 110016, Taiwan (R.O.C.)</p>
    <p style="margin:0;">
      <a href="https://gs.amazon.com.tw/" style="color:#FF6200; text-decoration:none;">Website</a> &#65372;
      <a href="https://www.facebook.com/AmazonGlobalSellingTaiwan" style="color:#FF6200; text-decoration:none;">Facebook</a> &#65372;
      <a href="https://page.line.me/920bogra?openQrModal=true" style="color:#FF6200; text-decoration:none;">LINE</a> &#65372;
      <a href="https://www.youtube.com/@amazonglobalsellingtw" style="color:#FF6200; text-decoration:none;">YouTube</a> &#65372;
      <a href="https://gs.amazon.com.tw/learn" style="color:#FF6200; text-decoration:none;">Seller University</a>
    </p>
  </div>

</div>

</body>
</html>'''

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = "【邀請提報】2026年亞馬遜 4月月促 + 5月積分狂歡節"

# Set HTML first, then add inline attachment
mail.HTMLBody = html

# Add logo as inline attachment with Content-ID
att = mail.Attachments.Add(logo_path, 1, 0)  # olByValue=1, Position=0
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "amazonlogo"
)
# Also hide it from attachment list
att.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
    True
)

msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)
print(f"Done: {msg_path}")
