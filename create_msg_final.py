import win32com.client
import os
from PIL import Image, ImageDraw, ImageFont
import io

# Step 1: Create Amazon-style white logo PNG
img = Image.new('RGBA', (280, 60), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

try:
    font = ImageFont.truetype("arial.ttf", 36)
except:
    font = ImageFont.load_default()

# "amazon" in white
draw.text((10, 5), "amazon", fill=(255, 255, 255, 255), font=font)

# Orange smile arrow
draw.arc([30, 36, 200, 58], start=0, end=180, fill=(255, 98, 0, 255), width=3)
draw.polygon([(195, 40), (207, 48), (195, 54)], fill=(255, 98, 0, 255))

logo_path = os.path.join(os.getcwd(), "amazon_logo_white.png")
img.save(logo_path, format='PNG')
print(f"Logo saved: {logo_path}")

# Step 2: Read HTML and replace img src with CID reference
with open("email_template.html", "r", encoding="utf-8") as f:
    html = f.read()

old_img = '<img src="https://upload.wikimedia.org/wikipedia/commons/a/a9/Amazon_logo.svg" alt="Amazon" style="height:32px; filter: brightness(0) invert(1);">'
new_img = '<img src="cid:amazon_logo" alt="Amazon" style="height:40px;">'
html = html.replace(old_img, new_img)

# Step 3: Create .msg with embedded image via Outlook COM
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = "【邀請提報】2026年亞馬遜 4月月促 + 5月積分狂歡節"
mail.HTMLBody = html

# Attach logo as hidden inline attachment
attachment = mail.Attachments.Add(logo_path, 1)  # 1 = olByValue
# Set Content-ID for CID reference
attachment.PropertyAccessor.SetProperty(
    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
    "amazon_logo"
)

msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)
print(f"MSG file saved to: {msg_path}")
