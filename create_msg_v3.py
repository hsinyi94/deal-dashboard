import win32com.client
import os
import requests
import base64

# Use a reliable PNG source for Amazon logo (white on transparent)
# Try multiple sources
logo_urls = [
    "https://d1.awsstatic.com/logos/aws-logo-lockups/poweredbyaws/PB_AWS_logo_RGB_REV_SQ.8c88ac215fe4e441dc42865dd6962ed4f444a90d.png",
]

logo_b64 = None

# Alternative: create a simple white "amazon" text as PNG using Pillow
from PIL import Image, ImageDraw, ImageFont
import io

# Create a transparent image with white "amazon" text + smile
img = Image.new('RGBA', (280, 60), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Use a clean font
try:
    font = ImageFont.truetype("arial.ttf", 36)
except:
    font = ImageFont.load_default()

# Draw "amazon" in white
draw.text((10, 8), "amazon", fill=(255, 255, 255, 255), font=font)

# Draw the smile arrow underneath (simplified orange curve)
# Simple orange arrow from 'a' to 'z'
draw.arc([30, 38, 200, 58], start=0, end=180, fill=(255, 98, 0, 255), width=3)
# Arrow tip
draw.polygon([(195, 42), (205, 48), (195, 52)], fill=(255, 98, 0, 255))

buf = io.BytesIO()
img.save(buf, format='PNG')
logo_b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
print(f"Logo created, size: {len(buf.getvalue())} bytes")

# Read HTML and replace the img tag
with open("email_template.html", "r", encoding="utf-8") as f:
    html = f.read()

old_img = '<img src="https://upload.wikimedia.org/wikipedia/commons/a/a9/Amazon_logo.svg" alt="Amazon" style="height:32px; filter: brightness(0) invert(1);">'
new_img = f'<img src="data:image/png;base64,{logo_b64}" alt="Amazon" style="height:40px;">'
html = html.replace(old_img, new_img)

# Create .msg
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = "【邀請提報】2026年亞馬遜 4月月促 + 5月積分狂歡節"
mail.HTMLBody = html

msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)
print(f"MSG file saved to: {msg_path}")
