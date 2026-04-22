import win32com.client
import os
import requests
import base64

# Step 1: Download Amazon logo PNG (white version for dark background)
# Use a white PNG version of the Amazon logo
logo_url = "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Amazon_logo.svg/320px-Amazon_logo.svg.png"
resp = requests.get(logo_url)
if resp.status_code == 200:
    logo_b64 = base64.b64encode(resp.content).decode('utf-8')
    print(f"Logo downloaded, size: {len(resp.content)} bytes")
else:
    print(f"Failed to download logo: {resp.status_code}")
    # Fallback: convert SVG to PNG locally
    import cairosvg
    svg_url = "https://upload.wikimedia.org/wikipedia/commons/a/a9/Amazon_logo.svg"
    svg_resp = requests.get(svg_url)
    png_data = cairosvg.svg2png(bytestring=svg_resp.content, output_width=320)
    logo_b64 = base64.b64encode(png_data).decode('utf-8')
    print(f"Logo converted from SVG, size: {len(png_data)} bytes")

# Step 2: Read HTML and replace the img tag with embedded base64 PNG
with open("email_template.html", "r", encoding="utf-8") as f:
    html = f.read()

# Replace the SVG img with inline base64 PNG (white filter applied via inversion)
old_img = '<img src="https://upload.wikimedia.org/wikipedia/commons/a/a9/Amazon_logo.svg" alt="Amazon" style="height:32px; filter: brightness(0) invert(1);">'
new_img = f'<img src="data:image/png;base64,{logo_b64}" alt="Amazon" style="height:32px;">'
html = html.replace(old_img, new_img)

# Step 3: Create .msg via Outlook COM
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = "【邀請提報】2026年亞馬遜 4月月促 + 5月積分狂歡節"
mail.HTMLBody = html

msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)
print(f"MSG file saved to: {msg_path}")
