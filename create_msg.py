import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 = olMailItem

mail.Subject = "【邀請提報】2026年亞馬遜 4月月促 + 5月積分狂歡節"

# Read HTML content
html_path = os.path.join(os.getcwd(), "email_template.html")
with open(html_path, "r", encoding="utf-8") as f:
    html_body = f.read()

mail.HTMLBody = html_body

# Save as .msg
msg_path = os.path.join(os.getcwd(), "2026_亞馬遜4月月促_5月積分狂歡節.msg")
mail.SaveAs(msg_path, 3)  # 3 = olMSG format

print(f"MSG file saved to: {msg_path}")
