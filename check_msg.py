import win32com.client
import os

outlook = win32com.client.Dispatch('Outlook.Application')
msg = outlook.CreateItemFromTemplate(os.path.join(os.getcwd(), '2026_亞馬遜4月月促_5月積分狂歡節.msg'))
html = msg.HTMLBody

idx = html.find('class="header"')
if idx >= 0:
    print(html[idx:idx+500])
else:
    print('header class not found')

# Also search for img
idx2 = html.find('<img')
if idx2 >= 0:
    print("\n--- img tag found ---")
    print(html[idx2:idx2+300])
else:
    print("\nNo <img> tag found in .msg HTML")
