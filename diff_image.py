"""生成差異圖，標記出原始有文字但轉換後沒有的區域（紅色）。"""
import fitz
import struct

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

zoom = 2
mat = fitz.Matrix(zoom, zoom)

# 只檢查第 5 頁作為範例
for page_idx in [0, 4, 9, 16]:  # 第 1, 5, 10, 17 頁
    orig_pix = orig_doc[page_idx].get_pixmap(matrix=mat, alpha=False)
    conv_pix = conv_doc[page_idx].get_pixmap(matrix=mat, alpha=False)

    w, h = orig_pix.width, orig_pix.height

    # 建立差異圖
    diff_pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, w, h), 1)
    diff_pix.set_rect(diff_pix.irect, (255, 255, 255, 255))

    orig_data = orig_pix.samples
    conv_data = conv_pix.samples

    # 逐像素比較
    diff_count = 0
    # 找出「原始有深色像素但轉換後是白色/淺色」的區域
    missing_pixels = 0
    added_pixels = 0

    for y in range(h):
        for x in range(w):
            idx = (y * w + x) * 3
            or_, og, ob = orig_data[idx], orig_data[idx+1], orig_data[idx+2]
            cr, cg, cb = conv_data[idx], conv_data[idx+1], conv_data[idx+2]

            orig_bright = (or_ + og + ob) / 3
            conv_bright = (cr + cg + cb) / 3

            # 原始有深色內容但轉換後變白/淺 = 文字消失
            if orig_bright < 200 and conv_bright > 230:
                missing_pixels += 1

            # 轉換後有深色內容但原始沒有 = 新增文字
            if conv_bright < 200 and orig_bright > 230:
                added_pixels += 1

    total = w * h
    print(f"第 {page_idx+1} 頁: "
          f"消失像素 {missing_pixels} ({missing_pixels/total*100:.3f}%), "
          f"新增像素 {added_pixels} ({added_pixels/total*100:.3f}%)")

orig_doc.close()
conv_doc.close()
