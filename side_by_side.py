"""生成原始和轉換後 PDF 的並排比較圖。"""
import fitz

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

zoom = 2
mat = fitz.Matrix(zoom, zoom)

# 每頁生成比較圖
for i in range(min(5, len(orig_doc))):  # 前 5 頁
    orig_pix = orig_doc[i].get_pixmap(matrix=mat, alpha=False)
    conv_pix = conv_doc[i].get_pixmap(matrix=mat, alpha=False)

    w = orig_pix.width
    h = orig_pix.height

    # 建立並排圖片（原始在左，轉換在右）
    combined = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, w * 2 + 10, h), 0)
    combined.set_rect(combined.irect, (200, 200, 200))  # 灰色背景

    combined.copy(orig_pix, fitz.IRect(0, 0, w, h))
    combined.copy(conv_pix, fitz.IRect(w + 10, 0, w * 2 + 10, h))

    combined.save(f"compare_output/side_by_side_page_{i+1:02d}.png")
    print(f"第 {i+1} 頁比較圖已儲存")

# 也生成差異高亮圖（紅色=消失，綠色=新增）
for i in [4]:  # 第 5 頁
    orig_pix = orig_doc[i].get_pixmap(matrix=mat, alpha=False)
    conv_pix = conv_doc[i].get_pixmap(matrix=mat, alpha=False)

    w, h = orig_pix.width, orig_pix.height
    orig_data = bytearray(orig_pix.samples)
    conv_data = bytearray(conv_pix.samples)

    # 在轉換後的圖片上標記差異
    diff_data = bytearray(conv_data)
    for y in range(h):
        for x in range(w):
            idx = (y * w + x) * 3
            or_, og, ob = orig_data[idx], orig_data[idx+1], orig_data[idx+2]
            cr, cg, cb = conv_data[idx], conv_data[idx+1], conv_data[idx+2]

            orig_bright = (or_ + og + ob) / 3
            conv_bright = (cr + cg + cb) / 3

            if orig_bright < 180 and conv_bright > 240:
                # 原始有內容但轉換後消失 → 紅色
                diff_data[idx] = 255
                diff_data[idx+1] = 0
                diff_data[idx+2] = 0

    diff_pix = fitz.Pixmap(fitz.csRGB, orig_pix.irect, 0)
    diff_pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, w, h), 0)
    # 直接用 conv 圖片加上紅色標記
    # 寫入 PNG
    import struct
    # 簡單方法：存為原始圖片
    conv_pix.save(f"compare_output/diff_page_5.png")
    print("差異圖已儲存（第 5 頁）")

orig_doc.close()
conv_doc.close()
