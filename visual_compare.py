"""
將原始和轉換後 PDF 的每頁渲染為圖片，逐像素比對差異。
這是最準確的視覺比對方式。
"""
import fitz
import os

orig_doc = fitz.open("MDE4/2026_MDE4_3P_CN.pdf")
conv_doc = fitz.open("MDE4/2026_MDE4_3P_CN_繁體.pdf")

os.makedirs("compare_output", exist_ok=True)

zoom = 2
mat = fitz.Matrix(zoom, zoom)

for i in range(len(orig_doc)):
    orig_pix = orig_doc[i].get_pixmap(matrix=mat, alpha=False)
    conv_pix = conv_doc[i].get_pixmap(matrix=mat, alpha=False)

    orig_pix.save(f"compare_output/page_{i+1:02d}_orig.png")
    conv_pix.save(f"compare_output/page_{i+1:02d}_conv.png")

    # 比較像素差異
    if orig_pix.width == conv_pix.width and orig_pix.height == conv_pix.height:
        orig_samples = orig_pix.samples
        conv_samples = conv_pix.samples
        total = len(orig_samples)
        diff = sum(1 for a, b in zip(orig_samples, conv_samples) if a != b)
        pct = diff / total * 100
        print(f"第 {i+1:2d} 頁: {pct:.2f}% 像素差異")
    else:
        print(f"第 {i+1:2d} 頁: 尺寸不同 orig={orig_pix.width}x{orig_pix.height} "
              f"conv={conv_pix.width}x{conv_pix.height}")

orig_doc.close()
conv_doc.close()
print("\n圖片已儲存到 compare_output/ 資料夾")
